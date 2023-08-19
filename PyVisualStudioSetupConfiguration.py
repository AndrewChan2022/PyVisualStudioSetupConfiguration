#
# Copyright 2023 Pixar
#
# Licensed under the Apache License, Version 2.0 (the "Apache License")
# with the following modification; you may not use this file except in
# compliance with the Apache License and the following modification to it:
# Section 6. Trademarks. is deleted and replaced with:
#
# 6. Trademarks. This License does not grant permission to use the trade
#    names, trademarks, service marks, or product names of the Licensor
#    and its affiliates, except as required to comply with Section 4(c) of
#    the License and to reproduce the content of the NOTICE file.
#
# You may obtain a copy of the Apache License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the Apache License with the above modification is
# distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied. See the Apache License for the specific
# language governing permissions and limitations under the Apache License.
#

# Utilities for finding Visual Studio installation information
#
# this file should keep as single file independent of others

import os, subprocess, winreg, atexit, re
import ctypes
import ctypes.wintypes
import json
from enum import IntEnum

__all__ = [
    "GetAllVSInstanceInfo",     # get list of VSInstanceInfo
    "VSInstanceInfo",           # result item of GetAllVSInstanceInfo
]

class SAFEARRAYBOUND(ctypes.Structure):
    _fields_ = [("cElements" , ctypes.c_ulong),
                ("lLbound" , ctypes.c_long)]

class SAFEARRAY(ctypes.Structure):
    _fields_ = [("cDims", ctypes.c_ushort),
                ("fFeatures", ctypes.c_ushort),
                ("cbElements", ctypes.c_ulong),
                ("cLocks", ctypes.c_ulong),
                ("pvData", ctypes.c_void_p),
                ("rgsabound", SAFEARRAYBOUND * 1)]

CoCreateInstance = None
CoUninitialize = None
SafeArrayDestroy = None
def InitCom():
    global CoCreateInstance
    global CoUninitialize
    global SafeArrayDestroy
    try:
        ole32 = ctypes.OleDLL('Ole32.dll')
        CoInitialize = ole32.CoInitialize
        CoUninitialize = ole32.CoUninitialize
        CoCreateInstance = ole32.CoCreateInstance

        oleAut32Dll = ctypes.WinDLL("OleAut32.dll")
        SafeArrayDestroy = oleAut32Dll.SafeArrayDestroy
        SafeArrayDestroy.argtypes = (ctypes.POINTER(SAFEARRAY),)
        SafeArrayDestroy.restype = ctypes.c_long

        _ = CoInitialize(None)
    except Exception as e:
        print("COM init fail:", e)

# InitCOM be called for thread who imports this file first time
InitCom()

# UninitCom is registered to be called when Python is shut down
@atexit.register
def UninitCom(func = CoUninitialize):
    if func:
        try: func()
        except WindowsError: pass

def CreateComObject(clsid, interface = None):
    if not interface:
        interface = IUnknown
    iid = interface._iid
    obj = ctypes.c_void_p(None)
    clsctx = 1
    try: rc = CoCreateInstance(ctypes.byref(clsid), 0, clsctx, ctypes.byref(iid), ctypes.byref(obj))
    except: return None

    if rc != 0:
        return None
    return interface(obj)

def COMMETHOD(idx, restype, name, *argtypes):
    """pack args to tuple argtypes, then unpack in GenerateComMethod"""
    return (idx, restype, name, argtypes)
    
def GenerateComMethod(cls, instance, functions):
    """generate COM function as class instance method"""
    if cls != IUnknown:
        GenerateComMethod(cls.__bases__[0], instance, functions)
        startIndex = cls.__bases__[0].GetComMethodCount()
    else:
        startIndex = 0

    # generate method
    for method in cls._methods:
        idx, restype, name, rawargtypes = method
        argtypes = list(map(lambda arg: arg[1], rawargtypes))
        interfaceThisType = ctypes.c_void_p
        comfunc = ctypes.WINFUNCTYPE(restype, interfaceThisType, *argtypes)(functions[startIndex + idx])
        # set attribute of instance, not cls, otherwise be overrided by others
        setattr(instance, "_%s__com_%s" % (cls.__name__, name), comfunc)

class GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8)]
    
    def __init__(self, name=None):
        if name is not None:
            s = name.strip("{}")
            m = re.search('([0-9A-Fa-f]{8})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{12})', s)
            if m:
                s = s.replace("-", "")
                self.Data1 = int(s[0:8], 16)
                self.Data2 = int(s[8:12], 16)
                self.Data3 = int(s[12:16], 16)
                for i in range(8):
                    self.Data4[i] = int(s[16 + i * 2 : 16 + i * 2 + 2], 16)

    def __repr__(self):
        return 'GUID("%s")' % str(self)

    def __str__(self):
        return "{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}"%(
            self.Data1, self.Data2, self.Data3, 
            self.Data4[0], self.Data4[1], self.Data4[2], self.Data4[3],
            self.Data4[5], self.Data4[5], self.Data4[6], self.Data4[7])

class IUnknown(object):
    """base class for com interface
    TODO: should use meta class to simplify
    """

    # private config data of com interface
    _iid = GUID("{00000000-0000-0000-C000-000000000046}")
    _methods = [
        COMMETHOD(0, ctypes.HRESULT, "QueryInterface",
                  (['in'], ctypes.POINTER(GUID), "riid"), 
                  (['out'], ctypes.POINTER(ctypes.c_void_p), "ppvObject")),
        COMMETHOD(1, ctypes.HRESULT, "AddRef"),
        COMMETHOD(2, ctypes.HRESULT, "Release"),
        ]
    _method_count = len(_methods)

    # utils to generate com function
    @classmethod
    def GetComMethodCount(cls):
        if cls == IUnknown:
            return cls._method_count
        else:
            return cls._method_count + cls.__bases__[0].GetComMethodCount()

    def __init__(self, comInterface = None, owner = True) -> None:
        """use RAII to manage COM reference, register Release at destructor
        comInterface: nullptr or 'this' pointer of COM interface
        owner: False if comInterface already has owner, so AddRef also to be its onwer
        """
        atexit.register(self._AutoCleanComReference_)
        if comInterface is None:
            self._IThis = ctypes.c_void_p()
            return
        self.SetComInterface(comInterface)
        if comInterface and not owner:
            # if not owner of the object, AddRef to be owner
            self.AddRef()

    def _AutoCleanComReference_(self):
        if self._IThis:
            # unnecessary, will be released when thread over
            try: self.Release()
            except Exception as e: pass # print(e)

    def SetComInterface(self, comInterface):
        self._IThis = comInterface
        if not comInterface:
            return
        
        # interface point to object, first element is vtable
        VTable = ctypes.cast(comInterface, ctypes.POINTER(ctypes.c_void_p))
        wk = ctypes.c_void_p(VTable[0])
        functions = ctypes.cast(wk, ctypes.POINTER(ctypes.c_void_p))
        GenerateComMethod(type(self), self, functions)

    # com interface methods
    def QueryInterface(self, interface):
        iid = interface._iid
        p = ctypes.c_void_p()
        rc = self.__com_QueryInterface(self._IThis, ctypes.byref(iid), ctypes.byref(p))
        if rc != 0:
            return None
        return interface(p)
    
    def AddRef(self):
        return self.__com_AddRef(self._IThis)

    def Release(self):
        return self.__com_Release(self._IThis)



class ISetupConfiguration(IUnknown):
    """official interface: https://github.com/microsoft/vs-setup-samples
    official example: https://github.com/Microsoft/vswhere
    """
    _iid = GUID("{42843719-DB4C-46C2-8E7C-64F1816EFD5B}")
    _methods = [
        COMMETHOD(0, ctypes.HRESULT, "EnumInstances",
                        (['out'], ctypes.POINTER(ctypes.c_void_p), "ppEnumInstances")),
        COMMETHOD(1, ctypes.HRESULT, "GetInstanceForCurrentProcess",
                        (['out'], ctypes.POINTER(ctypes.c_void_p), "ppInstance")),
        COMMETHOD(2, ctypes.HRESULT, "GetInstanceForPath",
                           (['in'], ctypes.wintypes.LPCWSTR, "wzPath"),
                           (['out'], ctypes.POINTER(ctypes.c_void_p), "ppInstance")),
        ]
    _method_count = len(_methods)

    def EnumInstances(self, interface):
        p = ctypes.c_void_p()
        rc = self.__com_EnumInstances(self._IThis, ctypes.byref(p))
        if rc != 0:
            return None
        return interface(p)

class ISetupConfiguration2(ISetupConfiguration):
    _iid = GUID("{26AAB78C-4A60-49D6-AF3B-3C35BC93365D}")
    _methods = [
        COMMETHOD(0, ctypes.HRESULT, "EnumAllInstances",
                  (['out'], ctypes.POINTER(ctypes.c_void_p), "ppEnumInstances")),
        ]
    _method_count = len(_methods)

    def EnumAllInstances(self, interface):
        p = ctypes.c_void_p()
        rc = self.__com_EnumAllInstances(self._IThis, ctypes.byref(p))
        if rc != 0:
            return None
        return interface(p)

class IEnumSetupInstances(IUnknown):
    _iid = GUID("{6380BCFF-41D3-4B2E-8B2E-BF8A6810C848}")
    _methods = [
        COMMETHOD(0, ctypes.HRESULT, "Next",
                (['in'], ctypes.c_ulong, "celt"),
                (['out'], ctypes.POINTER(ctypes.c_void_p), "rgelt"),
                (['out'], ctypes.POINTER(ctypes.c_ulong), "pceltFetched")),
        COMMETHOD(1, ctypes.HRESULT, "Skip",
                (['in'], ctypes.c_ulong, "celt")),
        COMMETHOD(2, ctypes.HRESULT, "Reset"),
        COMMETHOD(3, ctypes.HRESULT, "Clone",
                (['out'], ctypes.POINTER(ctypes.c_void_p), "ppenum")),
    ]
    _method_count = len(_methods)

    def Next(self, celt, rgelt, rgeltResultList, pceltFetched = None):
        """python cannot pass by reference, so put to rgeltResultList"""
        if isinstance(rgeltResultList, list):
            rgeltResultList.clear()

        p = ctypes.c_void_p()
        if rgelt._IThis is not None:
            # copy pointer, create different object pointing to same address
            p = ctypes.c_void_p(rgelt._IThis.value)
        rc = self.__com_Next(self._IThis, celt, ctypes.byref(p), pceltFetched)

        if rc != 0 or not p: # nullptr is not None, but condition get False
            return None
        
        rgeltResult = ISetupInstance(p)
        if isinstance(rgeltResultList, list):
            rgeltResultList.append(rgeltResult)
        return rgeltResult

class ISetupInstance(IUnknown):
    _iid = GUID("{B41463C3-8866-43B5-BC33-2B0676F7F42E}")
    _methods = [
        COMMETHOD(0, ctypes.HRESULT, "GetInstanceId",
                            (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrInstanceId"),
                            ),
        COMMETHOD(1, ctypes.HRESULT, "GetInstallDate",
                           (['out'], ctypes.wintypes.LPFILETIME, "pInstallDate"),
                           ),
        COMMETHOD(2, ctypes.HRESULT, "GetInstallationName",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrInstallationName"),
                           ),
        COMMETHOD(3, ctypes.HRESULT, "GetInstallationPath",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrInstallationPath"),
                           ),
        COMMETHOD(4, ctypes.HRESULT, "GetInstallationVersion",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrInstallationVersion"),
                           ),
        COMMETHOD(5, ctypes.HRESULT, "GetDisplayName",
                           (['in'], ctypes.wintypes.LCID, "lcid"),
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrDisplayName"),
                           ),
        COMMETHOD(6, ctypes.HRESULT, "GetDescription",
                           (['in'], ctypes.wintypes.LCID, "lcid"),
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrDescription"),
                           ),
        COMMETHOD(7, ctypes.HRESULT, "ResolvePath",
                           (['in'], ctypes.wintypes.LPCWSTR, "pwszRelativePath"),
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrAbsolutePath"),
                           ),
    ]
    _method_count = len(_methods)

    def GetInstallationPath(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetInstallationPath(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value
    
    def GetInstallationVersion(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetInstallationVersion(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value


class ISetupInstance2(ISetupInstance):

    _iid = GUID("{89143C9A-05AF-49B0-B717-72E218A2185C}")
    _methods = [
        COMMETHOD(0, ctypes.HRESULT, "GetState",
                           (['out'], ctypes.POINTER(ctypes.c_int), "pState"),
                           ),
        COMMETHOD(1, ctypes.HRESULT, "GetPackages",
                           (['out'], ctypes.POINTER(ctypes.POINTER(SAFEARRAY)), "ppsaPackages"),
                           ),
        COMMETHOD(2, ctypes.HRESULT, "GetProduct",
                           (['out'], ctypes.POINTER(ctypes.c_void_p), "ppPackage"),
                           ),
        COMMETHOD(3, ctypes.HRESULT, "GetProductPath",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrProductPath"),
                           ),
    ]
    _method_count = 4

    def GetState(self):
        state = ctypes.c_int()
        rc = self.__com_GetState(self._IThis, ctypes.byref(state))
        if rc != 0:
            return None
        return state.value
    
    def GetProduct(self, interface):
        p = ctypes.c_void_p()
        rc = self.__com_GetProduct(self._IThis, ctypes.byref(p))
        if rc != 0 or not p:
            return None
        return interface(p)
    
    def GetPackages(self, dataType = SAFEARRAY):
        p = ctypes.POINTER(dataType)()
        rc = self.__com_GetPackages(self._IThis, ctypes.byref(p))
        if rc != 0 or not p:
            return None
        return p.contents

class ISetupPackageReference(IUnknown):
    _iid = GUID("{da8d8a16-b2b6-4487-a2f1-594ccccd6bf5}")
    _methods = [
        COMMETHOD(0, ctypes.HRESULT, "GetId",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrId"),
                           ),
        COMMETHOD(1, ctypes.HRESULT, "GetVersion",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrVersion"),
                           ),
        COMMETHOD(2, ctypes.HRESULT, "GetChip",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrChip"),
                           ),
        COMMETHOD(3, ctypes.HRESULT, "GetLanguage",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrLanguage"),
                           ),
        COMMETHOD(4, ctypes.HRESULT, "GetBranch",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrBranch"),
                           ),
        COMMETHOD(5, ctypes.HRESULT, "GetType",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrType"),
                           ),
        COMMETHOD(5, ctypes.HRESULT, "GetUniqueId",
                           (['out'], ctypes.POINTER(ctypes.c_wchar_p), "pbstrUniqueId"),
                           ),
    ]
    _method_count = len(_methods)

    def GetId(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetId(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value
    
    def GetVersion(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetVersion(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value
    
    def GetChip(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetChip(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value
    
    def GetLanguage(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetLanguage(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value
    
    def GetBranch(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetBranch(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value
    
    def GetType(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetType(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value
    
    def GetUniqueId(self):
        ps = ctypes.c_wchar_p()
        rc = self.__com_GetUniqueId(self._IThis, ctypes.byref(ps))
        if rc != 0:
            return None
        return ps.value

class EInstanceState(IntEnum):
    # The instance state has not been determined.
    eNone = 0,
    
    # The instance installation path exists.
    eLocal = 1,

    # A product is registered to the instance.
    eRegistered = 2,

    # No reboot is required for the instance.
    eNoRebootRequired = 4,

    # The instance represents a complete install.
    eComplete = 0xffffffff  #MAXUINT

def ReadWinreg(rootkey, path, valueName):
    """wrap of winreg to return None instead of exception
    key: winreg.HKEY_LOCAL_MACHINE
    path: r"SOFTWARE\Microsoft\Windows\CurrentVersion"
    valueName: name of one value
    """
    try: 
        # open folder
        key = winreg.OpenKeyEx(rootkey, path)

        # open item
        item = winreg.QueryValueEx(key, valueName)

        # close folder
        winreg.CloseKey(key)

        return item
    except:
        return None


def GetEWDKAllVSInstanceInfo(skipEWDK):
    "Enterprise Windows Driver Kit"
    if skipEWDK:
        return []

    envEnterpriseWDK = os.getenv("EnterpriseWDK") or ""
    envDisableRegistryUse = os.getenv("DisableRegistryUse") or ""
    if envEnterpriseWDK.lower() == "true" and envDisableRegistryUse.lower() == "true":
        envWindowsSdkDir81 = os.getenv("WindowsSdkDir_81") or ""
        envVSVersion = os.getenv("VisualStudioVersion")
        envVsInstallDir = os.getenv("VSINSTALLDIR")
        envVCToolsVersion = os.getenv("VCToolsVersion")
        envTargetArch = os.getenv("VSCMD_ARG_TGT_ARCH")
        
        vsInstanceInfo = VSInstanceInfo()
        vsInstanceInfo.bWin10SDK = True
        vsInstanceInfo.bWin81SDK = envWindowsSdkDir81 != ""
        vsInstanceInfo.Version = envVSVersion
        vsInstanceInfo.VSInstallLocation = envVsInstallDir
        vsInstanceInfo.VCToolsetVersion = envVCToolsVersion
        vsInstanceInfo.chip = envTargetArch

        if envVSVersion:
            return [vsInstanceInfo]
    
    return []

def _ComGetOneVSInstanceInfo(setupInstance2, needChipInfo):
    if not setupInstance2:
        return None
    vsInstanceInfo = VSInstanceInfo()

    # get installation state
    state = setupInstance2.GetState()
    if not state:
        return None
    
    # get installation version
    version = setupInstance2.GetInstallationVersion()
    if not version:
        return None
    vsInstanceInfo.Version = version  # no need Wstr to Str
    
    # get installation path
    # Reboot may have been required before the installation path was created
    if state & int(EInstanceState.eLocal) == int(EInstanceState.eLocal):
        installationPath = setupInstance2.GetInstallationPath()
        if not installationPath:
            return None
        vsInstanceInfo.VSInstallLocation = installationPath # no need Wstr to Str
    
    # Check if a compiler is installed with this instance.
    vcRoot = vsInstanceInfo.VSInstallLocation
    if not os.path.isdir(vcRoot):
        return None
    # Microsoft.VCToolsVersion.default.txt maybe override by later installed version 
    vcToolsVersionFilePath = os.path.join(vcRoot, "VC/Auxiliary/Build/Microsoft.VCToolsVersion.default.txt")
    if os.path.isfile(vcToolsVersionFilePath):
        with open(vcToolsVersionFilePath) as f:
            vcToolsVersion = f.readline().strip()
            vcToolsDir = os.path.join(vcRoot, "VC/Tools/MSVC/", vcToolsVersion)
            if os.path.isdir(vcToolsDir):
                vsInstanceInfo.VCToolsetVersion = vcToolsVersion

    # enumerate all packages to find win10SDKInstalled win81SDkInstalled and chip info
    # Reboot may have been required before the product package was registered
    if needChipInfo and (state & int(EInstanceState.eRegistered) == int(EInstanceState.eRegistered)):
        # check at least one package exist
        package0 = setupInstance2.GetProduct(ISetupPackageReference) # error
        if not package0:
            return None

        # get packages array
        packages = setupInstance2.GetPackages()
        if not packages:
            return None

        # enumrate each package
        lower = packages.rgsabound[0].lLbound
        count = packages.rgsabound[0].cElements
        pvData = ctypes.cast(packages.pvData, ctypes.POINTER(ctypes.c_void_p))
        for i in range(lower, lower + count):
            # get package ISetupPackageReference by its IUnknown interface
            package = IUnknown(pvData[i], False) # pvData[i] already owned by packages
            package = package.QueryInterface(ISetupPackageReference) if package else None
            if not package:
                continue
            
            packageId = package.GetId()
            packageType = package.GetType()
            chip = package.GetChip()
            win10SDKComponent = "Microsoft.VisualStudio.Component.Windows10SDK"
            win81SDKComponent = "Microsoft.VisualStudio.Component.Windows81SDK"
            vsProductPrefix = "Microsoft.VisualStudio.Product."  # Community/Professional/Enterprise
            if win10SDKComponent in packageId and packageType == "Component":
                vsInstanceInfo.bWin10SDK = True                    
            if win81SDKComponent == packageId and packageType == "Component":
                vsInstanceInfo.bWin81SDK = True                
            if vsProductPrefix in packageId and packageType == "Product":
                vsInstanceInfo.chip = chip
        
        # maybe unnecessary
        SafeArrayDestroy(ctypes.pointer(packages))
    return vsInstanceInfo

def ComGetAllVSInstanceInfo(needChipInfo = True):
    vsInstances = []

    # remove try, should no except
    try:
        ### get instance enumerator
        clsid = GUID("{177F0C4A-1CD3-4DE7-A32C-71DBBB9FA36D}")
        unknown = CreateComObject(clsid, IUnknown)
        if not unknown:
            return vsInstances
        setupConfig = unknown.QueryInterface(ISetupConfiguration)
        setupConfig2 = setupConfig.QueryInterface(ISetupConfiguration2)
        enumInstances = setupConfig2.EnumAllInstances(IEnumSetupInstances)

        ### search all Visual Studio Version by instance enumerator
        setupInstance = ISetupInstance()
        resultSetupInstances = []
        while enumInstances.Next(1, setupInstance, resultSetupInstances):
            setupInstance = resultSetupInstances[0]
            setupInstance2 = setupInstance.QueryInterface(ISetupInstance2)
            if setupInstance2:
                vsInstanceInfo = _ComGetOneVSInstanceInfo(setupInstance2, needChipInfo)
                vsInstances.append(vsInstanceInfo) if vsInstanceInfo is not None else None
    except Exception as e:
        print("ComGetAllVSInstanceInfo fail:", e)
    
    # sort
    vsInstances.sort(key = lambda x: x.getVersion(), reverse=True)
    return vsInstances

def VSWhereGetAllVSInstanceInfo():
    vsInstances = []

    for programFile in ["ProgramFiles(x86)", "ProgramFiles"]:
        programFilePath = os.getenv(programFile)
        if programFilePath:
            vsWherePath = programFilePath + "/Microsoft Visual Studio/Installer/vswhere.exe"
            
            if os.path.isfile(vsWherePath):

                try: result = subprocess.check_output(f'"{vsWherePath}" -format json')
                except : return vsInstances
                try: vsWhereInfo = json.loads(result)
                except: vsWhereInfo = []

                for item in vsWhereInfo:
                    vsInstanceInfo = VSInstanceInfo()
                    vsInstanceInfo.VSInstallLocation = item.get('installationPath')
                    vsInstanceInfo.Version = item.get('installationVersion') or ""
                    vsInstanceInfo.VCToolsetVersion = None
                    vsInstanceInfo.bWin10SDK = None
                    vsInstanceInfo.bWin81SDK = None
                    vsInstanceInfo.chip = None

                    vsInstances.append(vsInstanceInfo)
    
    #sort
    vsInstances.sort(key = lambda x: x.getVersion(), reverse=True)
    return vsInstances

def RegeditGetAllVSInstanceInfo():
    vsInstances = []
    vsGenerators = [("14.0", "Visual Studio 14 2015"),
                    ("12.0", "Visual Studio 12 2013"),
                    ("11.0", "Visual Studio 11 2012"),
                    ("9.0", "Visual Studio 9 2008")]
    vsVariants   = ['VisualStudio\\', 'VCExpress\\',  'WDExpress\\']
    vsEntries    = [("", "InstallDir"), ("\\Setup\\VC", "ProductDir")]

    for vsGenerator in vsGenerators:
        for vsVariant in vsVariants:
            for vsEntry in vsEntries:
                # HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VisualStudio\14.0\Setup\VC;ProductDir
                key = "SOFTWARE\\Microsoft\\" + vsVariant + vsGenerator[0] + vsEntry[0]
                valueName = vsEntry[1]
                
                dir = ReadWinreg(winreg.HKEY_LOCAL_MACHINE, key, valueName)
                if dir:
                    vsInstanceInfo = VSInstanceInfo()
                    vsInstanceInfo.Version = vsGenerator[0]
                    vsInstanceInfo.VSInstallLocation = dir
                    vsInstances.append(vsInstanceInfo)
    
    # remove duplicate and sort
    versions = set()
    vsInstances = [(versions.add(item.getVersion()), item)[1] for item in vsInstances if item.getVersion() not in versions]
    vsInstances.sort(key = lambda x: x.getVersion(), reverse=True)
    return vsInstances

def EnvGetAllVSInstanceInfo():
    vsInstances = []
    oldVSVersions = [
        ("Visual Studio 14 2015", "14"),
        ("Visual Studio 12 2013", "12"),
        ("Visual Studio 11 2012", "11"),
        ("Visual Studio 9 2008", "9")]

    for item in oldVSVersions:
        version = item[1]
        
        # VS120COMNTOOLS
        envVSCommonToolsDir = os.getenv("VS" + version + "0COMNTOOLS") or ""

        if envVSCommonToolsDir:
            if os.path.isdir(envVSCommonToolsDir):
                vsInstanceInfo = VSInstanceInfo()
                vsInstanceInfo.Version = version
                vsInstanceInfo.VSInstallLocation = os.path.abspath(os.path.join(envVSCommonToolsDir, '../..'))
                vsInstances.append(vsInstanceInfo)
    
    # sort
    vsInstances.sort(key = lambda x: x.getVersion(), reverse=True)
    return vsInstances
    

def GetCMakeDefaultVSInstanceInfo():
    """return list of one VSInstanceInfo"""
    output = subprocess.getoutput("cmake --help")
    if output:
        m = re.search('\* *(Visual Studio (\d+) (\d+))', output)
        if m:
            vsMajor = m.group(2)
            vsInstance = VSInstanceInfo()
            vsInstance.Version = vsMajor
            return [vsInstance]

    return []

class VSInstanceInfo:
    def __init__(self) -> None:
        """exampel:
        # VSInstallLocation   : d:\Program Files\Microsoft Visual Studio\2022\Community
        # Version             : 17.4.33213.308
        # versionMajor        : 17
        # VCToolsetVersion    : 14.34.31933
        # bWin10SDK           : True
        # bWin81SDK           : False
        # chip                : x64"""
        self.VSInstallLocation = None   # option, string
        self.Version = None             # mandatory, string
        self.VCToolsetVersion = None    # option
        self.bWin10SDK = None           # option
        self.bWin81SDK = None           # option
        self.chip = None                # option or string

    def getVersion(self):
        return self.Version or ""
    
    # string of version major
    def getVerionMajor(self):
        if self.Version:
            numbers = self.Version.split('.')
            return numbers[0] if len(numbers) else None
        return None

    def __repr__(self):
        s = 'VSInstanceInfo :\n'
        s += f'    VSInstallLocation   : {self.VSInstallLocation}\n'
        s += f'    Version             : {self.Version}\n'
        s += f'    versionMajor        : {self.getVerionMajor()}\n'
        s += f'    VCToolsetVersion    : {self.VCToolsetVersion}\n'
        s += f'    bWin10SDK           : {self.bWin10SDK}\n'
        s += f'    bWin81SDK           : {self.bWin81SDK}\n'
        s += f'    chip                : {self.chip}\n'
        return s


vsInstancesCache = None

def GetAllVSInstanceInfo(needChipInfo = True, skipEWDK = True, ignoreCache = False):
    """return list of VSInstanceInfo object
    needChipInfo: will enumerate package to get chip so a little slow
    skipEWDK: skip EWDK(Enterprise Windows Driver Kit)
    ignoreCache: not use cache stored by last call, only true for performance benchmark"""

    # get from cache if already searched
    global vsInstancesCache
    if vsInstancesCache and not ignoreCache:
        return vsInstancesCache
    
    # maybe need skip EWDK(Enterprise Windows Driver Kit)
    vsInstancesCache = GetEWDKAllVSInstanceInfo(skipEWDK)
    if len(vsInstancesCache) != 0:
        return vsInstancesCache

    # find new version by windows COM with full info
    # if fail, find with vswhere with only path and version info
    # find old version by register table, with only path version info
    # if fail, get from env
    vsInstancesCache = ComGetAllVSInstanceInfo(needChipInfo) or VSWhereGetAllVSInstanceInfo()
    vsInstancesCache += RegeditGetAllVSInstanceInfo() or EnvGetAllVSInstanceInfo()
    if len(vsInstancesCache) != 0:
        return vsInstancesCache
    
    # try to get CMake default generator
    vsInstancesCache = GetCMakeDefaultVSInstanceInfo()
    if len(vsInstancesCache) != 0:
        return vsInstancesCache
    
    return []
