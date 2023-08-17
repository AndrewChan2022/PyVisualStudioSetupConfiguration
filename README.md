# PyVisualStudioSetupConfiguration


# overview

This is pure python single file to search Visual Studio installation information.

The result is list of Visual Studio instances with version, path, and chip information.

Windows COM implement using python standard library ctypes, inspired by [comtypes](https://github.com/enthought/comtypes) and [pywin32](https://github.com/mhammond/pywin32).

Visual Studio search algorithm inspired by [CMake](https://github.com/Kitware/CMake).


# future work

1. use meta class to generate com function

2. use c_void_p class type to avoid explict to specify this pointer to COM interface


# usage

## code example

```python

from PyVisualStudioSetupConfiguration import GetAllVSInstanceInfo

# get list of VisualStudio Instance Info
vsInstances = GetAllVSInstanceInfo()

# print the list
print(vsInstances)

# get version of first instance
if len(vsInstances):
    version = vsInstances[0].getVersion()
    major   = vsInstances[0].getVerionMajor()
    chip    =  vsInstances[0].chip or ""

```

## example output

```text

[VSInstanceInfo :
    VSInstallLocation   : d:\Program Files\Microsoft Visual Studio\2022\Community
    Version             : 17.4.33213.308
    versionMajor        : 17
    VCToolsetVersion    : 14.34.31933
    bWin10SDK           : True
    bWin81SDK           : None
    chip                : x64
]
```

## VSInstanceInfo define

```python

class VSInstanceInfo:
    def __init__(self) -> None:

        # exampel:
        # VSInstallLocation   : d:\Program Files\Microsoft Visual Studio\2022\Community
        # Version             : 17.4.33213.308
        # versionMajor        : 17
        # VCToolsetVersion    : 14.34.31933
        # bWin10SDK           : True
        # bWin81SDK           : False
        # chip                : x64

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

```

# implement

## algorithm

algorithm to find VisualStudio Instance

```text

GetAllVSInstanceInfo:

    if driver develop machine:
        Get Instance for EWDK(Enterprise Windows Driver Kit)

    if found nothing:
        try to find new version from version 15 2017 by windows COM, with detail setup info
    
    if found nothing:
        try to find new version from verion 15 2017 by VSWhere.exe, with limit setup info

    if found nothing:
        try to find old version from register table
```

when get vsinstance, we can know the cmake generator based on the table:

cmake generators for VS major version with cmake --help:
```text

  Visual Studio 17 2022        = Generates Visual Studio 2022 project files.
                                 Use -A option to specify architecture.
  Visual Studio 16 2019        = Generates Visual Studio 2019 project files.
                                 Use -A option to specify architecture.
  Visual Studio 15 2017 [arch] = Generates Visual Studio 2017 project files.
                                 Optional [arch] can be "Win64" or "ARM".
  Visual Studio 14 2015 [arch] = Generates Visual Studio 2015 project files.
                                 Optional [arch] can be "Win64" or "ARM".
  Visual Studio 12 2013 [arch] = Generates Visual Studio 2013 project files.
                                 Optional [arch] can be "Win64" or "ARM".
  Visual Studio 11 2012 [arch] = Deprecated.  Generates Visual Studio 2012
                                 project files.  Optional [arch] can be
                                 "Win64" or "ARM".
  Visual Studio 9 2008 [arch]  = Deprecated.  Generates Visual Studio 2008

```


COM implement:

1. use ctypes to load COM lib,

2. then create COM object and get Interface.

3. 




# Acknowledgement

Visual Studio Search algorithm reference from:

* [CMake](https://github.com/Kitware/CMake)
    * [cmVSSetupHelper.cxx](https://github.com/Kitware/CMake/blob/7b49424489b7c1a6ba5487e6dfcf227be74e6720/Source/cmVSSetupHelper.cxx)
* [vs-setup-samples](https://github.com/microsoft/vs-setup-samples)

Official COM Interface and Document of Microsoft.VisualStudio.Setup.Configuration

* [Setup.Configuration.h](https://www.nuget.org/packages/Microsoft.VisualStudio.Setup.Configuration.Native/)
* [Microsoft.VisualStudio.Setup.Configuration](https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualstudio.setup.configuration?view=visualstudiosdk-2022)


Windows COM implement reference from:

* [pywin32](https://github.com/mhammond/pywin32)
* [comtypes](https://github.com/enthought/comtypes)
