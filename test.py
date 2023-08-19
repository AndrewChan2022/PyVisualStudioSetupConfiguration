from PyVisualStudioSetupConfiguration import GetAllVSInstanceInfo
import PyVisualStudioSetupConfiguration as vssetup
import time

if __name__ == '__main__':


    print("********* GetAllVSInstanceInfo: **************")
    print(GetAllVSInstanceInfo())

    print("********* ComGetAllVSInstanceInfo: **************")
    print(vssetup.ComGetAllVSInstanceInfo())


    print("********* VSWhereGetAllVSInstanceInfo: **************")
    print(vssetup.VSWhereGetAllVSInstanceInfo())

    print("********* RegeditGetAllVSInstanceInfo: **************")
    print(vssetup.RegeditGetAllVSInstanceInfo())

    print("********* EnvGetAllVSInstanceInfo: **************")
    print(vssetup.EnvGetAllVSInstanceInfo())

    print("********* GetCMakeDefaultVSInstanceInfo: **************")
    print(vssetup.GetCMakeDefaultVSInstanceInfo())

    # benchmark
    print("*********performance benchmark:**************")
    tic = time.time()
    for i in range(11):
        vsinstances = GetAllVSInstanceInfo(needChipInfo=True, skipEWDK=True, ignoreCache=True)
        #if i % 10 == 0:
        #    print("i:", i, vsinstances)
    toc = time.time()
    ms = (toc - tic) * 1000
    print("GetAllVSInstanceInfo 10 times: %0.2fms\n"%(ms))
    