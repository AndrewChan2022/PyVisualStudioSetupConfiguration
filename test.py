from PyVisualStudioSetupConfiguration import GetAllVSInstanceInfo
import PyVisualStudioSetupConfiguration as vssetup


if __name__ == '__main__':


    print("********* GetAllVSInstanceInfo: **************")
    print(GetAllVSInstanceInfo())

    print("********* ComGetAllVSInstanceInfo: **************")
    print(vssetup.ComGetAllVSInstanceInfo())


    print("********* VSWhereGetAllVSInstanceInfo: **************")
    print(vssetup.VSWhereGetAllVSInstanceInfo())

    print("********* RegeditGetAllVSInstanceInfo: **************")
    print(vssetup.RegeditGetAllVSInstanceInfo())


    # benchmark
    print("*********performance benchmark:**************")
    for i in range(11):
        vsinstances = GetAllVSInstanceInfo(needChipInfo=True, skipEWDK=True, ignoreCache=True)
        if i % 10 == 0:
            print("i:", i, vsinstances)
