from PyVisualStudioSetupConfiguration import GetAllVSInstanceInfo

if __name__ == '__main__':

    print(GetAllVSInstanceInfo())

    # benchmark
    print("*********performance benchmark:**************")
    for i in range(11):
        vsinstances = GetAllVSInstanceInfo(needChipInfo=True, skipEWDK=True, ignoreCache=True)
        if i % 1 == 0:
            print("i:", i, vsinstances)
