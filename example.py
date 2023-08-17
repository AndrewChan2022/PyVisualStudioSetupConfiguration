from PyVisualStudioSetupConfiguration import GetAllVSInstanceInfo

if __name__ == '__main__':

    # get list of VisualStudio Instance Info
    vsInstances = GetAllVSInstanceInfo()

    # print the list
    print(vsInstances)

    # get version of first instance
    if len(vsInstances):
        vsInstance = vsInstances[0]

        s = 'vsInstances[0] :\n'
        s += f'    VSInstallLocation   : {vsInstance.VSInstallLocation}\n'
        s += f'    Version             : {vsInstance.getVersion()}\n'
        s += f'    versionMajor        : {vsInstance.getVerionMajor()}\n'
        s += f'    VCToolsetVersion    : {vsInstance.VCToolsetVersion}\n'
        s += f'    bWin10SDK           : {vsInstance.bWin10SDK}\n'
        s += f'    bWin81SDK           : {vsInstance.bWin81SDK}\n'
        s += f'    chip                : {vsInstance.chip}\n'
        print(s)

