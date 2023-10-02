import os
from comtypes.client import *

class SaveScenario:

    def __init__(self, stkRoot, scenario) -> None:
        if stkRoot == None or scenario == None:
            print("error, please input root and scenario")
            return

        self.root = stkRoot
        self.sc = scenario

    def saveTo(self, path):
        self.root.SaveAs(path)
        print("scenario saved to " + path)

    def getDefaultPath(self) -> str:
        userDir = self.root.ExecuteCommand('GetDirectory / DefaultUser').Item(0)
        scFolder = userDir + '\\' + self.sc.InstanceName
        os.mkdir(scFolder)
        defaultPath = scFolder + '\\' + self.sc.InstanceName
        return defaultPath
    
    def quitRoot(self):
        del(self.root)
        print("stk quit, bye~")
