import time
import os
from comtypes.client import *
from comtypes.gen import STKObjects
from win32api import GetSystemMetrics

class Scenario:

    # 如果使用已经打开的STK场景，会节省创建的时间
    def __init__1(self, scName: str = "tempScName", readScenario: bool = False) -> None:
        
        startTime = time.time()
        
        # 获取STK的UI界面
        if readScenario:
            uiApplication = GetActiveObject('STK11.Application')
        else:
            uiApplication = CreateObject('STK11.Application')
        uiApplication.Visible = True
        uiApplication.UserControl = True

        # 设置窗口的位置和大小
        uiApplication.Top = 0
        uiApplication.Left = int(GetSystemMetrics(0)/2)
        uiApplication.Width = int(GetSystemMetrics(0)/2)
        uiApplication.Height = int(GetSystemMetrics(1) - 50)

        # 获取 IAgStkObjectRoot 接口
        self.stkRoot = uiApplication.Personality2
        
        if not readScenario:
            # 创建场景
            self.stkRoot.newScenario(scName)

        # 获取当前场景
        self.sc = self.stkRoot.CurrentScenario
        
        # 设置场景时间
        sc2 = self.sc.QueryInterface(STKObjects.IAgScenario)
        sc2.SetTimePeriod("Today", "+24")

        # 将场景回溯到开始时间
        self.stkRoot.Rewind()

        # Print time spent for scenario creation
        print("Getting scenario using {totalTime: f} sec".format(totalTime = time.time() - startTime))
        
    # 启动STK的UI界面
    def __init__(self) -> None:
        startTime = time.time()
        try:
            uiApplication = GetActiveObject('STK11.Application')
        except Exception:
            uiApplication = CreateObject('STK11.Application')
        uiApplication.Visible = True
        uiApplication.UserControl = True
        # 设置窗口的位置和大小
        uiApplication.Top = 0
        uiApplication.Left = int(GetSystemMetrics(0)/2)
        uiApplication.Width = int(GetSystemMetrics(0)/2)
        uiApplication.Height = int(GetSystemMetrics(1) - 50)
        
        # 获取 IAgStkObjectRoot 接口
        self.stkRoot = uiApplication.Personality2
        
        print("Getting UI using {totalTime: f} sec".format(totalTime = time.time() - startTime))
        return
        
    def loadScenario(self, scName = "tmpScName" ,scPath = ""):
        # 判断是否关闭当前场景
        if self.stkRoot.Children.Count != 0:
            inputText = input("close the current scenario?([y])")
            if inputText != 'y':
                self.sc = self.stkRoot.CurrentScenario
                return
            else:
                self.stkRoot.CurrentScenario.Unload()
                
        # 重新计算场景数量
        if self.stkRoot.Children.Count == 0:
            if scPath == "":
                self.stkRoot.newScenario(scName)
            else:
                self.stkRoot.LoadScenario(scPath)
                
        # 获取当前场景
        self.sc = self.stkRoot.CurrentScenario
        
        # 设置场景时间
        sc2 = self.sc.QueryInterface(STKObjects.IAgScenario)
        sc2.SetTimePeriod("Today", "+24")

        # 将场景回溯到开始时间
        self.stkRoot.Rewind()
        return
        
        
    # 获取场景，可以是当前的，也可以是重新创建的
    def getScenario(self) -> STKObjects:
        return self.sc
    
    
    def getScenarioI(sc) -> STKObjects.IAgScenario:
        sc2 = sc.QueryInterface(STKObjects.IAgScenario)
        return sc2
    
    
    # save scenario
    def saveTo(self, path):
        self.stkRoot.SaveAs(path)
        print("scenario saved to " + path)

    # 获取默认存储路径
    def getDefaultPath(self) -> str:
        userDir = self.stkRoot.ExecuteCommand('GetDirectory / DefaultUser').Item(0)
        scFolder = userDir + '\\' + self.sc.InstanceName
        os.mkdir(scFolder)
        defaultPath = scFolder + '\\' + self.sc.InstanceName
        return defaultPath
    
    # 退出程序
    def quit(self):
        del(self.stkRoot)
        print("stk quit, bye~")