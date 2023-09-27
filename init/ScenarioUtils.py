import time
from comtypes.client import *
from comtypes.gen import STKObjects
from win32api import GetSystemMetrics

class Scenario:
    # 如果使用已经打开的STK场景，会节省创建的时间
    def __init__(self, readScenario: bool = False) -> None:
        
        self.readScenario = readScenario
        # 获取STK的UI界面
        if readScenario:
            uiApplication = GetActiveObject('STK11.Application')
        else:
            uiApplication = CreateObject("STK11.Application")
        uiApplication.Visible = True
        uiApplication.UserControl = True

        # 设置窗口的位置和大小
        uiApplication.Top = 0
        uiApplication.Left = 0
        uiApplication.Width = int(GetSystemMetrics(0)/2)
        uiApplication.Height = int(GetSystemMetrics(1) - 30)

        # 获取 IAgStkObjectRoot 接口
        self.stkRoot = uiApplication.Personality2

    def createScenario(self, scName: str):
        startTime = time.time()
        if not self.readScenario:
            # 创建场景
            self.stkRoot.newScenario(scName)

        # 获取当前场景
        sc = self.stkRoot.CurrentScenario
        sc2 = self.sc.QueryInterface(STKObjects.IAgScenario)

        # 设置场景时间
        sc2.SetTimePeriod("Today", "+24")

        # 将场景回溯到开始时间
        self.stkRoot.Rewind()

        # Print time spent for scenario creation
        print("Scenario creation using {totalTime: f} sec".format(totalTime = time.time() - startTime))
        
        return sc