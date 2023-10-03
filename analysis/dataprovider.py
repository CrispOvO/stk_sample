import pandas as pd
import matplotlib.pyplot as plt
from comtypes.client import *
from comtypes.gen import STKObjects

from misc.Scenario import Scenario

class DataProvider:
    
    def __init__(self, stkRoot, scenario) -> None:
        self.root = stkRoot
        self.sc = scenario
        self.sc2 = Scenario.getScenarioI(self.sc)


    # 获取dataprovider中的time speed radial in-track信息，并将输出存入result中
    # !注意想要获取参数名称的大小写不要输错，不然会报错
    def getResult(self, sat: STKObjects, elems = ["Time", "speed", "radial", "in-track"]) -> dict:
    
        # 获取笛卡尔速度,dataprovider中有一些我们想要获取的属性，不同provider中包含的属性不同
        cartV = sat.Dataproviders.Item("Cartesian Velocity")
        cartVI = cartV.QueryInterface(STKObjects.IAgDataProviderGroup)

        # 获取Cartesian Velocity下ICRF文件夹中的数据，icrf是一种坐标系，其内可获得的属性有xyz、time、speed等
        cartVicrf = cartVI.Group.Item("ICRF")
        cartVicrfI = cartVicrf.QueryInterface(STKObjects.IAgDataPrvTimeVar)

        # 设置timestep为60，并获取数据
        # rawResult type: STKObjects.IAgDrResult
        rawResult = cartVicrfI.ExecElements(self.sc2.StartTime, self.sc2.StopTime, 60, elems)
        
        # 从result中获取数据
        result = {}
        for i in range(len(elems)):
            result[elems[i]] = rawResult.DataSets.Item(i).GetValues()
        return result
    
    
    def showResult(self, result):
        # 使用pandas展示前五条数据信息
        dataframe = pd.DataFrame(result)
        print(dataframe.head(5))
        
        
    def drawResult(self, result, var):
        # 绘制速度随时间变化的图像
        plt.plot(result[var])
        plt.xlabel("Time [mins]")
        plt.ylabel("Speed [km/sec]")
        plt.title("Speed vs Time")
        plt.show()



