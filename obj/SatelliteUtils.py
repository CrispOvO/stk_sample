import time
from comtypes.client import *
from comtypes.gen import STKObjects


class Satellite:
    def __init__(self, scenario) -> None:
        self.scenario = scenario
    
    # 根据名字删除卫星  
    def delSatellite(self, satName: str):
        if self.scenario.Children.Contains(STKObjects.eSatellite, satName):
            self.scenario.Children.Item(satName).Unload()
            
    # 在当前场景创建卫星，传入轨道预报模型的编号，默认为7，双星模型
    def createSatellite(self, satName: str, propagator: int = 7, ):
        sat = self.scenario.Children.New(STKObjects.eSatellite, satName)
        # 将上一步生成的对象转为IAgXXX类型，New方法返回的是STKObjects类型的对象，创建后有三种选择：
        # 1、保持该对象类型不变
        # 2、将对象映射为IAgXXX对象
        # 3、将对象映射为AgXXX对象，AgXXX对象同时包含IAgXXX和STKObjects的接口，但是最初AgXXX对象存在一些bug
        sat2 = sat.QueryInterface(STKObjects.IAgSatellite)
        print(sat2)

    # 打印支持的轨道预报模型
    def printPropagatorSupport(sat2):
        print("\n"+str(sat2.PropagatorSupportedTypes) + "\n")


'''
# 打印支持的轨道预报模型


# 设置卫星轨道
# 轨道类型为7，表示卫星的轨道预报模型为双星模型，此时地球视为一个质点
print("propagator type: " + str(satelliteI.Propagator))

propagator = satelliteI.Propagator

# 明确了卫星轨道模型后，将轨道明确为双星模型的轨道
proTwoBodyI = propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody)

# 更新卫星的时间段
epoch = "17 Sep 2018 00:00:00.000"
proTwoBodyI.InitialState.Epoch = epoch

# 使用ICRF坐标系中的经典轨道元素分配卫星的轨道状态,设置轨道的六根数为:
# 半长轴sma 8000 km 离心率e 0 轨道倾角i 60 近心点辐角aop 0 升交点经度raan 0 真近点角ma 0
proTwoBodyI.InitialState.Representation.AssignClassical(3, 8000, 0, 60, 0, 0, 0)

# 在UI界面中画出卫星的轨迹
proTwoBodyI.Propagate()

# 下面使用connect命令创建另一个卫星ConnectSat
cmd = "New / */Satellite ConnectSat"

if scenario.Children.Contains(STKObjects.eSatellite, "ConnectSat"):
    scenario.Children.Item("ConnectSat").Unload()
    
stkRoot.ExecuteCommand(cmd)

# 使用命令完善卫星的信息
cmd = (
    'SetState */Satellite/ConnectSat Classical TwoBody "{}" "{}" 90 ICRF "{}" 7000000.0 0.01 90 270 0 10'.format(scenarioI.StartTime, scenarioI.StopTime, scenarioI.StartTime)
)
stkRoot.ExecuteCommand(cmd)

# 通过路径获取对象
sat2 = stkRoot.GetObjectFromPath("*/Satellite/ConnectSat")
sat2I = sat2.QueryInterface(STKObjects.IAgSatellite)

# 改变卫星的颜色
basicAtt = sat2I.Graphics.Attributes.QueryInterface(STKObjects.IAgVeGfxAttributesBasic)

# RGB十六进制转为十进制
basicAtt.Color = 16777215
'''

    