import time
from comtypes.client import *
from comtypes.gen import STKObjects

# 对象名称
SN1 = "scenario_1"
SaN1 = "satellite_1"
FN1 = "facility_1"
Target_facility_name = "Facility1"

# 使用已经打开的STK场景，会节省创建的时间
Read_Scenario = True

# 记录开始的时间
startTime = time.time()

# 获取STK的UI界面
if Read_Scenario:
    uiApplication = GetActiveObject('STK11.Application')
else:
    uiApplication = CreateObject("STK11.Application")
    
uiApplication.Visible = True
uiApplication.UserControl = True

# 获取 IAgStkObjectRoot 接口
stkRoot = uiApplication.Personality2

if not Read_Scenario:
    # 创建场景
    stkRoot.newScenario(SN1)

# 获取当前场景
scenario = stkRoot.CurrentScenario
scenarioI = scenario.QueryInterface(STKObjects.IAgScenario)

# Print time spent for scenario creation
print("Scenario creation using {totalTime: f} sec".format(totalTime = time.time() - startTime))

# 接下来在当前场景创建卫星，如果存在名字一样的卫星，则先删除它 
if scenario.Children.Contains(STKObjects.eSatellite, SaN1):
    scenario.Children.Item(SaN1).Unload()

satellite = scenario.Children.New(STKObjects.eSatellite, SaN1)

# 将上一步生成的对象转为IAgXXX类型，New方法返回的是STKObjects类型的对象，创建后有三种选择：
# 1、保持该对象类型不变
# 2、将对象映射为IAgXXX对象
# 3、将对象映射为AgXXX对象，AgXXX对象同时包含IAgXXX和STKObjects的接口，但是最初AgXXX对象存在一些bug
satelliteI = satellite.QueryInterface(STKObjects.IAgSatellite)

# 打印支持的轨道预报模型
print("\n"+str(satelliteI.PropagatorSupportedTypes) + "\n")

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

'''
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