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

# 设置卫星轨道
# 轨道类型为7，表示卫星的轨道预报模型为双星模型，此时地球视为一个质点
print("propagator type: " + str(satelliteI.Propagator))

propagator = satelliteI.Propagator

# 明确了卫星轨道模型后，将轨道转化为双星模型
proTwoBodyI = propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody)

# 设置轨道的六根数为：
# 半长轴 8000 km
# 离心率 0 -- 圆
# 轨道倾角 60
# 近心点辐角 0
# 升交点经度 0
# 真近点角 0
proTwoBodyI.InitialState.Representation.AssignClassical(3, 8000, 0, 60, 0, 0, 0)

proTwoBodyI.Propagate()

# 创建成功，日期 2023/9/3
