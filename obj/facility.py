import time
from comtypes.client import *
from comtypes.gen import STKObjects


# 名称缩写
SN1 = "scenario_1"
FN1 = "facility_1"

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

# Print time spent for scenario creation
print("Scenario creation using {totalTime: f} sec".format(totalTime = time.time() - startTime))

# 接下来在当前场景创建 facility， 如果存在名字一样的facility，则先删除它 
if scenario.Children.Contains(STKObjects.eFacility, FN1):
    scenario.Children.Item(FN1).Unload()

facility = scenario.Children.New(STKObjects.eFacility, FN1)

# 将上一步生成的对象转为IAgFacility类型，New方法返回的是STKObjects类型的对象，创建后有三种选择：
# 1、保持该对象类型不变
# 2、将对象映射为IAgXXX对象
# 3、将对象映射为AgXXX对象，AgXXX对象同时包含IAgXXX和STKObjects的接口，但是最初AgXXX对象存在一些bug
facilityI = facility.QueryInterface(STKObjects.IAgFacility)

# 设置设备的位置 (Latitude, Longitude, Altitude)
lat = 39.0095
lon = -76.896
alt = 0
facilityI.Position.AssignGeodetic(lat, lon, alt)

# 不使用地形数据
facilityI.UseTerrain = False

# 定义方向-仰角遮罩 (Azimuth-elevation mask) 以便使用该遮罩作为计算接入 (Access) 时的约束。
facilityI.SetAzElMask(1,0) # eTerrainData 

