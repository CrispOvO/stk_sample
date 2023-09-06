import time
from comtypes.client import *
from comtypes.gen import STKObjects


# 对象名称
SN1 = "scenario_1"
FN1 = "facility_1"
SaN1 = "satellite_1"

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

# 按行读取文件
file = open('../satellite_info/satellite_info_test.txt')
info = file.readlines()
file.close()

# 创建卫星对象
for i in range(0, len(info), 3):
    satName = info[i].strip()
    
    # *删除已经存在的，防止报错
    if scenario.Children.Contains(STKObjects.eSatellite, satName):
        scenario.Children.Item(satName).Unload()
    
    satellite = scenario.Children.New(STKObjects.eSatellite, satName).QueryInterface(STKObjects.IAgSatellite)
    # 设置卫星预报模型SGP4
    satellite.SetPropagatorType(4)
    # 将预报模型转化为IAg
    pgSGP4I = satellite.Propagator.QueryInterface(STKObjects.IAgVePropagatorSGP4)
    # !读取文件中对应编号的卫星轨道参数，需要绝对路径
    pgSGP4I.CommonTasks.AddSegsFromFile(info[i + 1].split(' ')[1], r'D:/APOvO/workspace/python_workspace/stk/satellite_info/satellite_info_test.txt')
    
    pgSGP4I.Propagate()

