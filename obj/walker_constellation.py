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

# 构建Walker星座需要有种子卫星，从现有场景中读取
# 执行connect命令，选择Delta类型形成Walker星座，总卫星数量24颗、4个轨道面、每个轨道面6颗卫星相位因子1
command = 'Walker */Satellite/{satn} Type Delta NumPlanes {np} NumSatsPerPlane {nspp} InterPlanePhaseIncrement {ippi} ColorByPlane Yes'.format(satn = SaN1, np = 4, nspp = 6, ippi = 1)
print("command: " + command)
stkRoot.ExecuteCommand(command)

# 注意需要STKObject类型的变量scenario
scenario.Children.Item(SaN1).Unload()