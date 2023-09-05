import time
from comtypes.client import *
from comtypes.gen import STKObjects
from win32api import GetSystemMetrics

# 对象名称
SN1 = "scenario_1"
SaN1 = "satellite_1"
FN1 = "facility_1"
Target_facility_name = "Facility1"

# 如果使用已经打开的STK场景，会节省创建的时间
Read_Scenario = False

# 记录开始的时间
startTime = time.time()

# 获取STK的UI界面
if Read_Scenario:
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
stkRoot = uiApplication.Personality2

if not Read_Scenario:
    # 创建场景
    stkRoot.newScenario(SN1)

# 获取当前场景
scenario = stkRoot.CurrentScenario
scenarioI = scenario.QueryInterface(STKObjects.IAgScenario)

# 设置场景时间
scenarioI.SetTimePeriod("Today", "+24")

# 将场景回溯到开始时间
stkRoot.Rewind()

# Print time spent for scenario creation
print("Scenario creation using {totalTime: f} sec".format(totalTime = time.time() - startTime))