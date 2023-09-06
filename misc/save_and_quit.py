import time
import os 
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

# 获取 IAgStkObjectRoot 接口
stkRoot = uiApplication.Personality2

scenario = stkRoot.CurrentScenario

userDir = stkRoot.ExecuteCommand('GetDirectory / DefaultUser').Item(0)
saveFolder = userDir + '\\' + scenario.InstanceName
os.mkdir(saveFolder)
savePath = saveFolder + '\\' + scenario.InstanceName
stkRoot.SaveAs(savePath)
print("scenario saved to " + savePath)

# 退出应用
del(stkRoot)
uiApplication.Quit()
del(uiApplication)

print("stk quit, bye~")