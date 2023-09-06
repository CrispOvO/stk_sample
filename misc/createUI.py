from comtypes.client import CreateObject
import time

#! 注意，必须是通过程序启动的场景，才能直接使用程序进行操作，手动启动的STKUI不能直接使用程序进行操作

# 记录开始的时间
startTime = time.time()

# 创建STK11桌面应用
uiApplication = CreateObject("STK11.Application")

# STK11可见
uiApplication.Visible = True
uiApplication.UserControl = True

# 获取IAgStkObjectRoot接口
stkRoot = uiApplication.Personality2

# 读取Scenario1.sc文件
stkRoot.LoadScenario(r'C:\Users\Crisp\Documents\STK 11 (x64)\scenario_1\scenario_1.sc')

print("scenario initialization using {} sec".format(time.time() - startTime))