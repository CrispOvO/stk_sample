import time
import pandas as pd
import matplotlib.pyplot as plt
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

# 获取卫星
sat = scenario.Children.Item(SaN1)
satI = sat.QueryInterface(STKObjects.IAgSatellite)

# 获取笛卡尔速度,dataprovider中有一些我们想要获取的属性，不同provider中包含的属性不同
cartV = sat.Dataproviders.Item("Cartesian Velocity")
cartVI = cartV.QueryInterface(STKObjects.IAgDataProviderGroup)

# 获取Cartesian Velocity下ICRF文件夹中的数据，icrf是一种坐标系，其内可获得的属性有xyz、time、speed等
cartVicrf = cartVI.Group.Item("ICRF")
cartVicrfI = cartVicrf.QueryInterface(STKObjects.IAgDataPrvTimeVar)

# 获取dataprovider中的time speed radial in-track信息，并将输出存入result中
# !注意想要获取参数名称的大小写不要输错，不然会报错
elems = ["Time", "speed", "radial", "in-track"]

# 设置timestep为60，并获取数据
result = cartVicrfI.ExecElements(scenarioI.StartTime, scenarioI.StopTime, 60, elems)

# 从result中获取数据
Time = result.DataSets.Item(0).GetValues()
speed = result.DataSets.Item(1).GetValues()

# 使用pandas展示前五条数据信息
dataframe = pd.DataFrame({"Time": Time, "Speed": speed})
print(dataframe.head(5))

# 绘制速度随时间变化的图像
plt.plot(speed)
plt.xlabel("Time [mins]")
plt.ylabel("Speed [km/sec]")
plt.title("Speed vs Time")
plt.show()

