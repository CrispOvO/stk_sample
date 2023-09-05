import time
from comtypes.client import *
from comtypes.gen import STKObjects, AgSTKVgtLib
from win32api import GetSystemMetrics

# 对象名称
SN1 = "scenario_1"
SaN1 = "satellite_1"
FN1 = "facility_1"
VeN = "toagi_greenbelt"
Target_facility_name = "Facility1"

# 如果使用已经打开的STK场景，会节省创建的时间
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

# 将场景回溯到开始时间
stkRoot.Rewind()

# Print time spent for scenario creation
print("Scenario creation using {totalTime: f} sec".format(totalTime = time.time() - startTime))

# 获取需要创建向量的卫星
sat = scenario.Children.Item(SaN1)
satI = sat.QueryInterface(STKObjects.IAgSatellite)

# 从向量工厂创建一个新的向量 type:IAgCrdnVectorFactory
vecFac = sat.Vgt.Vectors.Factory
# type: IAgCrdnVector 
try:
    toFac = vecFac.Create(VeN, "Description", AgSTKVgtLib.eCrdnVectorTypeDisplacement)
except:
    # in case the vector has already been created
    toFac = sat.Vgt.Vectors.Item(VeN)
    
toFacI = toFac.QueryInterface(AgSTKVgtLib.IAgCrdnVectorDisplacement)

# 指定向量的原点和指向地点
toFacI.Origin.SetPath("Satellite/{} Center".format(SaN1))
toFacI.Destination.SetPath("Facility/{} Center".format(FN1))

# 将创建的向量连接到卫星上
vecVo = satI.VO.Vector.RefCrdns.Add(0, "Satellite/{} {}".format(SaN1, VeN))
vecVo.Visible = True