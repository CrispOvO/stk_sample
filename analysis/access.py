import time
import numpy as np
import pandas as pd
from comtypes.client import *
from comtypes.gen import STKObjects


# 名称缩写
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

# Print time spent for scenario creation
print("Scenario creation using {totalTime: f} sec".format(totalTime = time.time() - startTime))

# 获取场景中的地面设施
fac = scenario.Children.Item(FN1)
facI = fac.QueryInterface(STKObjects.IAgFacility)

# 获取场景中的卫星
sat = scenario.Children.Item(SaN1)
satI = sat.QueryInterface(STKObjects.IAgSatellite)

# 创建卫星到地面基站的评估
access = sat.GetAccessToObject(facI)

# 进行计算
access.ComputeAccess()

# 取得计算的结果，为有连接的开始和结束时间
accessResult = access.ComputedAccessIntervalTimes.ToArray(0, -1)

# 将元组转化为数组，然后展示数据
accessResult = np.asarray(accessResult)

# x[:,?]代表在x中所有数组（维）中，取第？个数据
accessTable = pd.DataFrame({"Start Time": accessResult[:, 0], "End Time": accessResult[:, 1]})
print(accessTable)
