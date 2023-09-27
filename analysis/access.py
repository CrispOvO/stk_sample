import time
import numpy as np
import pandas as pd
from comtypes.client import *
from comtypes.gen import STKObjects


class Access:
    def __init__(self) -> None:
        pass
    
    # 在obj1和obj2之间建立链接，然后计算相关数据
    def getAccess(self, scenario, obj1Name, obj2):
        
        # 通过objName获取初始对象
        obj1 = scenario.Children.Item(obj1Name)
        obj1.QueryInterface(STKObjects.IAgStkObject)
        
        access = obj1.GetAccessToObject(obj2)  
        # 进行计算
        access.ComputeAccess()
        # 取得计算的结果，为相互连接的开始和结束时间
        accessResult = access.ComputedAccessIntervalTimes.ToArray(0, -1)
        # 将元组转化为数组，然后展示数据
        accessResult = np.asarray(accessResult)

        # x[:,?]代表在x中所有数组（维）中，取第？个数据
        accessTable = pd.DataFrame({"Start Time": accessResult[:, 0], "End Time": accessResult[:, 1]})

        print(accessTable)
        
        return 
