import time
from comtypes.client import *
from comtypes.gen import STKObjects


class Satellite:
    def __init__(self, scenario: STKObjects.IAgScenario, stkRoot) -> None:
        
        self.scenario = scenario
        self.scenario2 = self.scenario.QueryInterface(STKObjects.IAgScenario)
        self.stkRoot = stkRoot

    
    # 根据名字删除卫星
    def delSatellite(self, satName: str) -> None:
        
        if self.scenario.Children.Contains(STKObjects.eSatellite, satName):
            self.scenario.Children.Item(satName).Unload()
        else:
            print(f"satellite {satName} does not exist.")

            
    # 在当前场景创建卫星
    def createSatellite(self, satName: str) -> STKObjects.IAgSatellite:
        
        if self.scenario.Children.Contains(STKObjects.eSatellite, satName):
            print(f"can not create satellite because it already exists.")
            
        sat = self.scenario.Children.New(STKObjects.eSatellite, satName)
        # 将上一步生成的对象转为IAgXXX类型，New方法返回的是STKObjects类型的对象，创建后有三种选择：
        # 1、保持该对象类型不变
        # 2、将对象映射为IAgXXX对象
        # 3、将对象映射为AgXXX对象，AgXXX对象同时包含IAgXXX和STKObjects的接口，但是最初AgXXX对象存在一些bug
        sat2 = sat.QueryInterface(STKObjects.IAgSatellite)
        
        return sat2
    
    # 使用命令行创建卫星
    def createSatelliteUsingCMD(self, scenario) -> STKObjects.IAgSatellite:
        
        # 使用connect命令创建另一个卫星ConnectSat
        cmd = "New / */Satellite ConnectSat"

        if scenario.Children.Contains(STKObjects.eSatellite, "ConnectSat"):
            scenario.Children.Item("ConnectSat").Unload()
            
        self.stkRoot.ExecuteCommand(cmd)

        # 使用命令完善卫星的信息
        cmd = (
            f'SetState */Satellite/ConnectSat Classical TwoBody "{self.scenario2.StartTime}" "{self.scenario2.StopTime}" 90 ICRF "{self.scenario2.StartTime}" 7000000.0 0.01 90 270 0 10'
        )
        self.stkRoot.ExecuteCommand(cmd)

        # 通过路径获取对象
        sat = self.stkRoot.GetObjectFromPath("*/Satellite/ConnectSat")
        sat2 = sat.QueryInterface(STKObjects.IAgSatellite)
        return sat2
    
    # 通过卫星名字获取卫星
    def getSatelliteByName(self, satName: str) -> STKObjects.IAgSatellite:
        
        # 如果不存在，打印报错信息
        if not self.scenario.Children.Contains(STKObjects.eSatellite, satName):
            print(f"satellite {satName} does not exist.")
            return None
        
        sat = self.scenario.Children.Item(satName)
        sat2 = sat.QueryInterface(STKObjects.IAgSatellite)

        return sat2

            
        
    # 改变卫星的颜色
    def changeColor(self, sat2: STKObjects.IAgSatellite, color: int = 16777215):
        
        basicAtt = sat2.Graphics.Attributes.QueryInterface(STKObjects.IAgVeGfxAttributesBasic)
        # RGB十六进制转为十进制
        basicAtt.Color = color

        
    # 根据轨道6根数和轨道预报模型确认卫星轨道
    def setOrbit(self, sat2: STKObjects.IAgSatellite, ele: list = [11, 7000, 0.01, 90, 270, 90, 10] ,propagatorType: int = 7):
        
        propagator = sat2.Propagator
        prop = None
        
        # 根据编号确定轨道模型
        if propagatorType == 7:
            prop = propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody)
        prop.Propagate()
        
        # 确定轨道六根数
        coordSys = ele[0]; sma = ele[1]; e = ele[2]; i = ele[3]; aop = ele[4]; raan = ele[5]; ma = ele[6]
        prop.InitialState.Representation.AssignClassical(coordSys, sma, e, i, aop, raan, ma)
        prop.Propagate() 


    # 打印支持的轨道预报模型
    def printPropagatorSupport(self, sat2: STKObjects.IAgSatellite):
        print(*sorted(sat2.PropagatorSupportedTypes), sep = "\n")

    