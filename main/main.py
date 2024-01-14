import sys
import os
#! due to the path, you have to run this file at parent folder
sys.path.append(os.path.abspath(os.curdir))
from obj.Satellite import Satellite
from misc.Scenario import Scenario
# from analysis.dataprovider import DataProvider

if __name__ == '__main__':
    
    # sc
    # scName = "myGoodSc"
    # scPath = os.path.abspath(os.curdir) + "/sc_data/" + scName + ".sc"
    scName = "Square"
    scPath = r'E:\sc\Square\Square.sc'
    scenario = Scenario()
    scenario.loadScenario(scName, scPath)
    sc = scenario.getScenario()
    
    # # satellite
    # sat1Name = "satellite1"
    # satellite = Satellite(sc, scenario.stkRoot)
    # # sat1 = satellite.createSatellite(sat1Name)
    # sat1 = satellite.getSatelliteByName(sat1Name)
    
    
    # # dataprovider
    # dataprovider = DataProvider(scenario.stkRoot, sc)
    # result = dataprovider.getResult(sat1)
    # dataprovider.showResult(result)
    # dataprovider.drawResult(result, "speed")
    
    # save
    # scenario.saveTo(scPath)
    
    
    
    