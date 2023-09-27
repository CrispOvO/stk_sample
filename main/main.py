import sys
import os

sys.path.append(os.path.abspath(os.curdir))

from obj.SatelliteUtils import Satellite
from init.ScenarioUtils import Scenario
from analysis.access import Access

if __name__ == '__main__':
    
    # sc
    readScenario = True
    scName = "myGoodSc"
    ScenarioUtils = Scenario(readScenario)
    sc = ScenarioUtils.getScenario()
    
    # sat
    sat = Satellite(sc, ScenarioUtils.stkRoot)
    s1Name = "starlink0_0"
    s2Name = "ConnectSat"
    s1 = sat.getSatelliteByName(s1Name)
    s2 = sat.getSatelliteByName(s2Name)
    
    # access
    acc = Access()
    acc.getAccess(sc, s1Name, s2)
    
    
    
    