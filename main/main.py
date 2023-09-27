import sys
import os

sys.path.append(os.path.abspath(os.curdir))

from obj.SatelliteUtils import Satellite
from init.ScenarioUtils import Scenario

if __name__ == '__main__':
    readScenario = True
    scName = "myGoodSc"
    ScenarioUtils = Scenario(readScenario)
    sc = ScenarioUtils.getScenario()
    sat = Satellite(sc, ScenarioUtils.stkRoot)
    s1Name = "starlink0_0"
    sat.delSatellite(s1Name)
    s1 = sat.createSatellite(s1Name)
    sat.setOrbit(s1)