# -*- coding: utf-8 -*-
"""
Created on Thu Jan 28 13:01:45 2021

@author: Jackson Artis
"""

# This script is to access satellite data after each prop segment
import os
from comtypes.client import GetActiveObject
from comtypes.gen import STKObjects
from comtypes.gen import AgStkGatorLib

# Set Ephemeris location
folderpath = "C:\\Users\\jartis\\Documents\\STK 12\\Astrogator Scenarios"
os.chdir(folderpath)

# Grab STK Instance
app = GetActiveObject('STK12.Application')
root = app.Personality2
scenario = root.CurrentScenario
scenario2 = scenario.QueryInterface(STKObjects.IAgScenario)

# Grab Satellite Object
satName = 'GEO_Sat_EW'
sat = scenario.Children.Item(satName)
satelliteObj = sat.QueryInterface(STKObjects.IAgSatellite)

finalEphData = []
runs = [1,2,3]

for j in range(1000):
    # Begin Looping Process
    propagator = satelliteObj.Propagator.QueryInterface(AgStkGatorLib.IAgVADriverMCS)
    MCS = propagator.MainSequence 
    lastSeg = MCS.Item(MCS.Count - 2)
    propSeg = lastSeg.QueryInterface(AgStkGatorLib.IAgVAMCSPropagate)
    finalState = lastSeg.FinalState
    finalState.SetElementType(0)
    element = finalState.Element.QueryInterface(AgStkGatorLib.IAgVAElementCartesian)
    xPos = element.X
    yPos = element.Y
    zPos = element.Z
    xVel = element.Vx
    yVel = element.Vy
    zVel = element.Vz
    xval = 1
    yval = 2
    zval = 3
    newMan = MCS.InsertByName('Maneuver','-')
    maneuver = newMan.QueryInterface(AgStkGatorLib.IAgVAMCSManeuver)
    maneuver.SetManeuverType(0)
    maneuver.Maneuver.QueryInterface(AgStkGatorLib.IAgVAManeuverImpulsive).SetAttitudeControlType(4)
    thrustVec = maneuver.Maneuver.QueryInterface(AgStkGatorLib.IAgVAManeuverImpulsive).AttitudeControl.QueryInterface(AgStkGatorLib.IAgVAAttitudeControlImpulsiveThrustVector)
    thrustVec.AssignCartesian(xval,yval,zval)
    newProp = MCS.InsertByName('Propagate','-')
    propagate = newProp.QueryInterface(AgStkGatorLib.IAgVAMCSPropagate)
    stopCon = propagate.StoppingConditions.Item('Duration')
    stopCon2 = stopCon.QueryInterface(AgStkGatorLib.IAgVAStoppingConditionElement).Properties.QueryInterface(AgStkGatorLib.IAgVAStoppingCondition)
    stopCon2.Trip = 1
    propagator.RunMCS()
    
    k = 0
    if (j % 100 == 0):
        exporter = satelliteObj.ExportTools.GetEphemerisStkExportTool()
        exporter.Export(folderpath + '\\Satellite.e')
        MCS.RemoveAll()
        
        oldFile = open('Satellite.e', 'r');
        lines = oldFile.readlines()
        realLineLength = len(lines)
        lines.append('\n')
        print(realLineLength)
        print(len(lines))
        
        
        
        i = 0
        
        while lines[i-2] != '    EphemerisTimePosVel\t\t\n':
            i = i + 1
        if (k == 0):
            currentDataPoints = lines[i:len(lines)]
        else: 
            currentDataPoints = lines[i+1:len(lines)]
        finalEphData = finalEphData + currentDataPoints
        oldFile.close()
        k = k + 1
    
    

## Build the final ephemeris

# Create header
header1 = 'stk.v.4.3\nBEGIN Ephemeris\nNumberOfEphemerisPoints 1439\nScenarioEpoch           '+ scenario2.StartTime + '\nInterpolationMethod     Lagrange\n'
header2 = 'InterpolationOrder      7\nDistanceUnit\t\t\tKilometers\nCentralBody             Earth\nCoordinateSystem        Fixed\n'
header3 = 'EphemerisTimePosVel\n'

fullHeader = header1 + header2 + header3

newFile = open('finalEphemeris.e', 'w+')
newFile.write(fullHeader)

for i in range(len(finalEphData)):
    newFile.write(finalEphData[i])
    
newFile.close()

lastSat = scenario.Children.New(18, 'Final_Satellite')
lastSatObj = lastSat.QueryInterface(STKObjects.IAgSatellite)
lastSatObj.SetPropagatorType(6)
propagator = lastSatObj.Propagator.QueryInterface(STKObjects.IAgVePropagatorStkExternal)
propagator.Filename = folderpath + '\\' + 'finalEphemeris.e'
propagator.Propagate()




