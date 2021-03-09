# -*- coding: utf-8 -*-
"""
Created on Fri Jan 29 09:22:05 2021

@author: Jackson Artis
"""

from comtypes.client import GetActiveObject
from comtypes.gen import STKObjects
from comtypes.gen import STKUtil

# Grab STK Instance
app = GetActiveObject('STK12.Application')
root = app.Personality2
scenario = root.CurrentScenario
scenario2 = scenario.QueryInterface(STKObjects.IAgScenario)

# Grab Satellite Object
satName = 'Satellite3'
launchVehicle = 'LaunchVehicle1'
sat = scenario.Children.Item(satName)
satelliteObj = sat.QueryInterface(STKObjects.IAgSatellite)


# Set Constant Parameters
ecc = 0
argOfPerigee = 0
raan = 0
trueAnomaly = 0


# Set Varying Parameters
altitudes = [8000,9000,10000]
inclinations = [10,20,30]
numPlanes = [1,2,3]
numSats = [1,2,3]

for i in range(len(altitudes)):
    # Set Satellite Parameters
    propagator = satelliteObj.Propagator
    twoBody = propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody)
    keplerian = twoBody.InitialState.Representation.ConvertTo(1)
    keplerian2 = keplerian.QueryInterface(STKObjects.IAgOrbitStateClassical)
    keplerian2.SizeShapeType = 4 # This line will make it so that you specify SMA & ecc. To specificy apogee/perigee alt, make it keplerian2.SizeShapeType = 0
    shapeValues = keplerian2.SizeShape.QueryInterface(STKObjects.IAgClassicalSizeShapeSemimajorAxis)
    shapeValues.Eccentricity = ecc # If SizeShapeType 0, this line should be shapeValues.PerigeeAltitude = altitudes[i]
    shapeValues.SemiMajorAxis = altitudes[i] # If SizeShapeType 0, this line should be shapeValues.ApogeeAltitude  = altitudes[i]
    keplerian2.Orientation.Inclination = inclinations[i]
    keplerian2.Orientation.ArgOfPerigee = argOfPerigee
    keplerian2.Orientation.AscNode.QueryInterface(STKObjects.IAgOrientationAscNodeRAAN).Value = raan
    keplerian2.Location.QueryInterface(STKObjects.IAgClassicalLocationTrueAnomaly).Value = trueAnomaly
    satelliteObj.Propagator.QueryInterface(STKObjects.IAgVePropagatorTwoBody).InitialState.Representation.Assign(keplerian2)
    twoBody.Propagate()
    
    # Change Walker Constellation
    connectCommand = 'Walker */Satellite/' + satName + " " + 'Type Delta NumPlanes' 
    connectCommand = connectCommand + " " + str(numPlanes[i]) + " " + 'NumSatsPerPlane' + " " + str(numSats[i]) + " " 
    connectCommand = connectCommand + 'InterPlanePhaseIncrement 0 ColorByPlane Yes ConstellationName SatConstellation'   
    root.ExecuteCommand(connectCommand)
    
    # Chain Connect Command
    if (scenario.Children.GetElements(4).Count == 0): 
         newChain = 'New / */Chain NewChain'
         root.ExecuteCommand(newChain)
         
    # Grab Hold of Chain
    chain = scenario.Children.GetElements(4).Item(0).InstanceName
    
    # Add satellite and Launch vehivle
    satCommand = 'Chains */Chain/' + chain + " " + 'Add Constellation/SatConstellation' 
    LV = 'Chains */Chain/' + chain + " " + 'Add LaunchVehicle/' + launchVehicle 
    computeCommand = 'Chains */Chain/' + chain + " " + 'Compute'
    root.ExecuteCommand(satCommand)
    root.ExecuteCommand(LV)
    root.ExecuteCommand(computeCommand)
    
    # Get Data Providers
    chainObj = scenario.Children.GetElements(4).Item(0)
    TOA = chainObj.DataProviders.Item('Time Ordered Access').QueryInterface(STKObjects.IAgDataPrvInterval)
    toaDP = TOA.Exec(scenario2.StartTime, scenario2.StopTime)
    toaStartTimes = toaDP.DataSets.GetDataSetByName('Start Time').GetValues()
    toaEndTimes = toaDP.DataSets.GetDataSetByName('Stop Time').GetValues()
    toaDuration = toaDP.DataSets.GetDataSetByName('Duration').GetValues()
    toaStandName = toaDP.DataSets.GetDataSetByName('Strand Name').GetValues()
    
    

    # Remove satellite and Launch Vehicle
    satCommand = 'Chains */Chain/' + chain + " " + 'Remove Constellation/SatConstellation' 
    LV = 'Chains */Chain/' + chain + " " + 'Remove LaunchVehicle/' + launchVehicle 
    computeCommand = 'Chains */Chain/' + chain + " " + 'ClearAccesses'
    root.ExecuteCommand(satCommand)
    root.ExecuteCommand(LV)
    root.ExecuteCommand(computeCommand)