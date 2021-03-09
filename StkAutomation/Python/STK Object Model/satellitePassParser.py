# -*- coding: utf-8 -*-
"""
Created on Tue Jan 12 12:52:59 2021

@author: Jackson Artis
"""

# Import necessary modules
from comtypes.client import GetActiveObject
from comtypes.gen import STKObjects
from comtypes.gen import AgSTKVgtLib


# Grab STK Instance
app = GetActiveObject('STK12.Application')
root = app.Personality2
scenario = root.CurrentScenario
scenario2 = scenario.QueryInterface(STKObjects.IAgScenario)

# Store scenario specific variables
northFacName = 'Facility1'
southFacName = 'Facility2'
satName = 'Satellite1'
desiredMaxAngle = 5

# Grab hold of object references and paths
sat = scenario.Children.Item(satName)
satellite = sat.QueryInterface(STKObjects.IAgSatellite)
fac1 = scenario.Children.Item(northFacName)
fac2 = scenario.Children.Item(southFacName)
satPath = sat.Path
northFacPath = fac1.Path
southFacPath = fac2.Path


# Create AWB References
satLocalPath = 'Satellite/' + satName
fac1LocalPath = 'Facility/' + northFacName
fac2LocalPath = 'Facility/' + southFacName

#### 
# Compute Access Intervals and pick desired southern access intervals
#### 

# Compute Accesses
northFacAccess = fac1.GetAccessToObject(satellite);
northFacAccess.ComputeAccess
southFacAccess = fac2.GetAccessToObject(satellite);
southFacAccess.ComputeAccess;

# Pick the Right Southern Access Intervals based off of elevation angles
southAccess = southFacAccess.vgt.EventIntervalLists.Item('AccessIntervals')
southFacAccessIntervals = southAccess.FindIntervals()
accessTuple = ()
accessTuple = list(accessTuple)

for i in range(southFacAccess.ComputedAccessIntervalTimes.Count):
    startTime = southFacAccessIntervals.Intervals.Item(i).Start
    stopTime = southFacAccessIntervals.Intervals.Item(i).Stop
    aerData = southFacAccess.DataProviders.Item('AER Data').QueryInterface(STKObjects.IAgDataProviderGroup)
    aerData2 = aerData.Group.Item('Default')
    accessDP = aerData2.QueryInterface(STKObjects.IAgDataPrvTimeVar).Exec(startTime, stopTime, 60)
    elevationAngles = (accessDP.Intervals.Item(0).DataSets.GetDataSetByName('Elevation').GetValues())
    if (max(elevationAngles) >= desiredMaxAngle):
        accessTuple.append(startTime)
        accessTuple.append(stopTime)
        
accessTuple = tuple(accessTuple)


# Pull Orbit Start and Stop Times
passes = sat.DataProviders.Item('Passes').QueryInterface(STKObjects.IAgDataPrvInterval)
passesDP = passes.Exec(scenario2.StartTime, scenario2.StopTime)
passStartTimes = passesDP.DataSets.GetDataSetByName('Start Time').GetValues()
passEndTimes = passesDP.DataSets.GetDataSetByName('End Time').GetValues()

# Drill down to Time Tool for Satellite
timeToolIntervalLists = sat.Vgt.EventIntervalLists
timeToolIntervals = sat.Vgt.EventIntervals
timeToolEvent = sat.Vgt.Events


# Create Merged Interval List of North Passes and Access Intervals
if (timeToolIntervalLists.Contains('CheckingList') == False ): 
    checker = timeToolIntervalLists.Factory.CreateEventIntervalListMerged('CheckingList', 'Insert current pass and current access to see if merging them changes the pass interval')
    checker2 = checker.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListMerged)
    checker2.MergeOperation = 'eCrdnEventListMergeOperationMINUS'
    checkerIntervals = checker.FindIntervals()
else: 
    checker = timeToolIntervalLists.Item('CheckingList')
    checker2 = checker.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListMerged)
    checkerIntervals = checker.FindIntervals()


passes = timeToolIntervalLists.Item('PassIntervals');
passIntervals = passes.FindIntervals()
access = northFacAccess.vgt.EventIntervalLists.Item('AccessIntervals');
checker2.SetIntervalListA(passes)
checker2.SetIntervalListB(access)

# Create list of relevant orbits
if (timeToolIntervalLists.Contains('First_Chosen_Orbits') == False):
    finalOrbitList = timeToolIntervalLists.Factory.CreateEventIntervalListFixed('First_Chosen_Orbits','All orbits that meet criteria')
    finalOrbitList2 = finalOrbitList.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListFixed)
    finalOrbitListIntervals = finalOrbitList.FindIntervals()
else:
    finalOrbitList = timeToolIntervalLists.Item('First_Chosen_Orbits')
    finalOrbitList2 = finalOrbitList.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListFixed)
    finalOrbitListIntervals = finalOrbitList.FindIntervals()
 
timesTuple = {}
timesTuple = list(timesTuple)


## Take out orbits with North Accesses
for i in range(len(passStartTimes)):
    j = i
    startPoint = passIntervals.Intervals.Item(i).Start
    endPoint = passIntervals.Intervals.Item(i).Stop
    while (j < checkerIntervals.Intervals.Count) and (startPoint != checkerIntervals.Intervals.Item(j).Start): 
        j = j + 1
       
      
    if (j < checkerIntervals.Intervals.Count) and (endPoint == checkerIntervals.Intervals.Item(j).Stop): 
        timesTuple.append(startPoint)
        timesTuple.append(endPoint)
      

timesTuple = tuple(timesTuple)
finalOrbitList2.SetIntervals(timesTuple);


# Make Chosen Accesses an Interval List
if (timeToolIntervalLists.Contains('Chosen_Accesses') == False):
    chosenAccesses = timeToolIntervalLists.Factory.CreateEventIntervalListFixed('Chosen_Accesses','All accesses that meet criteria')
    chosenAccesses2 = chosenAccesses.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListFixed)
    chosenAccessesIntervals = chosenAccesses.FindIntervals()
else:
    chosenAccesses = timeToolIntervalLists.Item('Chosen_Accesses')
    chosenAccesses2 = chosenAccesses.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListFixed)
    chosenAccessesIntervals = chosenAccesses.FindIntervals()
 
chosenAccesses2.SetIntervals(accessTuple)


# Make an Intermediate, Checking Interval List
if (timeToolIntervalLists.Contains('CheckingList2') == False):
    checker3 = timeToolIntervalLists.Factory.CreateEventIntervalListMerged('CheckingList2', 'Insert current pass and elevation calculations to see if merging them changes the pass interval')
    checker4 = checker3.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListMerged)
    checker4.MergeOperation = 'eCrdnEventListMergeOperationMINUS'
    checker3Intervals = checker3.FindIntervals()
else:
    checker3 = timeToolIntervalLists.Item('CheckingList2')
    checker4 = checker3.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListMerged)
    checker3Intervals = checker3.FindIntervals()

checker4.SetIntervalListA(finalOrbitList)
checker4.SetIntervalListB(chosenAccesses)

# Make Ultimate List of selected orbits
if (timeToolIntervalLists.Contains('Ultimate_Chosen_Orbits') == False): 
    finalOrbitList3 = timeToolIntervalLists.Factory.CreateEventIntervalListFixed('Ultimate_Chosen_Orbits','All orbits that meet all criteria')
    finalOrbitList4 = finalOrbitList3.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListFixed)
    finalOrbitList3Intervals = finalOrbitList3.FindIntervals()
else:
    finalOrbitList3 = timeToolIntervalLists.Item('Ultimate_Chosen_Orbits')
    finalOrbitList4 = finalOrbitList3.QueryInterface(AgSTKVgtLib.IAgCrdnEventIntervalListFixed)
    finalOrbitList3Intervals = finalOrbitList3.FindIntervals()


orbitsTuple = list()
antiAccesslist = list()

for i in range(finalOrbitListIntervals.Intervals.Count):
      j = i
      startPoint = finalOrbitListIntervals.Intervals.Item(i).Start
      endPoint = finalOrbitListIntervals.Intervals.Item(i).Stop
      while (j < checker3Intervals.Intervals.Count) and (startPoint != checker3Intervals.Intervals.Item(j).Start): 
        j = j + 1;

      
      if (j < checker3Intervals.Intervals.Count) and (endPoint != checker3Intervals.Intervals.Item(j).Stop): 
          antiAccesslist.append(i)

          
j = 0

for i in range(finalOrbitListIntervals.Intervals.Count):
    if (j + 1 <= len(antiAccesslist) and i == antiAccesslist[j]):
      startPoint = finalOrbitListIntervals.Intervals.Item(i).Start
      endPoint = finalOrbitListIntervals.Intervals.Item(i).Stop
      orbitsTuple.append(startPoint)
      orbitsTuple.append(endPoint)
      j = j+1
 
orbitsTuple = tuple(orbitsTuple)
finalOrbitList4.SetIntervals(orbitsTuple)