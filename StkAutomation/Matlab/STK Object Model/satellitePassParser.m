%% Connect to STK and collect necessary references

% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Store scenario specific variables
northFacName = 'Facility1';
southFacName = 'Facility2';
satName = 'Satellite1';
desiredMaxAngle = 5;

% Grab hold of object references and paths
satellite = scenario.Children.Item(satName);
facility1 = scenario.Children.Item(northFacName);
facility2 = scenario.Children.Item(southFacName);

satPath = satellite.Path;
northFacPath = facility1.Path;
southFacPath = facility2.Path;

% Create AWB References
satLocalPath = strcat('Satellite/',satName);
fac1LocalPath = strcat('Facility/',northFacName);
fac2LocalPath = strcat('Facility/',southFacName);

%% Compute Access Intervals and pick the desired southern access intervals

% Compute Accesses
northFacAccess = facility1.GetAccessToObject(satellite);
northFacAccess.ComputeAccess
southFacAccess = facility2.GetAccessToObject(satellite);
southFacAccess.ComputeAccess;

% Pick the Right Southern Access Intervals based off of elevation angles
southAccess = southFacAccess.vgt.EventIntervalLists.Item('AccessIntervals');
accessTuple = {};
for i = 0:southAccess.FindIntervals.Intervals.Count-1
    startTime = southAccess.FindIntervals.Intervals.Item(i).Start;
    stopTime = southAccess.FindIntervals.Intervals.Item(i).Stop;
    accessDP = southFacAccess.DataProviders.Item('AER Data').Group.Item('Default').Exec(startTime, stopTime, 60);
    elevationAngles = abs(cell2mat(accessDP.Interval.Item(cast(0, 'int32')).DataSets.GetDataSetByName('Elevation').GetValues));
    if (max(elevationAngles) >= desiredMaxAngle)
        accessTuple = vertcat(accessTuple, startTime);
        accessTuple = vertcat(accessTuple, stopTime);
    end
end 
%% Create Time Components for North Facility

% Pull Orbit Start and Stop Times
passesDP = satellite.DataProviders.Item('Passes').Exec(scenario.StartTime, scenario.StopTime);
passStartTimes = passesDP.DataSets.GetDataSetByName('Start Time').GetValues;
passEndTimes = passesDP.DataSets.GetDataSetByName('End Time').GetValues;

% Drill down to Time Tool for Satellite
timeToolIntervalLists = satellite.vgt.EventIntervalLists;
timeToolIntervals = satellite.vgt.EventIntervals;
timeToolEvent = satellite.vgt.Events;

% Create Merged Interval List of North Passes and Access Intervals
if ~timeToolIntervalLists.Contains('CheckingList') 
    checker = timeToolIntervalLists.Factory.CreateEventIntervalListMerged('CheckingList', 'Insert current pass and current access to see if merging them changes the pass interval');
    checker.MergeOperation = 'eCrdnEventListMergeOperationMINUS';
else 
    checker = timeToolIntervalLists.Item('CheckingList');
end 
passes = timeToolIntervalLists.Item('PassIntervals');
access = northFacAccess.vgt.EventIntervalLists.Item('AccessIntervals');
checker.SetIntervalListA(passes);
checker.SetIntervalListB(access);


% Create list of relevant orbits
if ~timeToolIntervalLists.Contains('First_Chosen_Orbits') 
    finalOrbitList = timeToolIntervalLists.Factory.CreateEventIntervalListFixed('First_Chosen_Orbits','All orbits that meet criteria');
else
    finalOrbitList = timeToolIntervalLists.Item('First_Chosen_Orbits');
end 
timesTuple = {};



%% Take out orbits with North Accesses
  for i = 0:passes.FindIntervals.Intervals.Count-1
      j = i;
      startPoint = passes.FindIntervals.Intervals.Item(i).Start;
      endPoint = passes.FindIntervals.Intervals.Item(i).Stop;
      while (j < checker.FindIntervals.Intervals.Count) && (string(startPoint) ~= string(checker.FindIntervals.Intervals.Item(j).Start)) 
        j = j + 1;
      end 
      
      if (j < checker.FindIntervals.Intervals.Count) && (string(endPoint) == string(checker.FindIntervals.Intervals.Item(j).Stop)) 
          timesTuple = vertcat(timesTuple, startPoint);
          timesTuple = vertcat(timesTuple, endPoint);
      end
  end 
finalOrbitList.SetIntervals(timesTuple);

%% Create Time Components for South Facility

% Make Chosen Accesses an Interval List
if ~timeToolIntervalLists.Contains('Chosen_Accesses') 
    chosenAccesses = timeToolIntervalLists.Factory.CreateEventIntervalListFixed('Chosen_Accesses','All accesses that meet criteria');
else
    chosenAccesses = timeToolIntervalLists.Item('Chosen_Accesses');
end 
chosenAccesses.SetIntervals(accessTuple);

% Make an Intermediate, Checking Interval List
if ~timeToolIntervalLists.Contains('CheckingList2') 
    checker2 = timeToolIntervalLists.Factory.CreateEventIntervalListMerged('CheckingList2', 'Insert current pass and elevation calculations to see if merging them changes the pass interval');
    checker2.MergeOperation = 'eCrdnEventListMergeOperationMINUS';
else
    checker2 = timeToolIntervalLists.Item('CheckingList2');
end 
checker2.SetIntervalListA(finalOrbitList);
checker2.SetIntervalListB(chosenAccesses);

% Make Ultimate List of selected orbits
if ~timeToolIntervalLists.Contains('Ultimate_Chosen_Orbits') 
    finalOrbitList2 = timeToolIntervalLists.Factory.CreateEventIntervalListFixed('Ultimate_Chosen_Orbits','All orbits that meet all criteria');
else
    finalOrbitList2 = timeToolIntervalLists.Item('Ultimate_Chosen_Orbits');
end 
orbitsTuple = {};
antiAccesslist = [];

for i = 0:finalOrbitList.FindIntervals.Intervals.Count-1
      j = i;
      startPoint = finalOrbitList.FindIntervals.Intervals.Item(i).Start;
      endPoint = finalOrbitList.FindIntervals.Intervals.Item(i).Stop;
      while (j < checker2.FindIntervals.Intervals.Count) && (string(startPoint) ~= string(checker2.FindIntervals.Intervals.Item(j).Start)) 
        j = j + 1;
      end 
      
      if (j < checker2.FindIntervals.Intervals.Count) && (string(endPoint) ~= string(checker2.FindIntervals.Intervals.Item(j).Stop)) 
          antiAccesslist = horzcat(antiAccesslist, i)
          endPoint
      end
end

j = 1;
for i = 0:finalOrbitList.FindIntervals.Intervals.Count-1
    if (j <= length(antiAccesslist) && i == antiAccesslist(j))
      startPoint = finalOrbitList.FindIntervals.Intervals.Item(i).Start;
      endPoint = finalOrbitList.FindIntervals.Intervals.Item(i).Stop;
      orbitsTuple = vertcat(orbitsTuple, startPoint);
      orbitsTuple = vertcat(orbitsTuple, endPoint);
      j = j+1;
    end 
end 

finalOrbitList2.SetIntervals(orbitsTuple)