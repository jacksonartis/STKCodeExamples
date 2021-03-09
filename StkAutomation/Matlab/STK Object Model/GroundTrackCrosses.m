% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% List scenario specific variables
satName = 'Satellite2';
scalarValue = root.ConversionUtility.NewQuantity('AngleRate' , 'deg/sec', 0);

% Get hold of satellite and create variables for necessary AWB sections
satellite = scenario.Children.Item(satName);

intervals = satellite.Vgt.EventIntervals;
intervalLists = satellite.Vgt.EventIntervalLists;

scalars = satellite.Vgt.CalcScalars;
conditions = satellite.Vgt.Conditions;

% Create Calculation Components
if scalars.Contains('Lat_Rate')
    lat_rate = scalars.Item('Lat_Rate');
else
lat_rate = scalars.Factory.CreateCalcScalarDataElementWithGroup('Lat_Rate', 'Follows Latitude rate', 'LLA State', 'Fixed', 'Lat Rate');
end 

if conditions.Contains('Increasing')
    increasing = conditions.Item('Increasing');
else 
    increasing = conditions.Factory.CreateConditionScalarBounds('Increasing', 'Rising Ground track');
end

if conditions.Contains('Decreasing')
    decreasing = conditions.Item('Decreasing');
else 
    decreasing = conditions.Factory.CreateConditionScalarBounds('Decreasing', 'Falling Ground track');
end

increasing.Scalar = lat_rate;
decreasing.Scalar = lat_rate;

increasing.Operation = 'eCrdnConditionThresholdOptionAboveMin';
decreasing.Operation = 'eCrdnConditionThresholdOptionBelowMax';

increasing.SetMinimum(scalarValue);
decreasing.SetMaximum(scalarValue);

% Create Time Components
passes = intervalLists.Item('PassIntervals');

if intervals.Contains('Current_Pass')
    currentPass = intervals.Item('Current_Pass');
else 
    currentPass = intervals.Factory.CreateEventIntervalFromIntervalList('Current_Pass','Pass X');
end

if intervals.Contains('Following_Pass')
    nextPass = intervals.Item('Following_Pass');
else 
    nextPass = intervals.Factory.CreateEventIntervalFromIntervalList('Following_Pass','Pass X+1');
end

currentPass.ReferenceIntervals = passes;
nextPass.ReferenceIntervals = passes;

currentPass.IntervalSelection = 'eCrdnIntervalSelectionFromStart';
nextPass.IntervalSelection = 'eCrdnIntervalSelectionFromStart';

if intervalLists.Contains('Before_Ascending_Node')
    beforeNode = intervalLists.Item('Before_Ascending_Node');
else
    beforeNode = intervalLists.Factory.CreateEventIntervalListCondition('Before_Ascending_Node', 'when the satellite is ascending in ground track');
end

if intervalLists.Contains('Ascending_Times')
    whenAscending = intervalLists.Item('Ascending_Times');
else 
    whenAscending = intervalLists.Factory.CreateEventIntervalListCondition('Ascending_Times', 'when the satellite is ascending in ground track');
end

if intervalLists.Contains('Descending_Times')
    whenDescending = intervalLists.Item('Descending_Times');
else
    whenDescending = intervalLists.Factory.CreateEventIntervalListCondition('Descending_Times', 'when the satellite is descending in ground track');
end

beforeNode.Condition = conditions.Item('AboveAscendingNode');
whenAscending.Condition = increasing;
whenDescending.Condition = decreasing;

% Create First Merged Interval Lists
if intervalLists.Contains('Upward_Motion')
    upward = intervalLists.Item('Upward_Motion');
else 
    upward = intervalLists.Factory.CreateEventIntervalListMerged('Upward_Motion', 'Ascending before ascending node');
end

if intervalLists.Contains('Downward_Motion')
    downward = intervalLists.Item('Downward_Motion');
else 
    downward = intervalLists.Factory.CreateEventIntervalListMerged('Downward_Motion', 'Descending before ascending node');
end

% Set first merged list
upward.SetIntervalListA(whenAscending);
upward.SetIntervalListB(beforeNode);
downward.SetIntervalListA(whenDescending);
downward.SetIntervalListB(beforeNode);

% Create Second set of Merged Lists
if intervalLists.Contains('This_Orbit_Rising')
    upward2 = intervalLists.Item('This_Orbit_Rising');
else 
    upward2 = intervalLists.Factory.CreateEventIntervalListMerged('This_Orbit_Rising', 'Rising times for orbit X');
end 

if intervalLists.Contains('Next_Orbit_Falling')
    downward2 = intervalLists.Item('Next_Orbit_Falling');
else
    downward2 = intervalLists.Factory.CreateEventIntervalListMerged('Next_Orbit_Falling', 'Rising times for orbit X + 1');
end

% Set second merged lists
upward2.SetIntervalA(currentPass);
upward2.SetIntervalListB(upward);
downward2.SetIntervalA(nextPass);
downward2.SetIntervalListB(downward)

currentPass.IntervalNumber = 1;
nextPass.IntervalNumber = 1 + 1;

currentStart = upward2.FindIntervals.Interval.Item(0).Start;
currentStop = upward2.FindIntervals.Interval.Item(0).Stop;

nextStart = downward2.FindIntervals.Interval.Item(0).Start;
nextStop = downward2.FindIntervals.Interval.Item(0).Stop;

%% Get output values
currentDP = satellite.DataProviders.GetDataPrvTimeVarFromPath('LLA State/Fixed').Exec(currentStart, currentStop, 1);
nextDP = satellite.DataProviders.GetDataPrvTimeVarFromPath('LLA State/Fixed').Exec(nextStart, nextStop, 1);


currentLatIncreasing = cell2mat(currentDP.DataSets.GetDataSetByName('Lat').GetValues);
currentLon = cell2mat(currentDP.DataSets.GetDataSetByName('Lon').GetValues);

nextLatDecreasing = cell2mat(nextDP.DataSets.GetDataSetByName('Lat').GetValues);
nextLon = cell2mat(nextDP.DataSets.GetDataSetByName('Lon').GetValues);
nextTime = string(cell2mat(nextDP.DataSets.GetDataSetByName('Time').GetValues));

maxInc = fix(length(currentLatIncreasing)/3)*2;

matchFound = false; 
i = 1;
relevantIndex = 1;
tester = 1;
while ~matchFound && i < maxInc
   lat = nextLatDecreasing(i);
   lon = nextLon(i);
   values = lat - currentLatIncreasing;
   values(values < 0) = [];
   j = length(values);
   if (lon - currentLon(j) < 0.1) && (lon - currentLon(j) > 0)
       relevantIndex = i;
       matchFound = true;
       tester = j;
   end
   i = i+1;
end

timeOfCross = nextTime(relevantIndex);
latOfCross = nextLatDecreasing(relevantIndex);
lonOfCross = nextLon(relevantIndex);


%% Set facility
if scenario.Children.GetElements('eFacility').Contains('crossPoint')
    crossPoint = scenario.Children.Item('crossPoint');
else 
    crossPoint = scenario.Children.New('eFacility', 'crossPoint');
end
% IAgFacility facility: Facility Object
crossPoint.Position.AssignGeodetic(latOfCross, lonOfCross, 0);

% Accuracy of ~.28km











