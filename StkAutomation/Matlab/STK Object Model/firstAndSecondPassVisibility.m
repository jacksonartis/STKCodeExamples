% Connect to STK
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Grab sensor in question
sat = scenario.Children.Item('Satellite111');
sensor = sat.Children.Item('Sensor1');
visibility = sensor.Vgt.Volumes.Item('Visibility');

% Build first pass visibility
areaTarget = scenario.Children.New('eAreaTarget','First_Pass_Point');
sensorDP = sensor.DataProviders.Item('Pattern Intersection').Exec(scenario.StartTime, scenario.StopTime, 86400);
lats = cell2mat(sensorDP.DataSets.GetDataSetByName('Latitude').GetValues);
lon = cell2mat(sensorDP.DataSets.GetDataSetByName('Longitude').GetValues);

% Second boundary
secondBoundDP = sensorDP.Sections.Item(int32(1)).Intervals.Item(int32(0));
lats2 = cell2mat(secondBoundDP.DataSets.GetDataSetByName('Latitude').GetValues);
lon2 = cell2mat(secondBoundDP.DataSets.GetDataSetByName('Longitude').GetValues);

for i = 1:length(lats)
   areaTarget.AreaTypeData.Add(lats(i),lon(i)); 
end

for i = 1:length(lats2)
   areaTarget.AreaTypeData.Add(lats2(i),lon2(i)); 
end



% Calculate second pass visibility
firstPassGrid =  areaTarget.Vgt.VolumeGrids.Factory.CreateVolumeGridBearingAlt('firstPass','First Pass');
secondPassGrid = areaTarget.Vgt.VolumeGrids.Factory.CreateVolumeGridConstrained('secondPass','Second Pass');
secondPassGrid.ReferenceGrid = firstPassGrid;
secondPassGrid.Constraint = visibility;

secondPassInt = sensor.Vgt.EventIntervals.Factory.CreateEventIntervalFixed('SecondPass','Second Pass');
secondPassInt.SetInterval(scenario.StartTime, scenario.StopTime);
fixedIntRef = strcat(secondPassInt.Path, " ", 'Interval');

volume = scenario.Children.New('eVolumetric', 'Combined_Visibility');
volume.VolumeGridDefinition.VolumeGrid = secondPassGrid.Path;
volume.VolumeAnalysisInterval.AnalysisInterval = fixedIntRef;

