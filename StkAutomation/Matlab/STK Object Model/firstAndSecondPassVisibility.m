% Connect to STK
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Grab sensor in question
sat = scenario.Children.Item('Satellite1');
sensor = sat.Children.Item('Sensor1');
visibility = sensor.Vgt.Volumes.Item('Visibility');

% Build first pass visibility
areaTarget = scenario.Children.New('eAreaTarget','First_Pass_Point');
sensorDP = sensor.DataProviders.Item('Pattern Intersection').Exec(scenario.StartTime, scenario.StopTime, 86400);
lats = cell2mat(sensorDP.DataSets.GetDataSetByName('Latitude').GetValues);
lon = cell2mat(sensorDP.DataSets.GetDataSetByName('Longitude').GetValues);
at = scenario.Children.New('eAreaTarget','AreaTarget2');

for i = 1:length(lats)
   areaTarget.AreaTypeData.Add(lats(i),lon(i)); 
end


% Calculate second pass visibility
firstPassGrid =  areaTarget.Vgt.VolumeGrids.Factory.CreateVolumeGridBearingAlt('firstPass','First Pass');
secondPassGrid = areaTarget.Vgt.VolumeGrids.Factory.CreateVolumeGridConstrained('secondPass','Second Pass');
secondPassGrid.ReferenceGrid = firstPassGrid;
secondPassGrid.Constraint = visibility;
volume = scenario.Children.New('eVolumetric', 'Combined_Visibility');
volume.VolumeGridDefinition.VolumeGrid = secondPassGrid.Path;
