% Connect to STK
app = actxGetRunningServer("STK12.Application");
root = app.Personality2;
scenario = root.CurrentScenario;

% Get
sat = scenario.Children.Item('Satellite111');
sensor = sat.Children.Item('Sensor1');
sensorDP = sensor.DataProviders.Item('Pattern Intersection').Exec(scenario.StartTime, scenario.StopTime, 3600);


lats = cell2mat(sensorDP.DataSets.GetDataSetByName('Latitude').GetValues);
lon = cell2mat(sensorDP.DataSets.GetDataSetByName('Longitude').GetValues);
at = scenario.Children.New('eAreaTarget','AreaTarget2');

for i = 1:length(lats)
   at.AreaTypeData.Add(lats(i),lon(i)); 
end
