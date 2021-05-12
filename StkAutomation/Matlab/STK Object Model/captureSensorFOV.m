% Connect to STK
app = actxGetRunningServer("STK12.Application");
root = app.Personality2;
scenario = root.CurrentScenario;

% Get
sat = scenario.Children.Item('Satellite1');
sensor = sat.Children.Item('Sensor1');
sensorDP = sensor.DataProviders.Item('Pattern Intersection').Exec(scenario.StartTime, '10 May 2021 16:01:00.000', 1);
sensorDP = sensorDP.Intervals.Item(int32(0));

lats = cell2mat(sensorDP.DataSets.GetDataSetByName('Latitude').GetValues);
lon = cell2mat(sensorDP.DataSets.GetDataSetByName('Longitude').GetValues);
at = scenario.Children.New('eAreaTarget','AreaTarget2');

for i = 1:length(lats)
   at.AreaTypeData.Add(lats(i),lon(i)); 
end
