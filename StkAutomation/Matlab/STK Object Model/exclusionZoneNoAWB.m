% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

values = inputdlg({'From Object','To Object','Exclusion Zone (deg)'},...
              'Customer', [1 30; 1 30; 1 30]);
          
fromObj = scenario.Children.Item(string(values{1}));
toObj = scenario.Children.Item(string(values{2}));
newSensor = fromObj.Children.New('eSensor','limiter');
newSensor.Pattern.ConeAngle = str2double(values{3});

access = newSensor.GetAccessToObject(toObj);
aerDP = access.DataProviders.Item('Access Data').Exec(scenario.StartTime, scenario.StopTime);
startTimes = string(cell2mat(aerDP.DataSets.GetDataSetByName('Start Time').GetValues));
stopTimes = string(cell2mat(aerDP.DataSets.GetDataSetByName('Stop Time').GetValues));
intConstraint = fromObj.AccessConstraints.AddNamedConstraint('Intervals');

for i = 1:length(startTimes)
    intConstraint.Intervals.Add(startTimes(i),stopTimes(i));
end

intConstraint.ExclIntvl = 1;