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
intervals = access.Vgt.EventIntervalLists.Item('AccessIntervals');

intConstraint = fromObj.AccessConstraints.AddNamedConstraint('Intervals');


SetConstraint */Satellite/Sat283 Intervals Include TimeComponent "Satellite/Sat1 LightingIntervals.Sunlight Interval List"

parentPath = strcat(fromObj.ClassType(2:length(fromObj.ClassType)),'/',fromObj.InstanceName);
fullPath = strcat(fromObj.ClassType(2:length(fromObj.ClassType)),'/',fromObj.InstanceName, '/Sensor/limiter');

command = strcat(parentPath, " ", 'Intervals Include TimeComponent', " ", fullPath, intervals.QualifiedPath);
intConstraint.ExclIntvl = 1;
