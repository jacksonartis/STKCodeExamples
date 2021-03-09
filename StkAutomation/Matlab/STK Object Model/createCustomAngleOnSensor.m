% Connect to STK
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Input Object Names
obj1Name = 'Satellite1';
obj2Name = 'Facility1';
obj3Name = 'Facility2';

% Get Object References
obj1 = scenario.Children.Item(obj1Name);
obj2 = scenario.Children.Item(obj2Name);
obj3 = scenario.Children.Item(obj3Name);

% Open AWB Factories
vectors = obj1.Vgt.Vectors.Factory;
angles = obj1.Vgt.Angles.Factory;
points = obj1.Vgt.Points;

% Get Center points
obj1Center = points.Item('Center');
obj2Center = obj2.Vgt.Points.Item('Center');
obj3Center = obj3.Vgt.Points.Item('Center');

% Create new Geometric components
vector1 = vectors.CreateDisplacementVector(strcat('To',obj2Name),obj1Center,obj2Center);
vector2 = vectors.CreateDisplacementVector(strcat('To',obj3Name),obj1Center,obj3Center);
newAngle = angles.Create('AngleBetweenVectors','Angle between to displacement vectors', 0);
newAngle.FromVector.SetPath(vector1.Path);
newAngle.toVector.SetPath(vector2.Path);








