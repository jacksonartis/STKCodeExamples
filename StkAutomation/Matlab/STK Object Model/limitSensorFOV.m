% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Variables
answers = inputdlg({'Parent Object Name','Sensor Name'},...
              'Customer', [1 30; 1 30]);
satName = string(answers{1});
sensorName = string(answers{2});
sat = scenario.Children.Item(satName);
sensor = sat.Children.Item(sensorName);
whichAngle = questdlg('Which Half Angle would you like to be non-uniform?', ...
	'Choices', ...
	'Horizontal','Vertical','No');
switch whichAngle
        case 'Horizontal'
            axis = 'Y';
        case 'Vertical'
            axis = 'X';
end 
values = inputdlg({'Half Angle 1','Half Angle 2'},...
              'Customer', [1 30; 1 30]);
max_angle = max(str2double(values));
min_angle = -min(str2double(values));

% Get Handles
keepGoing = true;
i = 1;
while keepGoing
    answers = inputdlg({'To Access Object Name'},...
                  'Customer', [1 30]);
    objName = string(answers{1});
    obj = scenario.Children.Item(objName);

    % Create Angle
    vectorName = strcat('to',objName);
    vector = sensor.Vgt.Vector.Factory.CreateDisplacementVector(vectorName, sensor.Vgt.Point.Item('Center'), obj.Vgt.Point.Item('Center'));
    angleName = strcat('ConstrainingAngle',string(i));
    angle = sensor.Vgt.Angle.Factory.Create(angleName, '', 'eCrdnAngleTypeDihedralAngle');
    angle.PoleAbout.SetVector(sensor.Vgt.Vector.Item(strcat('Body.',axis)));
    angle.ToVector.SetVector(vector);
    angle.SignedAngle = 1;
    anglePath = angle.Path;

    % Add Constraints
    angleConstraint = sensor.AccessConstraints.AddConstraint('eCstrCrdnAngle');
    angleConstraint.EnableMin = 1;
    angleConstraint.EnableMax = 1;
    angleConstraint.Max = max_angle;
    angleConstraint.Min = min_angle;

    angleConstraint.Reference = anglePath;
    i = i + 1;
    continueAns = questdlg('Add More Objects?', ...
	'Choices', ...
	'Yes','No','No');
    switch continueAns
        case 'Yes'
            keepGoing = true;
        case 'No'
            keepGoing = false;
    end 
end 



