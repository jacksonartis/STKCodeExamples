% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;
constellationName = 'Constellation1';
vertAngle = 20;
horzAngle = 20;

% Connect to Constellation
const = scenario.Children.Item(constellationName);

for i = 1:const.Objects.Count
    currentObj = const.Objects.Item(int32(i-1)).LinkedObject;
    currentObj.Pattern.HorizontalHalfAngle = horzAngle;
    currentObj.Pattern.VerticalHalfAngle = vertAngle;
    solarExclusion = currentObj.AccessConstraints.AddConstraint('eCstrLOSSunExclusion');
    solarExclusion.Angle = 45;
end
