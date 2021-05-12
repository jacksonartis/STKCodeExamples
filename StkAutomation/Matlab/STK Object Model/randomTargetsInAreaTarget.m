% Start STK
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Grab Area Target and decide which objects to add
values = inputdlg({'Area Target Name','Number of Targets'},...
              'Customer', [1 30; 1 30]);
          
areaTarget = scenario.Children.Item(values{1});
numberOfTargets = round(str2double(values{2}));

if areaTarget.AccessConstraints.IsNamedConstraintActive('ElevationAngle')
       areaTarget.AccessConstraints.GetActiveNamedConstraint('ElevationAngle').Angle = 90;
else
       cons = areaTarget.AccessConstraints.AddNamedConstraint('ElevationAngle');
       cons.Angle = 90;
end

continueAns = questdlg('Add Facilities or Targets', ...
	'Choices', ...
	'Facility','Target','Target');
    switch continueAns
        case 'Facility'
            objectType = 'Facility';
            enumeration = 'eFacility';
        case 'Target'
            objectType = 'Target';
            enumeration = 'eTarget';
    end 


% Find Extents
points = cell2mat(areaTarget.AreaTypeData.ToArray);
lats = points(:,1);
longs = points(:,2);
maxLat = max(lats);
minLat = min(lats);
maxLong = max(longs);
minLong = min(longs);

% Add Objects
for i = 1:numberOfTargets
    latValue = minLat + (maxLat-minLat)*rand;
    longValue = minLong + (maxLong-minLong)*rand;
    name = strcat(objectType, string(i));
    newObj = scenario.Children.New(enumeration,name);
    newObj.Position.AssignGeocentric(latValue,longValue,1);
    access = areaTarget.GetAccessToObject(newObj);
    access.ComputeAccess;
    while ~access.ComputedAccessIntervalTimes.Count
        latValue = minLat + (maxLat-minLat)*rand;
        longValue = minLong + (maxLong-minLong)*rand;
        newObj.Position.AssignGeocentric(latValue,longValue,1);
        access.ComputeAccess;
    end
    access.ClearAccess;
end 

