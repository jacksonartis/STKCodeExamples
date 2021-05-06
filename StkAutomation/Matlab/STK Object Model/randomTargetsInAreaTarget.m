% Start STK
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Grab Area Target and decide which objects to add
values = inputdlg({'Area Target Name','Number of Targets'},...
              'Customer', [1 30; 1 30]);
          
areaTarget = scenario.Children.Item(values{1});
numberOfTargets = round(str2double(values{2}));

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
covDef = scenario.Children.New('eCoverageDefinition','Test');
covDef.Grid.BoundsType = 'eBoundsCustomBoundary';
covDef.Grid.Bounds.BoundaryObjects.AddObject(areaTarget);
covDef.Grid.Resolution.LatLon = 0.5;
pointDP = covDef.DataProviders.Item('Grid Point Locations').Exec;
lats = cell2mat(pointDP.DataSets.GetDataSetByName('Latitude').GetValues);
longs = cell2mat(pointDP.DataSets.GetDataSetByName('Longitude').GetValues);
covDef.Unload;
maxLat = max(lats);
minLat = min(lats);
maxLong = max(longs);
minLong = min(longs);

% Add Objects
for i = 1:numberOfTargets
    latValue = minLat + (maxLat-minLat)*rand;
    longValue = minLong + (maxLong-minLong)*rand;
    in = inpolygon(latValue, longValue, lats, longs);
    while ~in
        latValue = minLat + (maxLat-minLat)*rand;
        longValue = minLong + (maxLong-minLong)*rand;
        in = inpolygon(latValue, longValue, lats, longs);
    end
    name = strcat(objectType, string(i));
    newObj = scenario.Children.New(enumeration,name);
    newObj.Position.AssignGeocentric(latValue,longValue,0);
end 