% Grab STK Instance
app = actxGetRunningServer('STK11.Application');
root = app.Personality2;
scenario = root.CurrentScenario;
coverageDefinitionName = 'CoverageDefinition1';
facilityName = 'Eval_Facility2';
assetName = 'TAP1';
timestep = 86400;


% Add Facility
if scenario.Children.Contains('eFacility', facilityName)
    fac = scenario.Children.Item(facilityName);
else
    fac = scenario.Children.New('eFacility',facilityName);
end
facCenter = fac.Vgt.Points.Item('Center');
nedX = fac.Vgt.Vectors.Item('NorthEastDown.X');
nedZ = fac.Vgt.Vectors.Item('NorthEastDown.Z');
bodyPlane = fac.Vgt.Planes.Item('Body.XY');

% Get Asset
asset = scenario.Children.Item(assetName);
assetCenter = asset.Vgt.Points.Item('Center');
startTime = asset.Vgt.Events.Item('AvailabilityStartTime').ReferenceEventInterval.FindInterval.Interval.Start;
stopTime = asset.Vgt.Events.Item('AvailabilityStopTime').ReferenceEventInterval.FindInterval.Interval.Stop;

% Configure Connect Command
covDef = scenario.Children.Item(coverageDefinitionName);
gridConstraintCommand = strcat('Cov */CoverageDefinition/',coverageDefinitionName," ",'Grid GridConstraint Facility AltAtTerrain Facility/', fac.InstanceName, " ", 'UseActualObject');   
root.ExecuteCommand(gridConstraintCommand);

% Add AWB Components
if fac.Vgt.Vectors.Contains('Range')
    rangeVec = fac.Vgt.Vectors.Item('Range');
else
    rangeVec = fac.Vgt.Vectors.Factory.CreateDisplacementVector('Range',facCenter,assetCenter);
end
rangeVec.Apparent = 1;

if fac.Vgt.Angles.Contains('Azimuth')
    azimuthAngle = fac.Vgt.Angles.Item('Azimuth');
else
    azimuthAngle = fac.Vgt.Angles.Factory.Create('Azimuth','Azimuth Angle','eCrdnAngleTypeDihedralAngle');
end
azimuthAngle.FromVector.SetVector(nedX);
azimuthAngle.ToVector.SetVector(rangeVec);
azimuthAngle.PoleAbout.SetVector(nedZ);

if fac.Vgt.Angles.Contains('Elevation')
    elevationAngle = fac.Vgt.Angles.Item('Elevation');
else
    elevationAngle = fac.Vgt.Angles.Factory.Create('Elevation','Elevation Angle','eCrdnAngleTypeToPlane');
end
elevationAngle.ReferenceVector.SetVector(rangeVec);
elevationAngle.ReferencePlane.SetPlane(bodyPlane);

if fac.Vgt.CalcScalars.Contains('RangeValues')
    rangeCalc = fac.Vgt.CalcScalars.Item('RangeValues');
else
    rangeCalc = fac.Vgt.CalcScalars.Factory.CreateCalcScalarVectorMagnitude('RangeValues', 'Range');
end
rangeCalc.InputVector = rangeVec;

if fac.Vgt.CalcScalars.Contains('AzimuthValues')
    azimuthCalc = fac.Vgt.CalcScalars.Item('AzimuthValues');
else
    azimuthCalc =  fac.Vgt.CalcScalars.Factory.CreateCalcScalarAngle('AzimuthValues','Azimuth');
end
azimuthCalc.InputAngle = azimuthAngle;

if fac.Vgt.CalcScalars.Contains('ElevationValues')
    elevationCalc = fac.Vgt.CalcScalars.Item('ElevationValues');
else
    elevationCalc =  fac.Vgt.CalcScalars.Factory.CreateCalcScalarAngle('ElevationValues','Elevation');
end
elevationCalc.InputAngle = elevationAngle;

% Add Rate Scalar
if fac.Vgt.CalcScalars.Contains('RangeRate')
    rangeRate = fac.Vgt.CalcScalars.Item('RangeRate');
else
   rangeRate = fac.Vgt.CalcScalars.Factory.CreateCalcScalarDerivative('RangeRate','Rate of change in range');
end
rangeRate.Scalar = rangeCalc;

if fac.Vgt.CalcScalars.Contains('AzimuthRate')
    azRate = fac.Vgt.CalcScalars.Item('AzimuthRate');
else
   azRate = fac.Vgt.CalcScalars.Factory.CreateCalcScalarDerivative('AzimuthRate','Rate of change in azimuth');
end
azRate.Scalar = azimuthCalc;

if fac.Vgt.CalcScalars.Contains('ElevationRate')
   elRate = fac.Vgt.CalcScalars.Item('ElevationRate');
else
   elRate = fac.Vgt.CalcScalars.Factory.CreateCalcScalarDerivative('ElevationRate','Rate of change in elevation');
end
elRate.Scalar = elevationCalc;

% Add and Configure Figures of Merit
if covDef.Children.Contains('eFigureOfMerit', 'Range')
    rangeFOM = covDef.Children.Item('Range');
else
    rangeFOM = covDef.Children.New('eFigureOfMerit','Range');
end
rangeFOM.SetScalarCalculationDefinition(rangeCalc.Path);
rangeFOM.Definition.TimeStep = 1;

if covDef.Children.Contains('eFigureOfMerit', 'RangeRate')
    rangeRateFOM = covDef.Children.Item('RangeRate');
else
    rangeRateFOM = covDef.Children.New('eFigureOfMerit','RangeRate');
end
rangeRateFOM.SetScalarCalculationDefinition(rangeRate.Path);
rangeRateFOM.Definition.TimeStep = 1;

if covDef.Children.Contains('eFigureOfMerit', 'Azimuth')
    azFOM = covDef.Children.Item('Azimuth');
else
    azFOM = covDef.Children.New('eFigureOfMerit','Azimuth');
end
azFOM.SetScalarCalculationDefinition(azimuthCalc.Path);
azFOM.Definition.TimeStep = 1;

if covDef.Children.Contains('eFigureOfMerit', 'AzimuthRate')
    azRateFOM = covDef.Children.Item('AzimuthRate');
else
    azRateFOM = covDef.Children.New('eFigureOfMerit','AzimuthRate');
end
azRateFOM.SetScalarCalculationDefinition(azRate.Path);
azRateFOM.Definition.TimeStep = 1;

if covDef.Children.Contains('eFigureOfMerit', 'Elevation')
    elFOM = covDef.Children.Item('Elevation');
else
    elFOM = covDef.Children.New('eFigureOfMerit','Elevation');
end
elFOM.SetScalarCalculationDefinition(elevationCalc.Path);
elFOM.Definition.TimeStep = 1;

if covDef.Children.Contains('eFigureOfMerit', 'ElevationRate')
    elRateFOM = covDef.Children.Item('ElevationRate');
else
    elRateFOM = covDef.Children.New('eFigureOfMerit','ElevationRate');
end
elRateFOM.SetScalarCalculationDefinition(elRate.Path);
elRateFOM.Definition.TimeStep = 1;



% Compute Coverage Definition
covDef.AssetList.Add(asset.Path)
covDef.ComputeAccesses;

% Get Timesteps
timeDP = asset.DataProviders.Item('Classical Elements').Group.Item('ICRF').Exec(startTime, stopTime, timestep);
timesteps = string(timeDP.DataSets.GetDataSetByName('Time').GetValues);

rangeDP = rangeFOM.DataProviders.Item('Time Value By Point');
rangeDP.PreData = timesteps(1);
lat = cell2mat(rangeDP.Exec.DataSets.GetDataSetByName('Latitude').GetValues);
lon = cell2mat(rangeDP.Exec.DataSets.GetDataSetByName('Longitude').GetValues);

% Pull Grid Point Data for Each Time
for i = 1:length(timesteps)
    rangeDP = rangeFOM.DataProviders.Item('Time Value By Point');
    rangeDP.PreData = timesteps(i);
    rangeValues = cell2mat(rangeDP.Exec.DataSets.GetDataSetByName('FOM Value').GetValues);
    
    azDP = azFOM.DataProviders.Item('Time Value By Point');
    azDP.PreData = timesteps(i);
    azValues = cell2mat(azDP.Exec.DataSets.GetDataSetByName('FOM Value').GetValues);
    
    elDP = elFOM.DataProviders.Item('Time Value By Point');
    elDP.PreData = timesteps(i);
    elValues = cell2mat(elDP.Exec.DataSets.GetDataSetByName('FOM Value').GetValues);
    
    rangeRateDP = rangeRateFOM.DataProviders.Item('Time Value By Point');
    rangeRateDP.PreData = timesteps(i);
    rangeRateValues = cell2mat(rangeRateDP.Exec.DataSets.GetDataSetByName('FOM Value').GetValues);
    
    azRateDP = azRateFOM.DataProviders.Item('Time Value By Point');
    azRateDP.PreData = timesteps(i);
    azRateValues = cell2mat(azRateDP.Exec.DataSets.GetDataSetByName('FOM Value').GetValues);
    
    elRateDP = elRateFOM.DataProviders.Item('Time Value By Point');
    elRateDP.PreData = timesteps(i);
    elRateValues = cell2mat(elRateDP.Exec.DataSets.GetDataSetByName('FOM Value').GetValues);
end


