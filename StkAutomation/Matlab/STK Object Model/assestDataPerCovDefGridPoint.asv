% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;
coverageDefinitionName = 'CoverageDefinition1';
assetName = 'Satellite1';
timestep = 86400;


% Add Facility
fac = scenario.Children.New('eFacility','Eval_Facility');
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
gridConstraintCommand = strcat('Cov */CoverageDefinition/',coverageDefinitionName," ",'Grid GridConstraint Facility AltAtTerrain Facility/Eval_Facility UseActualObject');   
root.ExecuteCommand(gridConstraintCommand);

% Add AWB Components
rangeVec = fac.Vgt.Vectors.Factory.CreateDisplacementVector('Range',facCenter,assetCenter);
rangeVec.Apparent = 1;

azimuthAngle = fac.Vgt.Angles.Factory.Create('Azimuth','Azimuth Angle','eCrdnAngleTypeDihedralAngle');
azimuthAngle.FromVector.SetVector(nedX);
azimuthAngle.ToVector.SetVector(rangeVec);
azimuthAngle.PoleAbout.SetVector(nedZ);

elevationAngle = fac.Vgt.Angles.Factory.Create('Elevation','Elevation Angle','eCrdnAngleTypeToPlane');
elevationAngle.ReferenceVector.SetVector(rangeVec);
elevationAngle.ReferencePlane.SetPlane(bodyPlane);

rangeCalc = fac.Vgt.CalcScalars.Factory.CreateCalcScalarVectorMagnitude('RangeValues', 'Range');
rangeCalc.InputVector = rangeVec;

azimuthCalc =  fac.Vgt.CalcScalars.Factory.CreateCalcScalarAngle('AzimuthValues','Azimuth');
azimuthCalc.InputAngle = azimuthAngle;

elevationCalc = fac.Vgt.CalcScalars.Factory.CreateCalcScalarAngle('ElevationValues','Elevation');
elevationCalc.InputAngle = elevationAngle;

% Add and Configure Figures of Merit
rangeFOM = covDef.Children.New('eFigureOfMerit','Range');
rangeFOM.SetScalarCalculationDefinition(rangeCalc.Path);
rangeFOM.Definition.TimeStep = 1;

azFOM = covDef.Children.New('eFigureOfMerit','Azimuth');
azFOM.SetScalarCalculationDefinition(azimuthCalc.Path);
azFOM.Definition.TimeStep = 1;

elFOM = covDef.Children.New('eFigureOfMerit','Elevation');
elFOM.SetScalarCalculationDefinition(elevationCalc.Path);
elFOM.Definition.TimeStep = 1;

% Compute Coverage Definition
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
end


