% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

values = inputdlg({'From Object','To Object'},...
              'Customer', [1 30; 1 30]);
facilityName = values{1};
assetName = values{2};
timestep = 86400;

fac = scenario.Children.Item(facilityName);
asset = scenario.Children.Item(assetName);

facCenter = fac.Vgt.Points.Item('Center');
nedX = fac.Vgt.Vectors.Item('NorthEastDown.X');
nedZ = fac.Vgt.Vectors.Item('NorthEastDown.Z');
bodyPlane = fac.Vgt.Planes.Item('Body.XY');

% Get Asset
asset = scenario.Children.Item(assetName);
assetCenter = asset.Vgt.Points.Item('Center');
startTime = asset.Vgt.Events.Item('AvailabilityStartTime').ReferenceEventInterval.FindInterval.Interval.Start;
stopTime = asset.Vgt.Events.Item('AvailabilityStopTime').ReferenceEventInterval.FindInterval.Interval.Stop;

% Add Vector
if fac.Vgt.Vectors.Contains('Range')
    rangeVec = fac.Vgt.Vectors.Item('Range');
else
    rangeVec = fac.Vgt.Vectors.Factory.CreateDisplacementVector('Range',facCenter,assetCenter);
end
rangeVec.Apparent = 1;

% Add Angles
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


% Add Calc Scalars
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