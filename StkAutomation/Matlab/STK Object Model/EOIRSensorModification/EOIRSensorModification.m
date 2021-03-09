% Access STK
app = actxGetRunningServer('STK12.application');
root = app.Personality2;

% Get Scenario
scenario = root.CurrentScenario;
satelliteList = scenario.Children.GetElements('eSatellite');

%% Organize values

s = struct; % s.a - spatial, s.b - spectral, s.c - optical, s.d - radiometric


% Spatial
spat_strings = string({'eNumPixAndPixelPitch'}); % alternates are eFOVandPixelPitch and eFOVandNumPix
spat_numbers = [2.175, 2.175, 128, 128, 65.276, 65.276, 0.5934, 0.5934]; % [FOV_Hor, FOV_Ver, NOP_Hor, NOP_Vert, RDP_Hor, RDP_Vert, IFOV_Hor, IFOV_Vert]
if spat_strings(1) == string('eNumPixAndPixelPitch')
    s.a = {spat_numbers(3:6)};
    
elseif spat_strings(1) == string('eFOVandPixelPitch')
    s.a = {spat_numbers(1:2),spat_numbers(5:6)}; 
    
else 
    s.a = {spat_numbers(1:4)};
    
end

% Spectral
spectral_shape = string('eDefault'); % alternate is 'eProvideRSR'
spec_numbers = [0.4, 0.7, 6]; % [low, high, number of intervals]

if spectral_shape == string('eProvideRSR')
    RSRUnits = 'eEnergyUnits'; % alternate is 'eQuantaUnits'
    spectral_response_path = "C:\Program Files\AGI\STK 12\EOIR_Databases\PropertyFiles\RSR_Files\Pan_RSR.srf";
end

s.b = spec_numbers;


% Optical
opt_strings = string({'eFNumberAndFocalLength', 'eDiffractionLimited', 'eBandEffectiveTransmission', 'eBandCenter'});
% alternates for index 1: eFNumberAndApertureDiameter and eFocalLengthAndApertureDiameter

% alternates for index 2: eNegligibleAberrations, eMildAberrations,
% eModerateAberrations, eCustomWavefrontError, eCustomPupilFunction,
% eCustomPSF, eCustomMTF

% alternates for index 3: eTransmissionDataFile

% alternates for index 4: eLowBandEdge, eHighBandEdge,
% eUserDefinedWavelength

opt_numbers = [2, 11, 5.5, 0, 5, 5, 0, 1, 0.55]; % [F/#, effectivefocallength, entrancePupil, RMS_wavefront, file_sampling, file_sampling_freq, longitudinalDefocus, band effective value, wavelength]
chosen_numbers = [];
dataFile = "C:\Program Files\AGI\STK 12\EOIR_Databases\PropertyFiles\Shape_Files\2dGaussian4StdDev64x64.csv";
spectralTransmission = "C:\Program Files\AGI\STK 12\EOIR_Databases\PropertyFiles\Optical_Transmission_Files\ARGlass_Trans.srf";



if opt_strings(1) == string('eFNumberAndFocalLength')
    chosen_numbers = horzcat(chosen_numbers, opt_numbers(1:2));
elseif opt_strings(1) == string('eFNumberAndApertureDiameter')
    chosen_numbers = horzcat(chosen_numbers, opt_numbers(1), opt_numbers(3));
else
    chosen_numbers = horzcat(chosen_numbers, opt_numbers(2:3));
end 



if opt_strings(2) == string('eCustomMTF')
    chosen_numbers = horzcat(chosen_numbers, opt_numbers(6));
elseif opt_strings(2) == string('eCustomPSF')
    chosen_numbers = horzcat(chosen_numbers, opt_numbers(5));
elseif opt_strings(2) == string('eCustomWavefrontError')
    chosen_numbers = horzcat(chosen_numbers, opt_numbers(4), opt_numbers(7));
else
    chosen_numbers = horzcat(chosen_numbers, opt_numbers(7));
end 

 if opt_strings(3) == string('eBandEffectiveTransmission')
       chosen_numbers = horzcat(chosen_numbers, opt_numbers(8));
 end


if opt_strings(4) == string('eUserDefinedWavelength')
        chosen_numbers = horzcat(chosen_numbers, opt_numbers(9));
end 

s.c = chosen_numbers;

% Radiometric
radio_number = 100;


%% Change Values

% Move through each satellite
for i = 1:satelliteList.Count
    currentSat = satelliteList.Item(int32(i-1));
    sensorList = currentSat.Children.GetElements('eSensor');
    
    % Move through each sensor
    for j = 1:sensorList.Count
       currentSensor = sensorList.Item(int32(j-1));
       pattern = currentSensor.Pattern;
       
       % Move through each band
       for k = 1:pattern.Bands.Count
          currentBand = pattern.Bands.Item(int32(k-1));
          setSpatial(cell2mat(s.a), spat_strings, currentBand);
          setSpectral(s.b, spectral_shape, currentBand)
          setOptic(s.c, opt_strings, currentBand, spectralTransmission, dataFile)
          setRadiometric(radio_number, currentBand)
       end
    end 
end
