close all;
clear all;
clc;

%% Launch STK

%Create an instance of STK
uiApplication = actxserver('STK12.application');
uiApplication.Visible = 1;

%Get our IAgStkObjectRoot interface
root = uiApplication.Personality2;

Scenario = root.Children.New('eScenario', 'Generic_Aircraft');

% Set Start time to Now and Stop to 1 Day later
%Get the system clock time and use that to set up the scenario's start and
%stop time.
tomorrow_date = datestr((now+1), 'dd mmm yyyy HH:MM:SS.FFF')
current_date = datestr((now), 'dd mmm yyyy HH:MM:SS.FFF')

scenObj.Epoch = current_date
scenObj.StopTime = tomorrow_date
scenObj.StartTime = current_date

%% Create Aircraft
aircraft = Scenario.Children.New('eAircraft', 'Aircraft1')

%% Create and Modify Radar Cross Section of Aircraft

% Define Radar Cross Section Handle of Aircraft 1 to be "RCS"
RCS = aircraft.RadarCrossSection;

% Set RCS to not be Inherit - Allows us to modify
RCS.Inherit = 0;    % VERY IMPORTANT PIECE OF CODE!!!!

%% Inspect and Alter the Frequency Bands

% Define RCS's frequency Bands to be "RCS_FBands"
RCS_FBands = RCS.Model.FrequencyBands;

% Returns the number of Frequency elements in the collection
Number_of_Bands = RCS_FBands.Count

% Add a Second Band - Indicate Min/Max Frequencies
RCS_FBands.Add(3.5, 4.5) % Frequency values in MHz -- Add(Min_Frequency, Max_Frequency)

%% Inspect and Change Individual Bands

% In STK a Radar Cross Section is determined across a Frequency band of
% 2.99 MHz to 3.0e+11 MHz. There will always be at least 1 band, and if
% there is only 1 band it's Minimum Frequency and Maximum Frequency can
% ONLY be set to these values. When you add a New Band, you need to
% indicate the Minimum and Maximum Frequencies of this New Band. By adding
% 1 New Band you have now cut the Frequency spectrum into three chunks or "Bands". The
% first band is [Absolute Min (2.99 MHz), Min of Band 2], the second band
% is [Min of Band 2, Max of Band 2] and the third band is [Max of Band 2,
% Absolute Max (3.0e+11 MHz)]. It is important to note that by creating 1
% new band you now have 3 bands to alter not just 2. If you add 2 bands you
% will have 4 to alter not 3, etc.
%         ** Note [x,y] means x = Min Frequency and y = Max Frequency **

% Once you have added a New Band and defined its Min and Max values you can
% view these Min/Max Frequencies but you can only change the Minimum Value!
% This is the same as in the GUI, only the minimum value is open for
% altering, the max value is grayed out. To change the Max Value of Band 2
% you need to change the Min Value of Band 3, and this pattern follows suit
% for all bands.

% Individual Bands can have their Compute Types and Swerling Cases both
% viewed and changed.

% Grab the First Band and Second Band
Band_1 = RCS_FBands.Item(0);
Band_2 = RCS_FBands.Item(1);
Band_3 = RCS_FBands.Item(2);

% Look at current Frequency Settings of Band_1
B1_Min = Band_1.MinimumFrequency  % Min Frequency in MHz
B1_Max = Band_1.MaximumFrequency % Max Frequency in MHz

% Change Frequency Settings of Band_2
Band_2.MinimumFrequency = 4.0  % Min Frequency in MHz

Band_3.MinimumFrequency = 4.8 % **** This is the Equivalent of changing the Max Frequency of Band 2! ****

% View the new Band Values to Verify
B1_Min = Band_1.MinimumFrequency  % Min Frequency in MHz
B1_Max = Band_1.MaximumFrequency % Max Frequency in MHz
B2_Min = Band_2.MinimumFrequency  % Min Frequency in MHz
B2_Max = Band_2.MaximumFrequency % Max Frequency in MHz
B3_Min = Band_3.MinimumFrequency  % Min Frequency in MHz
B3_Max = Band_3.MaximumFrequency % Max Frequency in MHz

%% Inspect and Alter the Compute Types

% View the Supported Compute Types for this RCS
Available_Types = Band_1.SupportedComputeStrategies

% Set the Compute Strategy to Constant Value
Band_1.SetComputeStrategy('Constant Value')
Band_3.SetComputeStrategy('Constant Value')
Band_3.SetComputeStrategy('Constant Value')

% Set the Constant Value for Bands
Band_1.ComputeStrategy.ConstantValue = 5; % Value is in dBsm
Band_2.ComputeStrategy.ConstantValue = 8;
Band_3.ComputeStrategy.ConstantValue = 3;

%% Inspect and Alter the Swerling Case

% Retrieve the Current Swerling Case
Band_1.SwerlingCase

% Set the Swirling Case to be III instead of 0
Band_1.SwerlingCase = 3   % Valid options are 0,1,2,3, and 4



