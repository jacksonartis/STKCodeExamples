% Author: Alex Ridgeway
% Organization: Analytical Graphics Inc.
% Date Created: 12/20/18
% Description: Grabs all satellites in a scenario creates a TLE File from
% their orbits. The TLE Output file will be saved to your Matlab folder
% Path.

clear all;
close all; clc;

% Get reference to running STK instance
% Get our IAgStkObjectRoot interface
uiApplication = actxGetRunningServer('STK12.application');
root = uiApplication.Personality2;
scenario = root.CurrentScenario;
children = scenario.Children;

%Gets a list of all the Satellites in the Scenario.
satname = GetObjectNames('Satellite');

start = scenario.StartTime;
%stopt = scenario.StopTime;

%Error Handling if you re-run code multiple times
try
    tempsat =root.GetObjectFromPath('Satellite/tempsat'); 
    tempsat.Unload;
catch
end

%Starting SSC Number for the Satellites, increments from this number.
ssc = 60000;
fid = fopen('TLEData.tle','wt');

for n = 1:length(satname)

         Lsatname = strcat('Satellite/',satname{n}{1});
         cmd = strcat('GenerateTLE */',Lsatname,' Point "',start,'" ', sprintf('%05.0f',ssc) , ' 20 0.01 SGP4 tempsat');
         root.ExecuteCommand(cmd);
         
         %Make sure TLE information is valid and propagated on dummy satellite
         tempsat =root.GetObjectFromPath('Satellite/tempsat');
         cmd1 = strcat('GenerateTLE */Satellite/tempsat Point "',start,'" ', sprintf('%05.0f',ssc) , ' 20 0.01 SGP4 tempsat');
         root.ExecuteCommand(cmd1);

         %Extract TLE information from dummy satellite 
         satDP = tempsat.DataProviders.Item('TLE Summary Data').Exec();
         TLEData = satDP.DataSets.GetDataSetByName('TLE').GetValues;
         tempsat.Unload;
         
         %Write TLE to file
         fprintf(fid, '%s\n%s\n', TLEData{1,1}, TLEData{2,1});
         ssc = ssc +1;
end
fclose(fid);