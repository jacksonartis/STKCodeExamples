% Open STK and load scenario
app = actxserver('STK12.Application');
root = app.Personality2;
[file,path] = uigetfile('.vdf');
fullfilepath = strcat(path,file);
root.Load(fullfilepath);
scenario = root.CurrentScenario;
scenName = scenario.InstanceName;
% Store Sensor names
name = inputdlg({'Satellite Name'},...
              'Customer', [1 30]);
sat = scenario.Children.Item(name{1});
sensorNames = [];

% Collect EOIR Sensor Names
for i = 1:sat.Children.Count
  if sat.Children.Item(int32(i-1)).PatternType == "eSnEOIR"
    sensorNames = horzcat(sensorNames, string(sat.Children.Item(int32(i-1)).InstanceName)); 
  end
end



% Create File Repositories
filePathName = inputdlg({'Foler Name'},...
              'Customer', [1 30]);
          
mkdir(string(filePathName{1}));

mkdir(strcat(filePathName{1},'\t0')) 
mkdir(strcat(filePathName{1},'\t0'),'\ShortIntegrationColor') 
mkdir(strcat(filePathName{1},'\t0'),'\LongIntegrationColor')  

output_folder = char(strcat(pwd,'\',filePathName{1}));

output_folder_short = [output_folder filesep() 't0' filesep() 'ShortIntegrationColor'];
output_folder_long = [output_folder filesep() 't0' filesep() 'LongIntegrationColor'];

% Create Connect Commands
commands_short = generate_save_image_commands(sensorNames, output_folder_short, name{1});
commands_long = generate_save_image_commands(sensorNames, output_folder_long, name{1});

for i = 1:length(all_commands_short)
    root.ExecuteCommand(commands_short(i));
end

update_all_satellite_sensor_integration_times(sat, 'slow'); 

for i = 1:length(all_commands_long)
    root.ExecuteCommand(commands_long(i));
end

