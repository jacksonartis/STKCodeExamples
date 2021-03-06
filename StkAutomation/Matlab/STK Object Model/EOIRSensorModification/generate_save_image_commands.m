function [connectCommands] = generate_save_image_commands(sensor_names,output_folder, satName)
    sensor_path_prefix = ['*/Satellite/' satName '/Sensor/'];
    output_folder = [output_folder filesep()];
    connectCommands = [];
    
    band_names{1} = 'ColorBlue';
    band_names{2} = 'ColorGreen';
    band_names{3} = 'ColorRed';
    
    for i=1:length(sensor_names)
        for j=1:length(band_names)
            connectCommands = vertcat(connectCommands, string(sprintf('EOIRDetails %s%s SaveSceneRawData "%s%s_%s.txt" %s',sensor_path_prefix,sensor_names{i},output_folder,sensor_names{i},band_names{j},band_names{j})));
         end
    end
end