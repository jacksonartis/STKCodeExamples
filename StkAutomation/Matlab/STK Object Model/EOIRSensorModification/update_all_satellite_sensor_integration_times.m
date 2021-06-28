function [] = update_all_satellite_sensor_integration_times(satellite, integration_time)      
    if integration_time == string('fast')
            integration_time = 10;
    elseif integration_time == string('slow')
            integration_time = 4500;
    end
    sensitivity_value = 1*10^-18;
    currentSat = satellite;
    sensorList = currentSat.Children.GetElements('eSensor');
       for j=1:sensorList.Count
           currentSensor = sensorList.Item(int32(j-1));
           if currentSensor.InstanceName ~= string('Sensor1')
               pattern = currentSensor.Pattern;
               for k=1:pattern.Bands.Count
                 currentBand = pattern.Bands.Item(int32(k-1));
                 currentBand.IntegrationTime = integration_time;
                 currentBand.Sensitivities.Item(int32(0)).EquivalentValue = sensitivity_value;
                end
            end 
        end
end