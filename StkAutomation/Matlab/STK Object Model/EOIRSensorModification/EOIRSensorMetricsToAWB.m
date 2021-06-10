%% Connect to STK
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;
scenChildren = scenario.Children;
scenChildObj = {};
targetObjects = {};


for i = 1:scenChildren.Count
    scenChildObj{1,i} = scenChildren.Item(int32(i-1)).InstanceName;
    if scenChildren.Item(int32(i-1)).ClassName == string('Aircraft') || scenChildren.Item(int32(i-1)).ClassName == string('Satellite') || scenChildren.Item(int32(i-1)).ClassName == string('Missile')
        targetObjects{1,i} = scenChildren.Item(int32(i-1)).InstanceName;
    end
end

[filepath, ~, ~] = fileparts(string(scenario.ScenarioFiles(1)));
fileEnding = 'Scalar.csc';

%% Object Selection

continueVariable = true;

while continueVariable
    
    % Select Current Parent Object
    [indx,tf] = listdlg('PromptString',{'Select A Parent Object.',...
        'Only One Object Can Be Selected At A Time.',''},...
        'SelectionMode','single','ListString',scenChildObj);
    parentObj = scenario.Children.Item(string(scenChildObj(indx)));
    newObjList = scenChildObj';
    newObjList(indx,:) = [];
    newObjList = newObjList';
    secondContinueVariable = true;
    parentChildren = parentObj.Children;
    parentChildObj = {};
    
    for i = 1:parentChildren.Count
        parentChildObj{1,i} = parentChildren.Item(int32(i-1)).InstanceName;
    end
    
    % Select Sensor to Use
    while secondContinueVariable
        [indx,tf] = listdlg('PromptString',{'Select A Sensor Object.',...
            'Only One Sensor Can Be Selected At A Time.',''},...
            'SelectionMode','single','ListString',parentChildObj);
        currSensor = parentObj.Children.Item(string(parentChildObj(indx)));
        parentChildObj = parentChildObj';
        parentChildObj(indx,:) = [];
        parentChildObj = parentChildObj';
        calcObj = currSensor.Vgt.CalcScalars;
        list = {'EOIR Sensor Optics','EOIR Sensor Performance','EOIR Sensor To Target Metrics',...
            'EOIR Sensor Image Quality'};
        [indx, tf] = listdlg('PromptString',{'Select the Data Provider'},'ListString',list);
        
        for i = 1:length(indx)
            spot = indx(i);
            currDP = currSensor.DataProviders.Item(string(list(spot)));
            prompt1 = char(strcat('Select the Data Elements within')); 
            prompt2 = char(strcat('the', " ", list(spot), " ", 'Data Provider'));
            prompt3 = char('you would like.');
            
            if spot == 1 || spot == 2 || spot == 4
                newDP = currDP.Exec();
                dataSetElements = newDP.DataSets.ElementNames';
                [indx2, tf2] = listdlg('PromptString',{prompt1; prompt2; prompt3},'ListString',dataSetElements);
                
                for q = 1:length(indx2)
                    value = cell2mat(newDP.DataSets.GetDataSetByName(string(dataSetElements(indx2(q)))).GetValues);
                    dimension = newDP.DataSets.GetDataSetByName(string(dataSetElements(indx2(q)))).Dimension;
                    name = strcat(currSensor.InstanceName, "_", string(dataSetElements(indx2(q))));
                    name = char(name);
                    name = name(find(~isspace(name)));
                    if ~calcObj.Contains(name)
                        newCalc = calcObj.Factory.CreateCalcScalarConstant(name, 'Scalar of Constant Value');
                        newCalc.Value = value;
                        newCalc.Dimension = dimension;
                    else
                        newCalc = calcObj.Item(name);
                        newCalc.Value = value;
                        newCalc.Dimension = dimension;
                    end
                end
                
                
            else
                
                % Select Target Object
                editedTarg = string(targetObjects);
                editedTarg(editedTarg == parentObj.InstanceName) = [];
                editedTarg = cellstr(editedTarg);
                [indx2,tf] = listdlg('PromptString',{'Select A Parent Object.',...
                    'Only One Object Can Be Selected At A Time.',''},...
                    'SelectionMode','single','ListString',editedTarg);
                
                targetObject = scenario.Children.Item(string(editedTarg(indx2)));
                longPath = targetObject.Path;
                splitPath = strsplit(longPath, "/");
                for j = 1:currSensor.Pattern.Bands.Count
                    bandName = currSensor.Pattern.Bands.Item(int32(j - 1)).BandName;
                    preDataPath = strcat(splitPath(length(splitPath) - 1),"/",splitPath(length(splitPath)), " ", bandName);
                    currDP.PreData = preDataPath;
                    
                    % Determine Timestep
                    values = inputdlg({'Data Time Step (seconds)'},...
                        'Customer', [1 30]);
                    timestep = str2double(values{1});
                    newDP = currDP.Exec(scenario.StartTime, scenario.StopTime, timestep);
                    dataSetElements = newDP.DataSets.ElementNames;
                    dataSetElements(1,:) = [];
                    dataSetElements = dataSetElements';
                    prompt1 = char(strcat('Select the Data Elements within')); 
                    prompt2 = char(strcat('the', " ", list(spot), " ", 'Data Provider'));
                    prompt3 = char('you would like.');
                    [indx3, tf2] = listdlg('PromptString',{prompt1; prompt2; prompt3},'ListString',dataSetElements);
                    times = cell2mat(newDP.DataSets.GetDataSetByName('Time').GetValues);
                    
                    for k = 1:length(indx3)
                        
                        % Create CSC file
                        elementValue = cell2mat(newDP.DataSets.GetDataSetByName(string(dataSetElements(indx3(k)))).GetValues);
                        dimension = newDP.DataSets.GetDataSetByName(string(dataSetElements(indx3(k)))).DimensionName;
                        dataPoints = length(elementValue);
                        newFilepath = strcat(filepath,'\',string(dataSetElements(indx3(k))), " ",fileEnding);
                        fid = fopen(newFilepath,'w+');
                        fprintf(fid, strcat('stk.v.12.0\n\nBEGIN Data\nTimeFormat\tUTCG\nUnitType\t\t',dimension,'\n'));
                        fprintf(fid, 'NumberOfIntervals 1\n');
                        fprintf(fid, strcat('\tBEGIN Interval\n\t  NumberOfPoints'," ",string(dataPoints)));
                        fprintf(fid, '\n\t  BEGIN TimeValues\n');
                        
                        for m = 1:dataPoints
                            fprintf(fid, strcat('\t\t',times(m,1:24),'\t',string(1*elementValue(m)),'\n'));
                        end
                        
                        fprintf(fid, '\n\tEND TimeValues');
                        fprintf(fid, '\n  END Interval\n');
                        fprintf(fid, 'END Data');
                        fclose(fid);
                        
                        % Add file as a custom scalar object
                        name = char(dataSetElements(indx3(k)));
                        name = name(find(~isspace(name)));
                        name = strcat(name, "_", bandName);
                        if ~calcObj.Contains(name)
                            calcObj.Factory.CreateCalcScalarFile(name,"Data Providers", newFilepath);
                        else
                            newCalc = calcObj.Item(name);
                            newCalc = newFilepath;
                        end
                    end
                end
            end
        end
        
        moreSensors = questdlg('Get Values for an additional sensor on this parent object?', ...
            'Choices', ...
            'Yes','No','No');
        switch moreSensors
            case 'Yes'
                if length(parentChildObj) > 0
                    secondContinueVariable = true;
                else
                    secondContinueVariable = false;
                    message1 = '***NOTICE***: You have chosen to pull values on an additional sensor when';
                    message2 = strcat('there are no additional sensors under the current Parent Object named:', " ",parentObj.InstanceName); ;
                    message3 = 'This script will next prompt you to pick a NEW PARENT OBJECT';
                    msgbox({message1; message2; message3});
                end
            case 'No'
                secondContinueVariable = false;
        end
    end
    moreParents = questdlg('Pull values from sensors on a different parent objeect??', ...
        'Choices', ...
        'Yes','No','No');
    switch moreParents
        case 'Yes'
            continueVariable = true;
        case 'No'
            continueVariable = false;
    end
end