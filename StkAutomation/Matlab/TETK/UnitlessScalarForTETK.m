% Access STK
app = actxGetRunningServer('STK12.application');
root = app.Personality2;

% Set Variables
objectOfInterest = 'Aircraft1';
unitlessValue = 'Mach';
dataProvider = 'Flight Profile By Time';
dataElement = 'Mach #';
timeStep = 1; % seconds
fileEnding = 'Scalar.csc'; % keep the same

% Access Scenario
scenario = root.CurrentScenario;
obj = scenario.Children.Item(objectOfInterest);
[filepath, ~, ~] = fileparts(string(scenario.ScenarioFiles(1)));

% Pull Data Providers
objDP = obj.DataProviders.GetDataPrvTimeVarFromPath(dataProvider).Exec(scenario.StartTime, scenario.StopTime, timeStep);
times = cell2mat(objDP.DataSets.GetDataSetByName('Time').GetValues);
elementValue = cell2mat(objDP.DataSets.GetDataSetByName(dataElement).GetValues);
dataPoints = length(elementValue);

% Create Scalar File
newFilepath = strcat(filepath,'\',unitlessValue,fileEnding);
fid = fopen(newFilepath,'w+');
fprintf(fid, 'stk.v.12.0\n\nBEGIN Data\nTimeFormat\tUTCG\nUnitType\t\tDistance\nValueUnit\t\tm\nValueRateUnit\tm/sec\n');
fprintf(fid, 'NumberOfIntervals 1\n');
fprintf(fid, strcat('\tBEGIN Interval\n\t  NumberOfPoints'," ",string(dataPoints)));
fprintf(fid, '\n\t  BEGIN TimeValues\n');

for i = 1:dataPoints
    fprintf(fid, strcat('\t\t',times(i,1:24),'\t',string(1*elementValue(i)),'\n'));
end

fprintf(fid, '\n\tEND TimeValues');
fprintf(fid, '\n  END Interval\n');
fprintf(fid, 'END Data');
fclose(fid);
 

% Create Dummy File Path
dummyFile = strcat(filepath,'\','dummy',fileEnding);
    
    
if obj.Vgt.CalcScalars.Contains(unitlessValue) == 0

    customScalar = obj.Vgt.CalcScalars.Factory.CreateCalcScalarFile(unitlessValue, 'Scalar based off of the data providers',newFilepath);
    fid2 = fopen(dummyFile, 'w+');
    fprintf(fid2, 'stk.v.12.0\n\nBEGIN Data\nTimeFormat\tUTCG\nUnitType\t\tDistance\nValueUnit\t\tm\nValueRateUnit\tm/sec\n');
    fprintf(fid2, 'NumberOfIntervals 1\n');
    fprintf(fid2, strcat('\tBEGIN Interval\n\t  NumberOfPoints'," ",string(1)));
    fprintf(fid2, '\n\t  BEGIN TimeValues\n');
    fprintf(fid2, strcat('\t\t',times(1,1:24),'\t',string(elementValue(1)),'\n'));
    fprintf(fid2, '\n\tEND TimeValues');
    fprintf(fid2, '\n  END Interval\n');
    fprintf(fid2, 'END Data');
    fclose(fid2);
else
     customScalar.Filename = dummyFile;
     customScalar.Filename = newFilepath;
end 

