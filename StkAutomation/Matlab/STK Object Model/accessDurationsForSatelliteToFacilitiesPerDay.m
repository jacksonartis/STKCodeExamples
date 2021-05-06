% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Sat Name
satName = 'Satellite_1_Primary_Mission_0600';
sensorName = 'Sensor1';
constName = 'Constellation1';
chainName = 'Chain1';

% Get one satellite
sat = scenario.Children.Item(satName);

% Add timekeeper sat
timeKeeper = scenario.Children.New('eSatellite','TimeKeeper');
command = strcat('SetState */Satellite/', timeKeeper.InstanceName, " ", 'Classical TwoBody',' "', scenario.StartTime, '" "', scenario.StopTime, '" ', '60 TrueOfDate', ' "', scenario.StartTime, '" ','42241100 0.0 0.0 0.0 0.0 270.101', ' "', scenario.StartTime, '"');
root.ExecuteCommand(command);

% Pull day times
satOrbits = timeKeeper.Vgt.EventIntervalLists.Item('PassIntervals');

% Connect to Constellation and Chain
if scenario.Children.Contains('eConstellation',constName)
    const = scenario.Children.Item(constName);
    const.Objects.RemoveAll;
else 
    const = scenario.Children.New('eConstellation',constName);
end

if scenario.Children.Contains('eChain',chainName)
    chain = scenario.Children.Item(chainName);
    chain.Objects.RemoveAll;
else 
    chain = scenario.Children.New('eChain',chainName);
end


% Add in facilities
facNames = [string('Dongara_Station_AUWA01_STDN_USDS'),string('Esrange_Station_KSX') ];
for i = 1:length(facNames)
   const.Objects.AddObject(scenario.Children.Item(facNames(i))); 
end

% Load and compute chain
chain.Objects.AddObject(sat);
chain.Objects.AddObject(const);
chain.ComputeAccess;

%% Get Max access duration between 1 satellite and any facility per orbit
% maxAccessDuration = 0;
% minAccessDuration = 1000;

accessMatrix = [];
accessNumberMatrix = [];
randomDay = [];
for i = 2:satOrbits.FindIntervals.Intervals.Count
    startTime = satOrbits.FindIntervals.Intervals.Item(int32(i-1)).Start;
    stopTime = satOrbits.FindIntervals.Intervals.Item(int32(i-1)).Stop;
    durDP = chain.DataProviders.Item('Complete Access').Exec(startTime,stopTime);
    if durDP.DataSets.Count > 0
        durations = cell2mat(durDP.DataSets.GetDataSetByName('Duration').GetValues);
        
        accessMatrix = vertcat(accessMatrix, sum(durations));
        accessNumberMatrix = vertcat(accessNumberMatrix, length(durations));
        if i == 10
            accessNum = cell2mat(durDP.DataSets.GetDataSetByName('Access Number').GetValues);
            startTimes = cell2mat(durDP.DataSets.GetDataSetByName('Start Time').GetValues);
            stopTimes = cell2mat(durDP.DataSets.GetDataSetByName('Stop Time').GetValues);
            dayOfData = horzcat(string(accessNum), string(startTimes), string(stopTimes), string(durations));
            startTime
            stopTime
        end
%         if min(durations) < minAccessDuration
%             minAccessDuration = min(durations);
%         end
%         if min(durations) > maxAccessDuration
%             maxAccessDuration = min(durations);
%         end
    end
end

% periodDP = sat.DataProviders.Item('Passes').Exec(scenario.StartTime, scenario.StopTime);
% period = periodDP.DataSets.GetDataSetByName('Period').GetValues;
period = 24*60*60;

% Print Access to Facility Object Data
maxAccessDuration = max(accessMatrix)
maxNumberOfAccess = max(accessNumberMatrix)
maxPercentageOfOrbit = (maxAccessDuration/period)*100

minAccessDuration = min(accessMatrix)
minNumberOfAccess = min(accessNumberMatrix)
minPercentageOfOrbit = (minAccessDuration/period)*100

totalAccess = sum(accessMatrix)
avgAccessPerDay = mean(accessMatrix)
avgAccessNumberPerDay = mean(accessNumberMatrix)


% dayInSeconds = 24*60*60;
% revolutions = round(dayInSeconds/period);
% accessValues = [];
% for i = 1:satOrbits.FindIntervals.Intervals.Count
%     startTime = satOrbits.FindIntervals.Intervals.Item(int32(i-1)).Start;
%     stopTime = satOrbits.FindIntervals.Intervals.Item(int32(i-1)).Stop;
%     durDP = chain.DataProviders.Item('Complete Access').Exec(startTime,stopTime);
%     if durDP.DataSets.Count > 0
%         durations = cell2mat(durDP.DataSets.GetDataSetByName('Duration').GetValues);
%         accessValues = vertcat(accessValues, durations);
%     end
% end

scenario.Children.Unload('eSatellite',timeKeeper.InstanceName);
