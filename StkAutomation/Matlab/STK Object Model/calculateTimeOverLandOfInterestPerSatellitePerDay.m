% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;

% Object Names
areaTargets = scenario.Children.GetElements('eAreaTarget');
constName = 'Land';
chainName = 'Chain1';

% Add timekeeper sat
timeKeeper = scenario.Children.New('eSatellite','TimeKeeper');
command = strcat('SetState */Satellite/', timeKeeper.InstanceName, " ", 'Classical TwoBody',' "', scenario.StartTime, '" "', scenario.StopTime, '" ', '60 TrueOfDate', ' "', scenario.StartTime, '" ','42241100 0.0 0.0 0.0 0.0 270.101', ' "', scenario.StartTime, '"');
root.ExecuteCommand(command);

if scenario.Children.Contains('eConstellation',constName)
    const = scenario.Children.Item(constName);
    %const.Objects.RemoveAll;
else 
    const = scenario.Children.New('eConstellation',constName);
end

if scenario.Children.Contains('eChain',chainName)
    chain = scenario.Children.Item(chainName);
    %chain.Objects.RemoveAll;
else 
    chain = scenario.Children.New('eChain',chainName);
end

% for i = 1:areaTargets.Count
%    currTarget = areaTargets.Item(int32(i-1));
%    const.Objects.AddObject(currTarget); 
%    if currTarget.AccessConstraints.IsNamedConstraintActive('ElevationAngle') && currTarget.AccessConstraints.GetActiveNamedConstraint('ElevationAngle').Angle ~= 90
%        currTarget.AccessConstraints.GetActiveNamedConstraint('ElevationAngle').Angle = 90;
%    else
%        cons = currTarget.AccessConstraints.AddNamedConstraint('ElevationAngle');
%        cons.Angle = 90;
%    end
% end

sat = scenario.Children.Item('Satellite_1_Primary_Mission_0600');
satOrbits = timeKeeper.Vgt.EventIntervalLists.Item('PassIntervals');

 %chain.Objects.AddObject(sat);
 %chain.Objects.AddObject(const);
 %chain.ComputeAccess;

maxAccessDuration = 0;
minAccessDuration = 100000;
accessMatrix = [];
for i = 2:satOrbits.FindIntervals.Intervals.Count
    startTime = satOrbits.FindIntervals.Intervals.Item(int32(i-1)).Start;
    stopTime = satOrbits.FindIntervals.Intervals.Item(int32(i-1)).Stop;
    durDP = chain.DataProviders.Item('Complete Access').Exec(startTime,stopTime);
    if durDP.DataSets.Count > 0
        durations = cell2mat(durDP.DataSets.GetDataSetByName('Duration').GetValues);
        accessMatrix = vertcat(accessMatrix, sum(durations));
%         if sum(durations) > maxAccessDuration
%             maxAccessDuration = sum(durations);
%         end
%         if sum(durations) < minAccessDuration
%             minAccessDuration = sum(durations);
%         end
    end
end

% periodDP = sat.DataProviders.Item('Passes').Exec(scenario.StartTime, scenario.StopTime);
% period = periodDP.DataSets.GetDataSetByName('Period').GetValues;
 period = 24*60*60;

maxOverLandPerDay = max(accessMatrix);
maxOverLandPerDayPercentage = (maxOverLandPerDay/period)*100

minOverLandPerDay = min(accessMatrix);
minOverLandPerDayPercentage = (minOverLandPerDay/period)*100

avgOverLandPerDay = mean(accessMatrix)
avgOverLandPerDayPercent = (avgOverLandPerDay/period)*100

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


