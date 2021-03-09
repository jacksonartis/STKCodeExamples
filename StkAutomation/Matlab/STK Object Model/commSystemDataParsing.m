% Access STK
app = actxGetRunningServer('STK12.application');
root = app.Personality2;

% Set Variables
transmitter = 'Transmitters';
receiver = 'Receivers';
jammers = 'Interference';
systemName = 'CommSystem1';
timeStep = 60;


% Get STK Object References
scenario = root.CurrentScenario;
transConst = scenario.Children.Item(transmitter);
receiveConst = scenario.Children.Item(receiver);
jammers = scenario.Children.Item(jammers);
commSys = scenario.Children.Item(systemName);
receiverPath = receiveConst.Objects.Item(int32(0)).Path;
parsedReceiverPath = split(receiverPath, '/');
truncReceiverPath = "";
for i = 6:length(parsedReceiverPath)
    if i ~= length(parsedReceiverPath)
        truncReceiverPath = strcat(truncReceiverPath,string(parsedReceiverPath(i)),'/');
    else
        truncReceiverPath = strcat(truncReceiverPath,string(parsedReceiverPath(i)));
    end
end
receiverParentObj = scenario.Children.Item(string(parsedReceiverPath(7)));


for i = 0:jammers.Objects.Count-1
    % Get Access Values
    objPath = jammers.Objects.Item(int32(i)).Path;
    parsedPath = split(objPath, '/');
    currentParentName = string(parsedPath(7));
    currentParent = scenario.Children.Item(currentParentName);
    currentAccess = currentParent.GetAccessToObject(receiverParentObj);
    accessIntervals = currentAccess.vgt.EventIntervalLists.Item('AccessIntervals');

    % Create All Values Matricies
    Time = [];
    linkToId = [];
    xmtrPower = [];
    xmtrGain = [];
    EIRP = [];
    xmtrAzimuth_Phi = [];
    xmtrElevation_Theta = [];
    rcvrAzimuth_Phi = [];
    rcvrElevation_theta = [];
    rcvdFrequency = [];
    rcvdIsoPower = [];
    carrierPowerAtRcvrInput = [];
    fluxDensity = [];
    rcvrGain = [];
    freeSpaceLoss = [];
    atmosLoss = [];
    rainLoss = [];
    cloudsFogLoss = [];
    tropScintillLoss = [];
    propLoss = [];
    train = [];
    tatmos = [];
    tCloudsFog = [];
    tropoScintill = [];
    tSun = [];
    tEarth = [];
    tCosmic = [];
    tOther = [];
    tAntenna = [];
    tEquiv = [];
    g_T = [];
    c_No = [];
    c_No_Io = [];
    bandwidth = [];
    bandwidth_Overlap = [];
    polRelAngle = [];
    polarizationEffic = [];
    c_N = [];
    c_N_I = [];
    eb_No = [];
    eb_No_Io = [];
    ber = [];
    ber_I = [];
    c_I = [];
    deltaT_T = [];
    pwrFluxDensity = [];

    for j = 0:accessIntervals.FindIntervals.Intervals.Count-1
        startTime = accessIntervals.FindIntervals.Intervals.Item(j).Start;
        stopTime = accessIntervals.FindIntervals.Intervals.Item(j).Stop;
        commProvCenter = commSys.DataProviders.GetDataPrvTimeVarFromPath('Link Information');
        commProvCenter.PreData = truncReceiverPath;
        rptElems = {'Time';'Link To ID';'Xmtr Power';'Xmtr Gain';'EIRP';'Xmtr Azimuth - Phi';'Xmtr Elevation - Theta';'Rcvr Azimuth - Phi';'Rcvr Elevation - Theta';'Rcvd. Frequency';'Rcvd. Iso. Power';'Flux Density';'Rcvr Gain';'Train';'Tatmos';'Tsun';'Tearth';'Tcosmic';'Tother';'Tantenna';'Tequiv';'g/T';'C/No';'C/(No+Io)';'Bandwidth';'Bandwidth Overlap';'C/N';'C/(N+I)';'Eb/No';'Eb/(No+Io)';'BER';'BER+I';'C/I';'DeltaT/T';'Pwr Flux Density'};
        commDP = commProvCenter.Exec(startTime, stopTime, timeStep);
        if commDP.DataSets.Count > 0
            Time = vertcat(Time, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(1))).GetValues));
            linkToId = vertcat(linkToId, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(2))).GetValues));
            xmtrPower = vertcat(xmtrPower, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(3))).GetValues));
            xmtrGain = vertcat(xmtrGain, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(4))).GetValues));
            EIRP = vertcat(EIRP, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(5))).GetValues));
            xmtrAzimuth_Phi = vertcat(xmtrAzimuth_Phi, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(6))).GetValues));
            xmtrElevation_Theta = vertcat(xmtrElevation_Theta, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(7))).GetValues));
            rcvrAzimuth_Phi = vertcat(rcvrAzimuth_Phi, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(8))).GetValues));
            rcvrElevation_theta = vertcat(rcvrElevation_theta, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(9))).GetValues));
            rcvdFrequency = vertcat(rcvdFrequency, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(10))).GetValues));
            rcvdIsoPower = vertcat(rcvdIsoPower, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(11))).GetValues));
            fluxDensity = vertcat(fluxDensity, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(12))).GetValues));
            rcvrGain = vertcat(rcvrGain, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(13))).GetValues));
            train = vertcat(train, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(14))).GetValues));
            tatmos = vertcat(tatmos, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(15))).GetValues));
            tSun = vertcat(tSun, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(16))).GetValues));
            tEarth = vertcat(tEarth, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(17))).GetValues));
            tCosmic = vertcat(tCosmic, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(18))).GetValues));
            tOther = vertcat(tOther, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(19))).GetValues));
            tAntenna = vertcat(tAntenna, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(20))).GetValues));
            tEquiv = vertcat(tEquiv, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(21))).GetValues));
            g_T = vertcat(g_T, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(22))).GetValues));
            c_No = vertcat(c_No, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(23))).GetValues));
            c_No_Io = vertcat(c_No_Io, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(24))).GetValues));
            bandwidth = vertcat(bandwidth, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(25))).GetValues));
            bandwidth_Overlap = vertcat(bandwidth_Overlap, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(26))).GetValues));
            c_N = vertcat(c_N, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(27))).GetValues));
            c_N_I = vertcat(c_N_I, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(28))).GetValues));
            eb_No = vertcat(eb_No, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(29))).GetValues));
            eb_No_Io = vertcat(eb_No_Io, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(30))).GetValues));
            ber = vertcat(ber, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(31))).GetValues));
            ber_I = vertcat(ber_I, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(32))).GetValues));
            c_I = vertcat(c_I, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(33))).GetValues));
            deltaT_T = vertcat(deltaT_T, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(34))).GetValues));
            pwrFluxDensity = vertcat(pwrFluxDensity, cell2mat(commDP.DataSets.GetDataSetByName(string(rptElems(35))).GetValues));
        end      
    end
    Time = string(Time);
    % Write to Excel Document
    filename = 'parsedCommSystemData.xlsx';
    id = 'MATLAB:xlswrite:AddSheet';
    warning('off',id);

     headers = rptElems.';
     headers = string(headers);
     xlswrite(filename,headers,1,'A1');

     sheet_Num = i+1;
     xlswrite(filename,Time,sheet_Num,'A2');
     xlswrite(filename,linkToId,sheet_Num,'B2');
     xlswrite(filename,xmtrPower,sheet_Num,'C2');
     xlswrite(filename,xmtrGain,sheet_Num,'D2');
     xlswrite(filename,EIRP,sheet_Num,'E2');
     xlswrite(filename,xmtrAzimuth_Phi,sheet_Num,'F2');
     xlswrite(filename,xmtrElevation_Theta,sheet_Num,'G2');
     xlswrite(filename,rcvrAzimuth_Phi,sheet_Num,'H2');
     xlswrite(filename,rcvrElevation_theta,sheet_Num,'I2');
     xlswrite(filename,rcvdFrequency,sheet_Num,'J2');
     xlswrite(filename,rcvdIsoPower,sheet_Num,'K2');
     xlswrite(filename,fluxDensity,sheet_Num,'L2');
     xlswrite(filename,rcvrGain,sheet_Num,'M2');
     xlswrite(filename,train,sheet_Num,'N2');
     xlswrite(filename,tatmos,sheet_Num,'O2');
     xlswrite(filename,tSun,sheet_Num,'P2');
     xlswrite(filename,tEarth,sheet_Num,'Q2');
     xlswrite(filename,tCosmic,sheet_Num,'R2');
     xlswrite(filename,tOther,sheet_Num,'S2');
     xlswrite(filename,tAntenna,sheet_Num,'T2');
     xlswrite(filename,tEquiv,sheet_Num,'U2');
     xlswrite(filename,g_T,sheet_Num,'V2');
     xlswrite(filename,c_No,sheet_Num,'W2');
     xlswrite(filename,c_No_Io,sheet_Num,'X2');
     xlswrite(filename,bandwidth,sheet_Num,'Y2');
     xlswrite(filename,bandwidth_Overlap,sheet_Num,'Z2');
     xlswrite(filename,c_N,sheet_Num,'AA2');
     xlswrite(filename,c_N_I,sheet_Num,'AB2');
     xlswrite(filename,eb_No,sheet_Num,'AC2');
     xlswrite(filename,eb_No_Io,sheet_Num,'AD2');
     xlswrite(filename,ber,sheet_Num,'AE2');
     xlswrite(filename,ber_I,sheet_Num,'AF2');
     xlswrite(filename,c_I,sheet_Num,'AG2');
     xlswrite(filename,deltaT_T,sheet_Num,'AH2');
     xlswrite(filename,pwrFluxDensity,sheet_Num,'AI2');
end