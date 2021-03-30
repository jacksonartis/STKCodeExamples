function ami_gui
% This function will create a GUI that allows your to build an astrogator
% satellite with an MCS populated by a either a .csv or .xml file

   %  Create and then hide the UI as it is being constructed.
   f = figure('Visible','off','Position',[620,400,500,300]);
   f.Name = 'Astrogator Maneuver Ingester GUI';
   
   %  Construct the components.
   load = uicontrol('Style','pushbutton','String','Load Scenario',...
           'Position',[175,260,100,25], 'Callback',@loadbutton_Callback);
   grab_sat = uicontrol('Style','pushbutton','String','Attach to Satellite',...
           'Position',[175,220,100,25], 'Callback',@grabsat_Callback);
   extension_text  = uicontrol('Style','text','String','Select File Extension',...
           'Position',[175,200,150,15]);
   select_data_type = uicontrol('Style','popupmenu',...
           'String',{'*.csv','*.xml'},...
           'Position',[175,170,100,25], 'Callback',@ext_Callback);
   select_data = uicontrol('Style','pushbutton',...
           'String','Select Burn File',...
           'Position',[175,130,100,25],'Callback',@selectfile_Callback); 
   prop = uicontrol('Style','pushbutton',...
           'String','Propagate',...
           'Position',[175,90,100,25], 'Callback', @prop_Callback); 
   clear = uicontrol('Style','pushbutton',...
           'String','Clear Variables',...
           'Position',[175,50,100,25], 'Callback',@clearall_Callback); 
   align([load,grab_sat,extension_text,select_data_type,select_data,prop,clear],'Center','None');
   
   % Make the UI visible.
   f.Visible = 'on';
   
   % Set Global Variables to be used
   global app
   global root
   global scenario
   global indx
   global sats
   global ext

    function loadbutton_Callback(source, eventdata)
        try
            app = actxGetRunningServer('STK12.Application');
            root = app.Personality2;
            scenario = root.CurrentScenario;
            msg = strcat('Connected to scenario named:'," ",scenario.InstanceName);
            msgbox(msg)
            
        catch
            app = actxserver('STK12.application');
            root = app.Personality2; 
            answers = inputdlg({'Scenario Name','Start Time','Stop Time'},...
              'Customer', [1 50; 1 50; 1 50]); 
            scenario = root.Children.New('eScenario',answers{1}); % Insert and Name the scenario
            scenario.SetTimePeriod(answers{2},answers{3}); % Set the initial Start and Stop Time
            scenStart = scenario.StartTime; % Save the Start time to so that it can be edited and pushed back to the scenario for each run
            root.Rewind();
        end
    end

    function grabsat_Callback(source, eventdata)
        satNum = scenario.Children.GetElements('eSatellite').Count;
        if satNum == 0
            satName = inputdlg({'Satellite Name'}, 'Customer',[1 50]);
            newSat = scenario.Children.New('eSatellite', string(satName{1}));
            indx = 1;
            sats = scenario.Children.GetElements('eSatellite');
        else 
            sats = scenario.Children.GetElements('eSatellite');
            list = [];
            for i = 1:satNum-1
               list = horzcat(list, string(sats.Item(int32(i-1)).InstanceName));
            end
            [indx,tf] = listdlg('ListString',list);
        end
    end

    function ext_Callback(source, eventdata)
        str = get(source, 'String');
        val = get(source, 'Value');
        
        switch str{val};
            case '*.csv'
                ext = string('*.csv')
            case '*.xml'
                ext = string('*.xml');
        end       
    end

    function selectfile_Callback(source, eventdata)
        [file,path] = uigetfile(ext);
        fullfilepath = strcat(path,file);
        [num,txt,raw] = xlsread(fullfilepath);
        array = size(txt);
        rowNum = array(1);
        start = txt(2:rowNum,3);
        ender = txt(2:rowNum,4);

        startTimes = cell2mat(start);
        endTimes = cell2mat(ender);
        
        compBrowser = scenario.ComponentDirectory.GetComponents('eComponentAstrogator').GetFolder('Engine Models');

        componentName = 'New Engine';
        try compBrowser.Item(componentName);
            % Propagator already exists
        catch
            % Duplicate default HPOP propagator
            engineComponent = compBrowser.Item('Constant Thrust and Isp');
            engineComponent.CloneObject();

            % Set name
            newComponent = compBrowser.Item('Constant Thrust and Isp1');
            newComponent.Name = componentName;
        end 
        
        for i = 1:length(indx)
            currentSat = sats.Item(int32(indx(i) - 1));
            if currentSat.PropagatorType ~= string('ePropagatorAstrogator')
                currentSat.SetPropagatorType('ePropagatorAstrogator');
            end
            orbElem = inputdlg({'X Position (km)','Y Position (km)','Z Position (km)', 'X Velocity (km/s)', 'Y Velocity (km/s)', 'Z Velocity (km/s)'},...
              'Customer', [1 30; 1 30; 1 30; 1 30; 1 30; 1 30]);
            MCS = currentSat.Propagator.MainSequence;
            InitState = MCS.Item('Initial State').Element;
            InitState.X = str2double(orbElem{1});
            InitState.Y = str2double(orbElem{2});
            InitState.Z = str2double(orbElem{3});
            InitState.Vx = str2double(orbElem{4});
            InitState.Vy = str2double(orbElem{5});
            InitState.Vz = str2double(orbElem{6});
            MCS.Remove('Propagate');
            firstPropagate = MCS.Insert(5,'Prop','-');
            firstPropagate.OverrideMaxPropagationTime = true;
            startYear = scenario.StartTime(8:11);
            startMonth = scenario.StartTime(4:6);
            startDay = scenario.StartTime(1:2);
            startTime = scenario.StartTime(13:24);
            fullScenStart = strcat(startYear,"/",startMonth,"/",startDay," ", startTime);

            for i = 1:rowNum-1
    
                % Get Thrust Values
                x_num = num(i,1);
                y_num = num(i,2);
                z_num = num(i,3);

                totalMag = sqrt(x_num^2 + y_num^2 + z_num^2);
                newComponent.Thrust = totalMag;

                x_value = x_num/totalMag;
                y_value = y_num/totalMag;
                z_value = z_num/totalMag;

                % Turn into correct date format
                year1 = startTimes(i,1:4);
                month1 = startTimes(i,6:7);
                day1 = startTimes(i,9:10);
                time1 = startTimes(i,12:23);
                fullStartDate = strcat(year1,"/",month1,"/",day1," ", time1);

                 if i == 1
                    FirstProp = etime(datevec(fullStartDate),datevec(fullScenStart));
                    firstPropagate.StoppingConditions.Item('Duration').Properties.Trip = FirstProp;
                 end


                year2 = endTimes(i,1:4);
                month2 = endTimes(i,6:7);
                day2 = endTimes(i,9:10);
                time2 = endTimes(i,12:23);
                fullEndDate = strcat(year2,"/",month2,"/",day2," ", time2);

                tripDuration = etime(datevec(fullEndDate),datevec(fullStartDate));

                name = strcat("Burn",string(i));
                newBurn = MCS.Insert(2,name,'-');
                newBurn.SetManeuverType(1);
                newBurn.Maneuver.SetAttitudeControlType(4);
                newBurn.Maneuver.AttitudeControl.ThrustVector.AssignXYZ(x_value,y_value,z_value);
                newBurn.Maneuver.Propagator.StoppingConditions.Item('Duration').Properties.Trip = tripDuration;
                commndPart1 = strcat("Astrogator */Satellite/",currentSat.InstanceName, " ", "SetValue MainSequence.SegmentList.");
                commandPart2 = strcat(name, ".");
                commandPart3 = "FiniteMnvr.EngineModel New_Engine";
                engineCommand = strcat(commndPart1, commandPart2, commandPart3);
                root.ExecuteCommand(engineCommand);

                if i ~= rowNum-1
                    year3 = startTimes(i+1,1:4);
                    month3 = startTimes(i+1,6:7);
                    day3 = startTimes(i+1,9:10);
                    time3 = startTimes(i+1,12:23);
                    nextStartDate = strcat(year3,"/",month3,"/",day3," ", time3);

                    propDuration = etime(datevec(nextStartDate),datevec(fullEndDate));
                    if propDuration ~= 0
                        newProp = currentSat.Propagator.MainSequence.Insert(5,'Prop','-');
                        newProp.StoppingConditions.Item('Duration').Properties.Trip = propDuration;
                    end 

                end 

            end 
        end
    end

    function prop_Callback(source, eventdata)
        for i = length(indx)
            currentSat = sats.Item(int32(indx(i) - 1));
            currentSat.Propagator.RunMCS;
        end
    end

    function clearall_Callback(source, eventdata)
        root = [];
        app = [];
    end
   
end