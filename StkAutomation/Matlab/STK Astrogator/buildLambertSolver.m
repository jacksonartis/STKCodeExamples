function  buildLambertSolver(root, startPlanet, depAng, arrAng, depRadius, arrRadius, depCirc, arrCirc,depGrav, flybyEpoch, timeOfFlight)
connectCommands = [];
command = strcat('ComponentBrowser */ Duplicate "Design Tools" "Lambert Solver" to', startPlanet);
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'LambertToolMode "Specify initial and final central bodies"');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'CentralBody Sun');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Departure.CentralBody Earth');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Arrival.CentralBody', " ", startPlanet);
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Departure.Angle', " ", string(depAng)," ", 'deg');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Arrival.Angle', " ", string(arrAng)," ", 'deg');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Departure.RadiusScaleFactor', " ", string(depRadius));
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Arrival.RadiusScaleFactor', " ", string(arrRadius));
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Departure.UseCircVel', " ", string(depCirc));
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Arrival.UseCircVel', " ", string(arrCirc));
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Propagator Earth_HPOP_Default_v10');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'SequenceName', " ", startPlanet, '_Flyby');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'MinimumTOF', " ", string(timeOfFlight), " " , 'day');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'InitEpoch', " ", string(flybyEpoch), " ", 'UTCG');
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'AutoAddToCB True');
connectCommands = horzcat(connectCommands, command);
if depCirc
    command = strcat('ComponentBrowser */ SetValue "Design Tools" to',startPlanet', " ", 'Departure.ConsiderGravity', " ", string(depGrav));
    connectCommands = horzcat(connectCommands, command);
end
command = strcat('ComponentBrowser */ LambertCompute "Design Tools" to',startPlanet);
connectCommands = horzcat(connectCommands, command);
command = strcat('ComponentBrowser */ LambertConstructSequence "Design Tools" to',startPlanet);
connectCommands = horzcat(connectCommands, command);


for i = 1:length(connectCommands)
    root.ExecuteCommand(connectCommands(i));
end
end

