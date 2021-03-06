function [objPaths] = GetObjectNames(objType,name)
% Author: John Thompson
% Organization: Analytical Graphics Inc.
% Date Created: 4/01/18
% Modified: 10/3/18 by Kyle Kochel
% Modified: 12/20/18 by Alex Ridgeway (changed the output from Full File
% Path to just the NAMES of the objects.
% Description: Grabs all objects of a specifed type and returns their file
% path. Optionally the objects can also be filtered by name.

%% Inputs:
% objType: Object Type in STK, i.e. Satellite, Transmitter, etc. [string]
% name: Optionally Filter by Objects Containing this Name [string]

% Example 1: Find all satellites in the scenario
% FilterObjectsByType('Satellite')
%% Code
% Attatch to an Existing Instance of STK
uiApplication = actxGetRunningServer('STK12.application');
root = uiApplication.Personality2;

% Write Objects in the Scenario to an XML File
xml = root.AllInstanceNamesToXML;
fileID = fopen('ObjectNames.xml','w');
fprintf(fileID,'%s',xml);
fclose(fileID);

% Grab All Objects in the XML of a User Specified Type
xmlDoc = xmlread('ObjectNames.xml');
allListItems = xmlDoc.getElementsByTagName('object');
objPaths = {};
k = 1;

% Grab All Objects of the Specified Type and Name
for i = 1:allListItems.getLength
    % Grab the Specifed Object Type
    ListItem = allListItems.item(i-1);
    if strcmp(char(ListItem.getAttribute('class')),char(objType)) == 1
        temp = strsplit(char(ListItem.getAttribute('path')),'/');
        % Store Objects
        if nargin == 2 && contains(temp{end}, name)
            fullpath = char(ListItem.getAttribute('path'));
            medpath = strsplit(fullpath,'/');
            shrtpath = medpath(end);
            objPaths{k} = shrtpath;
            k = k + 1;
        elseif nargin == 1
            fullpath = char(ListItem.getAttribute('path'));
            medpath = strsplit(fullpath,'/');
            shrtpath = medpath(end);
            objPaths{k} = shrtpath;
            k = k + 1;
        end
    end
end
end