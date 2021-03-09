% Grab STK Instance
app = actxGetRunningServer('STK12.Application');
root = app.Personality2;
scenario = root.CurrentScenario;
coverageDefinitionName = 'CoverageDefinition1';

% Get Coverage Definition and Line Targets
covDef = scenario.Children.Item('CoverageDefinition1');
lineTargets = scenario.Children.GetElements('eLineTarget');
covDef.Grid.BoundsType = 'eBoundsCustomBoundary';

% Add Line Targets for Custom Boundary
for i = 1:lineTargets.Count
    covDef.Grid.Bounds.BoundaryObjects.AddObject(lineTargets.Item(int32(i - 1)));
end 