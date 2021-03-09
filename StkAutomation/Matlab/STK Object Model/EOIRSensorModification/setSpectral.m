function setSpectral(matrix, stringArray, currentBand)
    currentBand.HighBandEdgeWL = matrix(2);
    currentBand.LowBandEdgeWL = matrix(1);
    currentBand.NumIntervals = matrix(3);
    
    if stringArray == string('eProvideRSR')
        currentBand.RSRUnits = RSRUnits;
        currentBand.SystemRelativeSpectralResponseFile = spectral_response_path;
    end 
end