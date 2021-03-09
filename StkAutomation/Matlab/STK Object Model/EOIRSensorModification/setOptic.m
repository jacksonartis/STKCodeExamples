function setOptic(matrix, opt_strings, currentBand, spectralTransmission, dataFile)

    inputChoices = string({'eFNumberAndFocalLength','eFNumberAndApertureDiameter','eFocalLengthAndApertureDiameter'});
    imageQuality = string({'eDiffractionLimited','eNegligibleAberrations','eMildAberrations','eModerateAberrations','eCustomWavefrontError','eCustomPupilFunction','eCustomPSF','eCustomMTF'});
    optTransmission = string({'eBandEffectiveTransmission','eTransmissionDataFile'});
    diffWavelength = string({'eLowBandEdge','eBandCenter','eHighBandEdge','eUserDefinedWavelength'});
    
    currentBand.OpticalInputMode = find(inputChoices == opt_strings(1))- 1;
    currentBand.ImageQuality = find(imageQuality == opt_strings(2)) - 1;
    currentBand.OpticalTransmissionMode = find(optTransmission == opt_strings(3)) - 1;
    currentBand.WavelengthType = find(diffWavelength == opt_strings(4)) - 1;
    index = 1;
    
    if opt_strings(1) == string('eFNumberAndFocalLength')
        currentBand.Fnumber = matrix(index);
        index = index + 1;
        currentBand.EffFocalL = matrix(index);
        index = index + 1;
    elseif opt_strings(1) == string('eFNumberAndApertureDiameter')
        currentBand.Fnumber = matrix(index);
        index = index + 1;
        currentBand.EntrancePDia = matrix(index);
        index = index + 1;
    else
        currentBand.EffFocalL = matrix(index);
        index = index + 1;
        currentBand.EntrancePDia = matrix(index);
        index = index + 1;
    end 



    if opt_strings(2) == string('eCustomMTF')
        currentBand.OpticalQualityDataFileFrequencySampling = matrix(index);
        index = index + 1;
        currentBand.OpticalQualityDataFile = dataFile;
    elseif opt_strings(2) == string('eCustomPSF')
        currentBand.OpticalQualityDataFileSpatialSampling = matrix(index);
        index = index + 1;
        currentBand.OpticalQualityDataFile = dataFile;
    elseif opt_strings(2) == string('eCustomWavefrontError')
        currentBand.RMSWavefrontError = matrix(index);
        index = index + 1;
        currentBand.LongDFocus = matrix(index);
        index = index + 1;
    else
        currentBand.LongDFocus = matrix(index);
        index = index + 1;
    end 

    
    if opt_strings(3) == string('eTransmissionDataFile')
        currentBand.OpticalTransmissionSpectralResponseFile = spectralTransmission;
    else
        currentBand.OpticalTransmission = matrix(index);
        index = index + 1;
    end

    if opt_strings(4) == string('eUserDefinedWavelength')
            currentBand.Wavelength = matrix(index);
    end
    
end

