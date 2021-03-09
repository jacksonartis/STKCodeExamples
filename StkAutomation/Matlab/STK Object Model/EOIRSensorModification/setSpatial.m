function setSpatial(matrix, stringArray, currentBand)

    if stringArray(1) == 'eNumPixAndPixelPitch'
        currentBand.SpatialInputMode = 2;
        currentBand.HorizontalPixels = matrix(1);
        currentBand.VerticalPixels = matrix(2);
        currentBand.HorizontalPP = matrix(3);
        currentBand.VerticalPP = matrix(4);
    elseif stringArray(1) == 'eFOVandPixelPitch'
        currentBand.SpatialInputMode = 1;
        currentBand.HorizontalHalfAngle = matrix(1);
        currentBand.VerticalHalfAngle = matrix(2);
        currentBand.HorizontalPP = matrix(3);
        currentBand.VerticalPP = matrix(4);
    else 
        currentBand.SpatialInputMode = 0;
        currentBand.HorizontalHalfAngle = matrix(1);
        currentBand.VerticalHalfAngle = matrix(2);
        currentBand.HorizontalPixels = matrix(3);
        currentBand.VerticalPixels = matrix(4);
    end
end
