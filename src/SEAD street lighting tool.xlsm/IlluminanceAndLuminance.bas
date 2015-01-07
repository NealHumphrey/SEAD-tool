Attribute VB_Name = "IlluminanceAndLuminance"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry
Global a


Sub finalMatrices()

' Calculate final luminance and illuminace at each grid point from all the fixtures
Dim lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing
Dim polesetback, ArmLength
Dim calcMethod As String, poleconfig As String


'Clearing all illuminance and luminance sheets
wksIlluminanceOutput.Rows("12:5000").ClearContents
wksLuminanceOutput.Rows("12:5000").ClearContents
wksLuminanceOutputCIE.Rows("12:5000").ClearContents

'geometryValues is an array to hold all of the data about the road geometry. This makes it easier to pass it in/out of subroutines
Dim geometryValues()
ReDim geometryValues(0 To 10, 0 To 2)
'Just to help users remember what is in the array, adding header data
'column headers in row 0
geometryValues(0, 0) = "Variable Name"
geometryValues(0, 1) = "Baseline"
geometryValues(0, 2) = "Upgrade"
'row headers in column 0
geometryValues(1, 0) = "NumLanes"
geometryValues(2, 0) = "LaneWidth"
geometryValues(3, 0) = "MedianWidth"
geometryValues(4, 0) = "MountingHeight"
geometryValues(5, 0) = "PoleSpacing"
geometryValues(6, 0) = "PoleSetback"
geometryValues(7, 0) = "ArmLength"
geometryValues(8, 0) = "FixtureArrangement"

ExitFlag = False
With wksRoadGeometry
    'Add the baseline values to the array
    geometryValues(1, 1) = .Range("bNumLanes").Value
    geometryValues(2, 1) = .Range("bLaneWidth").Value
    geometryValues(3, 1) = .Range("bMedianWidth").Value
    geometryValues(4, 1) = .Range("bMountingHeight").Value
    geometryValues(5, 1) = .Range("bPoleSpacing").Value
    geometryValues(6, 1) = .Range("bPoleSetback").Value
    geometryValues(7, 1) = .Range("bArmLength").Value
    If VarType(.Range("bFixtureArrangement").Value) = vbError Then ExitFlag = True Else geometryValues(8, 1) = .Range("bFixtureArrangement").Value
    
    'Add the upgrade values to the array
    geometryValues(1, 2) = .Range("uNumLanes").Value
    geometryValues(2, 2) = .Range("uLaneWidth").Value
    geometryValues(3, 2) = .Range("uMedianWidth").Value
    geometryValues(4, 2) = .Range("uMountingHeight").Value
    geometryValues(5, 2) = .Range("uPoleSpacing").Value
    geometryValues(6, 2) = .Range("uPoleSetback").Value
    geometryValues(7, 2) = .Range("uArmLength").Value
    If VarType(.Range("uFixtureArrangement").Value) = vbError Then ExitFlag = True Else geometryValues(8, 2) = .Range("uFixtureArrangement").Value
End With


'FLAG this should be eliminated once fully transitioned to the geometryValues array. This section exists for compatibility with methods that have not yet been updated to handle the geometryValues array. Those subs should be refactored for simplification, and then this section can be removed.
Dim intBaselineUpgradeChoice
If Sheets("FixtureData").Range("Base_Upgrade_Choice").Value = "Baseline" Then
    intBaselineUpgradeChoice = 1
ElseIf Sheets("FixtureData").Range("Base_Upgrade_Choice").Value = "Upgrade" Then
    intBaselineUpgradeChoice = 2
End If
NumberOfLanes = geometryValues(1, intBaselineUpgradeChoice)
lanewidth = geometryValues(2, intBaselineUpgradeChoice)
MedianLength = geometryValues(3, intBaselineUpgradeChoice)
FixtureHeight = geometryValues(4, intBaselineUpgradeChoice)
polespacing = geometryValues(5, intBaselineUpgradeChoice)
polesetback = geometryValues(6, intBaselineUpgradeChoice)
ArmLength = geometryValues(7, intBaselineUpgradeChoice)
poleconfig = geometryValues(8, intBaselineUpgradeChoice)
'end section of refactor

'End sub if there is an error in the data entry for pole configuration
If ExitFlag = True Then
    ErrorMessage = Sheet25.Range("poleconfigError").Value
    MsgBox (ErrorMessage)
    Exit Sub
End If

'Determine calculation method
calcMethod = Sheets("FixtureData").Range("iescieGraphChoice").Value

Dim LLF
LLF = Sheets("FixtureData").Range("H6").Value 'light loss factor

Dim outputXY()
Dim gammaArray()
Dim phi()
Dim phiArrayForITable()
Dim gammaArrayForITable()
Dim betaArray()

'number of grid points?
Dim ngp As Integer
ngp = TotalGridLength(calcMethod, FixtureHeight, polespacing) / GridSpace(calcMethod, polespacing)

'Making the grid and getting its positions into an array
'This grid has the XY coordinates of the entire road array (regardless of wheter values are calculated at those points). Logic in the angle and illuminance/luminance calcs decides which of these to use.
outputXY = makeGrid(NumberOfLanes, calcMethod, ngp, poleconfig, MedianLength, polespacing, lanewidth)

'Output the X and Y coordinates of each individual fixture.
'X is along the road. Y is across the road.
Dim fixtureX()
Dim fixtureY()
gridlength = TotalGridLength(calcMethod, FixtureHeight, polespacing)
fixtureX = FixturePosition(NumberOfLanes, poleconfig, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)(0)
fixtureY = FixturePosition(NumberOfLanes, poleconfig, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)(1)
'**FLAG small performance speedup - FixturePosition function recalculates each time it is called

'Tilt--------------------------
tiltOnX = 0 / 180 * WorksheetFunction.Pi        'the up down tilt
tiltOnY = 0 / 180 * WorksheetFunction.Pi        'towards or away from observer, i.e. twisting the arm
tiltOnZ = 0 / 180 * WorksheetFunction.Pi        'twisting the pole


' Every array prefixed L is used for Illuminance calculations and prefixed with R is used for Luminance calculations
Dim larray()
Dim LarrayMatrix()
ReDim LarrayMatrix(UBound(fixtureX))

Dim Rarray()
Dim RarrayMatrix()
ReDim RarrayMatrix(UBound(fixtureX))

Dim tempArray1()
Dim illuminanceFixture()
ReDim illuminanceFixture(UBound(fixtureX))

Dim temparray2()
Dim luminanceFixture()
ReDim luminanceFixture(UBound(fixtureX))

Dim LthisArray()
Dim LsumArray()
Dim LnextArray()
Dim RthisArray()
Dim RsumArray()
Dim RnextArray()
Dim rownum
rownum = 13 'This is the first row on the "Luminance Calcs CIE" tab where results are written. It gets incremented as each lane gets calculated in turn.

' If IES method
If calcMethod = "IES" Then
    '******************************************************************************************************
    '*******************************if IES method**********************************************************
    '******************************************************************************************************
    ' Calculate the angles, luminous intensity, reflectance and finally luminance/illuminance at each grid point for the contributions of each fixture. Save each array of contribution of all grid points for each fixture into an array of arrays.
    'Note, outputXY is the coordinates of each grid point (included unused gridpoints).
    'Each phi, gamma, beta array is redim'd to just the x coordinates that are between the two primary fixtures (but their location from the first calculated fixture is the same position in the array as in the outputXY array)
    '
    'for instance, if some poles are before and after the primary poles and the first pole is at (9) the two arrays would look like this:
    'X directionin outputXY: (1)(2) (3) (4) (5) (6) (7) (8) (9) (10) (11) (12) (13) (14) (15) (16)
    'x direction in others :                                (9) (10) (11) (12)
    
    '------------------------------------------------------------------
    'First, calculate Luminance and Illuminance at every grid point for every included fixture; these will be summed later. k interates through each fixture
    '------------------------------------------------------------------
        For k = LBound(fixtureX) To UBound(fixtureX)
            'Angle calculations
            phi = anglePhiWithTilt(fixtureX(k), fixtureY(k), outputXY, 0, 0, 0, calcMethod, intBaselineUpgradeChoice, geometryValues()) 'For the actual angle, tilt is zero since it is the angle of the light path itself
            phiArrayForITable = anglePhiWithTilt(fixtureX(k), fixtureY(k), outputXY, tiltOnX, tiltOnY, tiltOnZ, calcMethod, intBaselineUpgradeChoice, geometryValues()) 'for the I table, tilt is used to change which angle is used for the light intensity lookup
            gammaArray = angleGammaWithTilt(fixtureX(k), fixtureY(k), outputXY, 0, 0, 0, calcMethod, intBaselineUpgradeChoice, geometryValues())
            gammaArrayForITable = angleGammaWithTilt(fixtureX(k), fixtureY(k), outputXY, tiltOnX, tiltOnY, tiltOnZ, calcMethod, intBaselineUpgradeChoice, geometryValues())
            betaArray = angleBeta(phi(), calcMethod, fixtureX(k), fixtureY(k), outputXY, polespacing, lanewidth, FixtureHeight, 0) '0 is the "yo" observer location; logic in the function does not use yo when calc method is IES, this is onyl used for CIE
            
            'Luminous intensity calculations using quadratic interpolation
            larray = LintensityMatrix(ngp, poleconfig, fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, calcMethod, phiArrayForITable, gammaArrayForITable) 'FLAG
            LarrayMatrix(k) = larray
    
            'Road reflectance using quadratic interpolation
            Rarray = RMatrix(gridlength, poleconfig, fixtureX(k), fixtureY(k), outputXY(), polespacing, FixtureHeight, calcMethod, betaArray, gammaArray)
            RarrayMatrix(k) = Rarray
            
            ' Illuminance at every grid point by fixture k
            tempArray1 = Illum(larray, gammaArray, LLF, FixtureHeight)
            illuminanceFixture(k) = tempArray1
        
            ' Luminance at every grid point by fixture k
            temparray2 = Lum(larray, gammaArray, Rarray, LLF, FixtureHeight)
            luminanceFixture(k) = temparray2
            
            '```````````````````````DEBUG ONLY````````````````````````````````
            'output all calculated values to scratch sheet for debugging
            If gbDebug = True Then
                Dim intScratchLastRow As Integer, intScratchLastCol As Integer, rScratchLastCell As Range
                Set rScratchLastCell = wksScratch.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
                If rScratchLastCell Is Nothing Then Set rScratchLastCell = wksScratch.Range("A1")
                intScratchLastRow = 1 'rScratchLastCell.row
                intScratchLastCol = rScratchLastCell.column
                Dim rr As Integer 'the variable to keep track of current row
                
                ReDim aOutput(300, 100) As Variant
                
                aOutput(1, 1) = "Fixture" & k
                aOutput(2, 1) = "X distance"
                aOutput(2, 2) = fixtureX(k)
                aOutput(3, 1) = "Y distance"
                aOutput(3, 2) = fixtureY(k)
                
                aOutput(5, 1) = "X distances"
                
                Set rTarget = wksScratch.Cells(1, intScratchLastCol)
                rTarget.Resize(UBound(aOutput, 1), UBound(aOutput, 2)) = aOutput
                                
                rr = 6
                
                'X distances
                colNow = 0
                rowNow = 1
                For ii = LBound(phi, 1) To UBound(phi, 1)
                    aOutput(rr, 1 + colNow) = outputXY(0)(ii)
                    colNow = colNow + 1
                Next ii
                rr = rr + rowNow
                
                'Phi Array
                Call printIntermediateVariables(rr, aOutput, phi)
                Call printIntermediateVariables(rr, aOutput, phiArrayForITable)
                Call printIntermediateVariables(rr, aOutput, gammaArray)
                Call printIntermediateVariables(rr, aOutput, gammaArrayForITable)
                Call printIntermediateVariables(rr, aOutput, betaArray)
                Call printIntermediateVariables(rr, aOutput, larray)
                Call printIntermediateVariables(rr, aOutput, tempArray1) 'Illuminance
                Call printIntermediateVariables(rr, aOutput, temparray2) 'Luminance
                
                Set rTarget = wksScratch.Cells(intScratchLastRow, intScratchLastCol)
                rTarget.Resize(UBound(aOutput, 1), UBound(aOutput, 2)) = aOutput
            End If
            '```````````````````````END DEBUG ONLY````````````````````````````````
        Next
    
    '-----------------------------------------------------
    'Sum up the contributions of all the relevant fixtures
    '-----------------------------------------------------
        ReDim LsumArray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))
        ReDim RsumArray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))

        
        'Add up the contributions of all fixtures
        'Get the all gridpoint values for each fixture (..thisArray) and add the values in each grid point to the corresponding gridpoint in the sumArray
        For i = LBound(illuminanceFixture) To UBound(illuminanceFixture)
            LthisArray = illuminanceFixture(i)
            RthisArray = luminanceFixture(i)
            For j = LBound(LsumArray, 1) To UBound(LsumArray, 1)
                For k = LBound(LsumArray, 2) To UBound(LsumArray, 2)
                    LsumArray(j, k) = LsumArray(j, k) + LthisArray(j, k)        'Illuminance
                    RsumArray(j, k) = RsumArray(j, k) + RthisArray(j, k)        'Luminance, grid point j and k matches between thisArray and sumArray
                Next
            Next
        Next
    
    '-----------------------------------------------------
    'Output the values to the relevant sheet
    '-----------------------------------------------------
        ' getting the values into luminance and iluminance sheets
        Sheets("Illuminance Calcs").[B13].Resize(UBound(LsumArray, 1) - LBound(LsumArray, 1) + 1, UBound(LsumArray, 2) - LBound(LsumArray, 2) + 1).Value = LsumArray
        Sheets("Luminance Calcs").[B13].Resize(UBound(RsumArray, 1) - LBound(RsumArray, 1) + 1, UBound(RsumArray, 2) - LBound(RsumArray, 2) + 1).Value = RsumArray
        
        'add the labels
        Dim t As Integer
        Dim s As Integer
        s = 2
        For t = 1 To NumberOfLanes
            Sheets("Illuminance Calcs").Cells(12, s).Value = "Lane " & t & " - 1/4 lane"
            Sheets("Illuminance Calcs").Cells(12, s + 1).Value = "Lane " & t & " - 3/4 lane"
            Sheets("Luminance Calcs").Cells(12, s).Value = "Lane " & t & " - 1/4 lane"
            Sheets("Luminance Calcs").Cells(12, s + 1).Value = "Lane " & t & " - 3/4 lane"
            s = s + 2
        Next t

'******************************************************************************************************
'*******************************if CIE method**********************************************************
'******************************************************************************************************
ElseIf calcMethod = "CIE" Then
    '-----------
    'Luminance
    '-----------
    'running for all observer locations
    For m = 1 To NumberOfLanes
        yo = (2 * m - 1) * lanewidth / 2                    'yo = y of observer, in the center of each lane in turn
        If yo > NumberOfLanes * lanewidth / 2 Then
        yo = yo + MedianLength
        End If
        
        ' Running for all fixtures
        For k = LBound(fixtureX) To UBound(fixtureX)
        
            ' running for all observer locations.
            'Different luminance for different observer location is put into a sheet called Luminance Calcs CIE
            ' Least average values and uniformity values are taken from this sheet into main luminance calcs sheet
            
            'getting angle matrices
            phi = anglePhiWithTilt(fixtureX(k), fixtureY(k), outputXY, 0, 0, 0, calcMethod, intBaselineUpgradeChoice, geometryValues())     'use zero degree tilt for the reflectance calculations
            phiArrayForITable = anglePhiWithTilt(fixtureX(k), fixtureY(k), outputXY, tiltOnX, tiltOnY, tiltOnZ, calcMethod, intBaselineUpgradeChoice, geometryValues)   'use tilt for the light intensity lookup
            gammaArray = angleGammaWithTilt(fixtureX(k), fixtureY(k), outputXY, 0, 0, 0, calcMethod, intBaselineUpgradeChoice, geometryValues)
            gammaArrayForITable = angleGammaWithTilt(fixtureX(k), fixtureY(k), outputXY, tiltOnX, tiltOnY, tiltOnZ, calcMethod, intBaselineUpgradeChoice, geometryValues)
            betaArray = angleBeta(phi(), calcMethod, fixtureX(k), fixtureY(k), outputXY, polespacing, lanewidth, FixtureHeight, yo) 'yo is not used when the calc method is IES
            
            'getting luminous intensity using quadratic interpolation
            larray = LintensityMatrix(ngp, poleconfig, fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, calcMethod, phiArrayForITable, gammaArrayForITable)
            LarrayMatrix(k) = larray
            
            'Road reflectance using quadratic interpolation
            Rarray = RMatrix(gridlength, poleconfig, fixtureX(k), fixtureY(k), outputXY(), polespacing, FixtureHeight, calcMethod, betaArray, gammaArray)
            RarrayMatrix(k) = Rarray
            
            'Calculating luminance
            temparray2 = Lum(larray, gammaArray, Rarray, LLF, FixtureHeight)
            luminanceFixture(k) = temparray2
        
        Next k  'next fixture
        
        ' sum of luminnance from all fixtures
        ReDim RsumArray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))
        
        For i = LBound(luminanceFixture) To UBound(luminanceFixture)
        RthisArray = luminanceFixture(i)
            For j = LBound(RsumArray, 1) To UBound(RsumArray, 1)
                For k = LBound(RsumArray, 2) To UBound(RsumArray, 2)
                    ' Summing up luminace at each grid point from one observer location
                    RsumArray(j, k) = RsumArray(j, k) + RthisArray(j, k)
                Next k
            Next j
        Next i
        
        'getting luminance for one oberserver location into excel sheet
        Sheets("Luminance Calcs CIE").Range("F" & rownum).Resize(UBound(RsumArray, 1) - LBound(RsumArray, 1) + 1, UBound(RsumArray, 2) - LBound(RsumArray, 2) + 1).Value = RsumArray
        
        Dim c As Integer, sc As Integer, colNo As Integer, arr As Variant
        c = UBound(RsumArray, 2) + 1        'counter used to get mid-lane gridpoint
        colNo = 53                          'used for outputting the min/max of each lane to the sheet
        arr = RsumArray                     'arr is used to get all the grid points in one direction
        ReDim Preserve arr(LBound(RsumArray, 1) To UBound(RsumArray, 1), 1)     'FLAG I think that the Preserve can be removed, and the arr = rSumArray can also be removed, since the values for arr get assigned in the for loop below.
        
        '-------------------------------------
        'Find the summary statistics for this observer location (min, max, uniformity, etc.)
        '-------------------------------------
        Dim currentLaneAvg As Double, currentLaneMin As Double, currentLaneMax As Double, currentLaneMinAvgUniformity As Double, minRatio As Double
        
        'Calculate simple summary stats
        currentLaneAvg = Application.WorksheetFunction.Average(RsumArray)
        currentLaneMin = Application.WorksheetFunction.Min(RsumArray)
        currentLaneMax = Application.WorksheetFunction.max(RsumArray)
        If currentLaneMin <> 0 Then currentLaneMinAvgUniformity = currentLaneMin / currentLaneAvg Else currentLaneMinAvgUniformity = 0
        
        'Longitudinal uniformity is calculated along the centerline of each lane, rather than across the whole grid, and the lowest uniformity is used.
        'FLAG per section 8.3 this should only be done for the particular lane that the observer is in, because it says the observer should be in line with the midpoint. "for sc..." should instead be to determine which lane should be picked, and the overall max calc needs to be rethought.
        For sc = 2 To c - 1 Step 3
            Dim currentmax As Double, currentmin As Double, currentratio As Double
            arr = Application.WorksheetFunction.Index(RsumArray, 0, sc)                                 'get all grid points for one row, the middle of current lane (sc)
            currentmin = Application.WorksheetFunction.Min(arr)                                         'Min of this lane to sheet
            currentmax = Application.WorksheetFunction.max(arr)                                         'Max of this lane to sheet
            If currentmax <> 0 Then currentratio = currentmin / currentmax Else currentratio = 0        'Avoid the div0 error
            
            'identify lowest uniformity of all lanes
            If overallmax Is Nothing Then overallmax = currentmax           'the first time through
            If currentmax < minRatio Then minRatio = currentratio         'identifies the overall highest ratio by the time the for loop is over
            
            'Output data to the sheet (FLAG this could potentially be deleted, as long as it's not used elsewhere)
            Sheets("Luminance Calcs CIE").Cells(rownum, colNo).Value = currentmin
            Sheets("Luminance Calcs CIE").Cells(rownum + 1, colNo).Value = currentmax
            Sheets("Luminance Calcs CIE").Cells(rownum + 2, colNo).Value = currentratio
            
            colNo = colNo + 1
        Next sc

        'Output the summary stats to the sheet
        Sheets("Luminance Calcs CIE").Range("a" & rownum).Value = currentLaneAvg
        Sheets("Luminance Calcs CIE").Range("b" & rownum).Value = currentLaneMin
        Sheets("Luminance Calcs CIE").Range("c" & rownum).Value = currentLaneMax
        Sheets("Luminance Calcs CIE").Range("d" & rownum).Value = currentLaneMinAvgUniformity
        Sheets("Luminance Calcs CIE").Range("e" & rownum).Value = minRatio      'the lowest across all midpoints (this might be removed because it looks like only one lane should be used per observer location, see FLAG above)
        
        
        rownum = rownum + UBound(RsumArray, 1) - LBound(RsumArray, 1) + 1           'Values for the next lane get placed after all grid points for current lane
    Next m 'next lane
    
    'Add section headers to the sheet
    Dim t1 As Integer
    Dim s1 As Integer
    s1 = 6
    For t1 = 1 To NumberOfLanes
        Sheets("Luminance Calcs CIE").Cells(12, s1).Value = "L " & t1 & " - 1/6"
        Sheets("Luminance Calcs CIE").Cells(12, s1 + 1).Value = "L " & t1 & " - 3/6"
        Sheets("Luminance Calcs CIE").Cells(12, s1 + 2).Value = "L " & t1 & " - 5/6"
        s1 = s1 + 3
    Next t1
    
    '-----------
    'Illuminance
    '-----------
    'Grid changed to take only 3 longitudnal points across road instead of lane.
    LLF = Sheets("FixtureData").Range("H6").Value
    
    If MedianLength = 0 Then
        Dim outputY1(2)
        outputY1(0) = (lanewidth * NumberOfLanes + MedianLength) / 6
        outputY1(1) = (lanewidth * NumberOfLanes + MedianLength) / 2
        outputY1(2) = (lanewidth * NumberOfLanes + MedianLength) * 5 / 6
        outputXY(1) = outputY1
    Else
        Dim outputY2(5)
        outputY2(0) = (lanewidth * NumberOfLanes / 2) / 6
        outputY2(1) = (lanewidth * NumberOfLanes / 2) / 2
        outputY2(2) = (lanewidth * NumberOfLanes / 2) * 5 / 6
        outputY2(3) = (lanewidth * NumberOfLanes / 2) / 6 + MedianLength + (lanewidth * NumberOfLanes / 2)
        outputY2(4) = (lanewidth * NumberOfLanes / 2) / 2 + MedianLength + (lanewidth * NumberOfLanes / 2)
        outputY2(5) = (lanewidth * NumberOfLanes / 2) * 5 / 6 + MedianLength + (lanewidth * NumberOfLanes / 2)
        outputXY(1) = outputY2
    End If
    
    ' same routine again, but a different grid. FLAG this needs to be updated to work with tilt
    For k = LBound(fixtureX) To UBound(fixtureX)
        phi = anglePhiWithTilt(fixtureX(k), fixtureY(k), outputXY, 0, 0, 0, calcMethod, intBaselineUpgradeChoice, geometryValues())
        gammaArray = angleGammaWithTilt(fixtureX(k), fixtureY(k), outputXY, 0, 0, 0, calcMethod, intBaselineUpgradeChoice, geometryValues)
        betaArray = angleBeta(phi(), calcMethod, fixtureX(k), fixtureY(k), outputXY, polespacing, lanewidth, FixtureHeight, 0)
        
        'removing all the luminaries outside 5H distance
        If outputXY(0)(LBound(gammaArray)) - fixtureX(k) > 5 * FixtureHeight Or fixtureX(k) - outputXY(0)(UBound(gammaArray)) > 5 * FixtureHeight Then
            LLF = 0
        Else
            LLF = Sheets("FixtureData").Range("H6").Value
        End If
        
        larray = LintensityMatrix(ngp, poleconfig, fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, calcMethod, phi, gammaArray)
        LarrayMatrix(k) = larray
        
        tempArray1 = Illum(larray, gammaArray, LLF, FixtureHeight)
        illuminanceFixture(k) = tempArray1
    Next
    
    '*************************** got luminance and illuminance matrices at every grid point by every fixture till now *********
    ' sum of luminnance from all fixtures
    ReDim LsumArray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))
    For p = LBound(gammaArray, 1) To UBound(gammaArray, 1)
        For q = LBound(gammaArray, 2) To UBound(gammaArray, 2)
            LsumArray(p, q) = 0
            'RsumArray(p, q) = 0
        Next q
    Next p
    
    For i = LBound(illuminanceFixture) To UBound(illuminanceFixture)
    LthisArray = illuminanceFixture(i)
        For j = LBound(LsumArray, 1) To UBound(LsumArray, 1)
            For k = LBound(LsumArray, 2) To UBound(LsumArray, 2)
                    LsumArray(j, k) = LsumArray(j, k) + LthisArray(j, k)
            Next
        Next
    Next
    
    Sheets("Illuminance Calcs").[B13].Resize(UBound(LsumArray, 1) - LBound(LsumArray, 1) + 1, UBound(LsumArray, 2) - LBound(LsumArray, 2) + 1).Value = LsumArray
    'Sheets("Luminance Calcs").[B13].Resize(UBound(RsumArray, 1) - LBound(RsumArray, 1) + 1, UBound(RsumArray, 2) - LBound(RsumArray, 2) + 1).Value = RsumArray
        Sheets("Illuminance Calcs").Cells(12, 2).Value = "R - 1/6"
        Sheets("Illuminance Calcs").Cells(12, 3).Value = "R - 3/6"
        Sheets("Illuminance Calcs").Cells(12, 4).Value = "R - 5/6"
        
    If MedianLength > 0 Then
        Sheets("Illuminance Calcs").Cells(12, 5).Value = "R - 1/6"
        Sheets("Illuminance Calcs").Cells(12, 6).Value = "R - 3/6"
        Sheets("Illuminance Calcs").Cells(12, 7).Value = "R - 5/6"
    End If
    
End If  'CIE vs. IES
End Sub

Sub PrintArray(Data(), Cl As Range)
    Cl.Resize(UBound(Data, 1), UBound(Data, 2) + 1) = Data
End Sub


Function getDimension(var As Variant) As Integer
On Error GoTo Err:
    Dim i As Integer
    Dim tmp As Integer
    i = 0
    Do While True:
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    getDimension = i - 1
End Function

Function Illum(larray(), gammaArray(), LLF, FixtureHeight) As Variant
Dim tarray()
ReDim tarray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))

For i = LBound(gammaArray(), 1) To UBound(gammaArray(), 1)
    For j = LBound(gammaArray(), 2) To UBound(gammaArray(), 2)

        tarray(i, j) = larray(i, j) * Cos(gammaArray(i, j) * WorksheetFunction.Pi / 180) ^ 3 * LLF / FixtureHeight ^ 2
        'MsgBox tarray(i, j)
    Next
Next

Illum = tarray

End Function
Function Lum(larray(), gammaArray(), Rarray(), LLF, FixtureHeight) As Variant
Dim tarray()
ReDim tarray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))

For i = LBound(gammaArray(), 1) To UBound(gammaArray(), 1)
    For j = LBound(gammaArray(), 2) To UBound(gammaArray(), 2)

        tarray(i, j) = larray(i, j) * Rarray(i, j) * LLF / ((FixtureHeight ^ 2) * 10000)
    Next
Next

Lum = tarray

End Function

Sub printIntermediateVariables(rr As Integer, aOutput() As Variant, aArray() As Variant)
Dim colNow As Integer, rowNow As Integer
            colNow = 0
            rowNow = 0
            For ii = LBound(aArray, 1) To UBound(aArray, 1)
                'aOutput(rr, 1 + colNow) = outputXY(0)(ii)    'phi(ii, 1)
                rowNow = 1
                For jj = LBound(aArray, 2) To UBound(aArray, 2)
                    aOutput(rr + rowNow, 1 + colNow) = aArray(ii, jj)
                    rowNow = rowNow + 1
                Next jj
                colNow = colNow + 1
            Next ii
            
            rr = rr + rowNow
End Sub


