Attribute VB_Name = "AngleCalculations"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

Function angleGamma(fixtureX, fixtureY, gridXY, polespacing, FixtureHeight, angleOnX, angleOnY, angleOnZ, calculationmethod) As Variant
Dim phi() As Integer

'arrays to carry all x and y values
Dim outputX()
Dim outputY()
outputX = gridXY(0)
outputY = gridXY(1)

'grid start and end
If calculationmethod = "IES" Then
    istart = WorksheetFunction.Match(polespacing, outputX, True)
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
    Debug.Print "iStart in angleGamma is " & istart
    Debug.Print "iEnd in angleGamma is " & iend
    'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
    'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
End If

'X is along the road. Y is across the road.
numberOfX = iend - istart
numberOfY = UBound(outputY)

Dim gammaArray()
ReDim gammaArray(istart To iend, numberOfY)
m = outputX(1)
For i = istart To iend
For j = 0 To numberOfY
    distY = fixtureY                     'grid measurement including tilt
    dist = Distance(fixtureX, distY, outputX(i), outputY(j))
    'gamma at each grid point
    If dist <> 0 Then
        gammaArray(i, j) = (Atn(dist / FixtureHeight)) * 180 / WorksheetFunction.Pi
    Else
        gammaArray(i, j) = 0
    End If
Next
Next
angleGamma = gammaArray
End Function

Function angleGammaWithTilt(fixtureX, fixtureY, gridXY, polespacing, FixtureHeight, tiltOnX, tiltOnY, tiltOnZ, calculationmethod) As Variant
'gamma is the vertical angle coming out of the fixture
Dim phi() As Integer

'arrays to carry all x and y values - contains grid points
Dim outputX()
Dim outputY()
outputX = gridXY(0)
outputY = gridXY(1)

'grid start and end
If calculationmethod = "IES" Then
    istart = WorksheetFunction.Match(polespacing, outputX, True)
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
End If

'inputs needed for calculating distances. FLAG spreadsheet calls need to be removed from this function.
If Sheets("FixtureData").Range("A6").Value = "Baseline" Then
    lanewidth = Sheets("Road Geometry").Range("bLaneWidth").Value
    MedianLength = Sheets("Road Geometry").Range("bMedianWidth").Value
    NumberOfLanes = Sheets("Road Geometry").Range("bNumLanes").Value
    poleconfig = Sheets("Road Geometry").Range("bFixtureArrangement").Value
ElseIf Sheets("FixtureData").Range("A6").Value = "Upgrade" Then
    lanewidth = Sheets("Road Geometry").Range("uLaneWidth").Value
    MedianLength = Sheets("Road Geometry").Range("uMedianWidth").Value
    NumberOfLanes = Sheets("Road Geometry").Range("uNumLanes").Value
    poleconfig = Sheets("Road Geometry").Range("uFixtureArrangement").Value
End If


'X is along the road. Y is across the road.
numberOfX = iend - istart
numberOfY = UBound(outputY)

Dim gammaArray()
ReDim gammaArray(istart To iend, numberOfY)

'for debug purposes
Dim xArray()
ReDim xArray(istart To iend, numberOfY)
Dim yArray()
ReDim yArray(istart To iend, numberOfY)
Dim xPrimeArray()
ReDim xPrimeArray(istart To iend, numberOfY)
Dim yPrimeArray()
ReDim yPrimeArray(istart To iend, numberOfY)
Dim hPrimeArray()
ReDim hPrimeArray(istart To iend, numberOfY)


m = outputX(1)
For i = istart To iend
    For j = 0 To numberOfY
        '    v = tiltOnZ
        '    w = tiltOnY
        '    o = tiltOnX
        Dim x As Double, y As Double
        'FLAG - version 0. This is the one in the
'        x = (outputX(i) - fixtureX)
'        y = (outputY(j) - fixtureY)
        
        'FLAG - version 1. In this one, results for opposite street sides are not opposite (19.7 vers 24.8 for first grid point)
        x = (outputX(i) - fixtureX)
        If fixtureY > (lanewidth * NumberOfLanes / 2) Then  'if the fixture is located on the far side of the road
            y = (fixtureY - outputY(j))
        Else
            y = (outputY(j) - fixtureY)
        End If
        
        'FLAG - version 2. Produces the same results as version 1 for the two near fixtres, but traded across the road.
        'y = (fixtureY - outputY(j)) 'testing if it fixes it to overwrite
        
        xPrime = x * (Cos(tiltOnZ) * Cos(tiltOnY) - Sin(tiltOnZ) * Sin(tiltOnX) * Sin(tiltOnY)) + _
                 y * (Sin(tiltOnZ) * Cos(tiltOnY) + Cos(tiltOnZ) * Sin(tiltOnX) * Sin(tiltOnY)) + _
                 FixtureHeight * Cos(tiltOnX) * Sin(tiltOnY)
        yPrime = -x * Sin(tiltOnZ) * Cos(tiltOnX) + _
                    y * Cos(tiltOnZ) * Cos(tiltOnX) - _
                    FixtureHeight * Sin(tiltOnX)
        HPrime = -x * (Sin(tiltOnZ) * Sin(tiltOnX) * Cos(tiltOnY) + Cos(tiltOnZ) * Sin(tiltOnY)) - _
                 y * (Sin(tiltOnZ) * Sin(tiltOnY) - Cos(tiltOnZ) * Sin(tiltOnX) * Cos(tiltOnY)) + _
                 FixtureHeight * Cos(tiltOnX) * Cos(tiltOnY)
                    
        gammaTemp = Atn(((xPrime ^ 2 + yPrime ^ 2) ^ 0.5) / HPrime) * 180 / WorksheetFunction.Pi
        gammaArray(i, j) = gammaTemp
        
        
        'saving for debug purposes only
        xArray(i, j) = x
        yArray(i, j) = y
        xPrimeArray(i, j) = xPrime
        yPrimeArray(i, j) = yPrime
        hPrimeArray(i, j) = HPrime
            
        '    distY = fixtureY
        '    dist = Distance(fixtureX, distY, outputX(i), outputY(j))
        '    'gamma at each grid point
        '    If dist <> 0 Then
        '        gammaArray(i, j) = (Atn(dist / FixtureHeight)) * 180 / WorksheetFunction.Pi
        '    Else
        '        gammaArray(i, j) = 0
        '    End If
    Next
Next

'Temp code for debugging purposes
'If fixtureX = 35 Then 'Only worry about writing the two opposing fixtures
'    Dim rrow As Integer
'    rrow = 1
'    Dim aOutputGamma()
'    ReDim aOutput(300, 300) As Variant
'
'    Call printIntermediateVariables(rrow, aOutputGamma, xArray)
'    Call printIntermediateVariables(rrow, aOutputGamma, yArray)
'    Call printIntermediateVariables(rrow, aOutputGamma, xPrimeArray)
'    Call printIntermediateVariables(rrow, aOutputGamma, yPrimeArray)
'    Call printIntermediateVariables(rrow, aOutputGamma, hPrimeArray)
'
'    If fixtureY = 0.5 Then yColumn = 22
'    If fixtureY = 16.5 Then yColumn = 32
'    Set rTarget = wksScratch.Cells(100, yColumn)
'    rTarget.Resize(UBound(aOutput, 1), UBound(aOutput, 2)) = aOutput
'
'End If

angleGammaWithTilt = gammaArray

End Function

Function Distance(x1, y1, x2, y2)
Distance = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
End Function

Function anglePhi(fixtureX, fixtureY, gridXY(), polespacing, FixtureHeight, angleOnX, angleOnY, angleOnZ, calculationmethod) As Variant
'arrays to carry all x and y values
Dim outputX()
Dim outputY()
outputX = gridXY(0)
outputY = gridXY(1)

'inputs
If Sheets("FixtureData").Range("A6").Value = "Baseline" Then
    lanewidth = Sheets("Road Geometry").Range("bLaneWidth").Value
    MedianLength = Sheets("Road Geometry").Range("bMedianWidth").Value
    NumberOfLanes = Sheets("Road Geometry").Range("bNumLanes").Value
    poleconfig = Sheets("Road Geometry").Range("bFixtureArrangement").Value
ElseIf Sheets("FixtureData").Range("A6").Value = "Upgrade" Then
    lanewidth = Sheets("Road Geometry").Range("uLaneWidth").Value
    MedianLength = Sheets("Road Geometry").Range("uMedianWidth").Value
    NumberOfLanes = Sheets("Road Geometry").Range("uNumLanes").Value
    poleconfig = Sheets("Road Geometry").Range("uFixtureArrangement").Value
End If

'grid start and end
If calculationmethod = "IES" Then
    istart = WorksheetFunction.Match(polespacing, outputX, True)
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
    'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
    
    Debug.Print "iStart in anglePhi is " & istart
    Debug.Print "iEnd in anglePhi is " & iend
    'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
    'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
    '**FLAG** when testing CIE make sure this is correct
End If

numberOfX = iend - istart
numberOfY = UBound(outputY)
Dim phiArray()
ReDim phiArray(istart To iend, numberOfY)
m = outputX(1)
For i = istart To iend
For j = 0 To numberOfY
'distance between grid point and fixture point
'dist = Distance(fixtureX, fixtureY, outputX(i), outputY(j))
    
    ' if pole configuration is median mounted
    If poleconfig = "Median mounted" Then
        distY = fixtureY
        If fixtureY > (lanewidth * NumberOfLanes + MedianLength) / 2 Then           'pole is located on far side of road
                If distY - outputY(j) > 0 Then
                   phiArray(i, j) = 180 - Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi
                ElseIf distY - outputY(j) < 0 Then
                   phiArray(i, j) = (Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi)
                Else
                    phiArray(i, j) = 90
                End If
        ElseIf fixtureY < lanewidth * NumberOfLanes + MedianLength / 2 Then         'pole is located on near side of road
                If distY - outputY(j) > 0 Then
                   phiArray(i, j) = Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi
                ElseIf distY - outputY(j) < 0 Then
                   phiArray(i, j) = 180 - (Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi)
                Else
                    phiArray(i, j) = 90
                End If
        End If
    Else
    'For all configurations except median mounted
            distY = fixtureY                       'grid measurement including tilt
            x = (outputX(i) - fixtureX)
            If fixtureY < (lanewidth * NumberOfLanes / 2) Then  'if the fixture is located on the near side of the road
                y = (outputY(j) - fixtureY)
            Else
                y = (fixtureY - outputY(j))
            End If
            
            If y > 0 Then
                phiTemp = Atn(Abs(x) / Abs(y)) * 180 / WorksheetFunction.Pi 'FLAG changed to distY
            'FLAG check which version the section below was originally commented out in. This is needed for accurate calculations (comment inserted in 1.7.6)
            ElseIf fixtureY - outputY(j) < 0 Then
               phiTemp = 180 - (Atn(Abs(x) / Abs(y)) * 180 / WorksheetFunction.Pi)
            Else
                phiTemp = 90
            End If
            phiArray(i, j) = phiTemp
    End If

Next j
Next i
anglePhi = phiArray

End Function
Function anglePhiWithTilt(fixtureX, fixtureY, gridXY(), polespacing, FixtureHeight, tiltOnX, tiltOnY, tiltOnZ, calculationmethod) As Variant
'arrays to carry all x and y values
Dim outputX()
Dim outputY()
outputX = gridXY(0)
outputY = gridXY(1)

'inputs
If Sheets("FixtureData").Range("A6").Value = "Baseline" Then
    lanewidth = Sheets("Road Geometry").Range("bLaneWidth").Value
    MedianLength = Sheets("Road Geometry").Range("bMedianWidth").Value
    NumberOfLanes = Sheets("Road Geometry").Range("bNumLanes").Value
    poleconfig = Sheets("Road Geometry").Range("bFixtureArrangement").Value
ElseIf Sheets("FixtureData").Range("A6").Value = "Upgrade" Then
    lanewidth = Sheets("Road Geometry").Range("uLaneWidth").Value
    MedianLength = Sheets("Road Geometry").Range("uMedianWidth").Value
    NumberOfLanes = Sheets("Road Geometry").Range("uNumLanes").Value
    poleconfig = Sheets("Road Geometry").Range("uFixtureArrangement").Value
End If

'grid start and end
If calculationmethod = "IES" Then
    istart = WorksheetFunction.Match(polespacing, outputX, True)
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
    'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
    'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
    'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
    '**FLAG** when testing CIE make sure this is correct
End If

numberOfX = iend - istart
numberOfY = UBound(outputY)
Dim phiArray()
ReDim phiArray(istart To iend, numberOfY)
m = outputX(1)
For i = istart To iend
For j = 0 To numberOfY
'distance between grid point and fixture point
'dist = Distance(fixtureX, fixtureY, outputX(i), outputY(j))
    
    ' if pole configuration is median mounted
    If poleconfig = "Median mounted" Then
        distY = fixtureY
        If fixtureY > (lanewidth * NumberOfLanes + MedianLength) / 2 Then           'pole is located on far side of road
                If distY - outputY(j) > 0 Then
                   phiArray(i, j) = 180 - Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi
                ElseIf distY - outputY(j) < 0 Then
                   phiArray(i, j) = (Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi)
                Else
                    phiArray(i, j) = 90
                End If
        ElseIf fixtureY < lanewidth * NumberOfLanes + MedianLength / 2 Then         'pole is located on near side of road
                If distY - outputY(j) > 0 Then
                   phiArray(i, j) = Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi
                ElseIf distY - outputY(j) < 0 Then
                   phiArray(i, j) = 180 - (Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi)
                Else
                    phiArray(i, j) = 90
                End If
        End If
    Else
    'For all configurations except median mounted
            '    v = tiltOnZ
            '    w = tiltOnY
            '    o = tiltOnX
            Dim x As Double, y As Double
            'Get distances for calculating angle
            x = (outputX(i) - fixtureX)
            If fixtureY < (lanewidth * NumberOfLanes / 2) Then  'if the fixture is located on the near side of the road
                y = (outputY(j) - fixtureY)
            Else
                y = (fixtureY - outputY(j))
            End If
            
            'Calculate angle including tilt
            xPrime = x * (Cos(tiltOnZ) * Cos(tiltOnY) - Sin(tiltOnZ) * Sin(tiltOnX) * Sin(tiltOnY)) + _
                     y * (Sin(tiltOnZ) * Cos(tiltOnY) + Cos(tiltOnZ) * Sin(tiltOnX) * Sin(tiltOnY)) + _
                     FixtureHeight * Cos(tiltOnX) * Sin(tiltOnY)
            yPrime = -x * Sin(tiltOnZ) * Cos(tiltOnX) + _
                        y * Cos(tiltOnZ) * Cos(tiltOnX) - _
                        FixtureHeight * Sin(tiltOnX)
            HPrime = -x * (Sin(tiltOnZ) * Sin(tiltOnX) * Cos(tiltOnY) + Cos(tiltOnZ) * Sin(tiltOnY)) - _
                    y * (Sin(tiltOnZ) * Sin(tiltOnY) - Cos(tiltOnZ) * Sin(tiltOnX) * Cos(tiltOnY)) + _
                    FixtureHeight * Cos(tiltOnX) * Cos(tiltOnY)
            
            phiTemp = Atn(Abs(xPrime) / Abs(yPrime)) * 180 / WorksheetFunction.Pi     '
            
            'convert the angle if it is located behind the fixture
            If yPrime < 0 Then
                phiTemp = 180 - phiTemp
            Else
                phiTemp = phiTemp
            End If
'
'            If phiTemp >= 0 And phiTemp <= 90 Then
'                phiArray(i, j) = phiTemp
'            ElseIf phiTemp > 90 And phiTemp <= 180 Then
'                phiArray(i, j) = 180 - phiTemp
'            ElseIf phiTemp > 180 And phiTemp <= 270 Then
'                phiArray(i, j) = 540 - phiTemp
'            ElseIf phiTemp > 270 And phiTemp < 360 Then
'                phiArray(i, j) = phiTemp
'            ElseIf phiTemp = 360 Then
'                phiArray(i, j) = 0
'            End If
            
            phiArray(i, j) = phiTemp
            'FLAG - change to deal with varying angles

'            distY = fixtureY                       'grid measurement including tilt
'            If distY - outputY(j) <> 0 Then
'               phiArray(i, j) = Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi 'FLAG changed to distY
'            'ElseIf fixtureY - outputY(j) < 0 Then
'               'phiArray(i, j) = 180 - (Atn(Abs(fixtureX - outputX(i)) / Abs(fixtureY - outputY(j))) * 180 / WorksheetFunction.Pi)
'            Else
'                phiArray(i, j) = 90
'            End If
    End If

Next j
Next i
anglePhiWithTilt = phiArray

End Function

Function angleBeta(anglePhi(), calculationmethod As String, fixtureX, fixtureY, gridXY(), polespacing, lanewidth, FixtureHeight, yo) As Variant
'yo is the y of the observer, placed in the center of each lane in turn.
Dim xo
Dim outputX()
Dim outputY()
outputX = gridXY(0)     'the location of the grid points
outputY = gridXY(1)     'the location of the grid points

'grid start and end
If calculationmethod = "IES" Then
    istart = WorksheetFunction.Match(polespacing, outputX, True)
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
    'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
    'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
    'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
End If

numberOfX = iend - istart
numberOfY = UBound(outputY)

Dim betaArray()
ReDim betaArray(istart To iend, numberOfY)
'This is for IES method
If calculationmethod = "IES" Then
For i = istart To iend
For j = 0 To numberOfY
    If fixtureX - outputX(i) > 0 Then
        If anglePhi(i, j) >= 90 Then
            betaArray(i, j) = anglePhi(i, j) - 90                                   'agree
        Else
            betaArray(i, j) = 90 - anglePhi(i, j) '90 degrees=1.57079633radians     'agree
        End If
    Else
        If anglePhi(i, j) >= 90 Then
            betaArray(i, j) = 270 - anglePhi(i, j)                                  'this is between 90 and 180 - is this how Beta was designed?
        Else
            betaArray(i, j) = 90 + anglePhi(i, j) '90 degrees=1.57079633radians     'agree
        End If
    End If
    Next
    Next
    '  if CIE method
ElseIf calculationmethod = "CIE" Then
    For i = istart To iend
    For j = 0 To numberOfY
        xo = outputX(i) - 60
        m1 = (outputY(j) - yo) / (outputX(i) - xo)
        If (outputX(i) - fixtureX) = 0 Then
            m2 = 10000000 ' a very high value in place of infinity
        Else
            m2 = (outputY(j) - fixtureY) / (outputX(i) - fixtureX)
        End If
        
        If m1 * m2 = -1 Then
            betaArray(i, j) = 90
        Else
            If m2 >= 10000000 Then
                betaArray(i, j) = 180
            Else
                If fixtureX - outputX(i) > 0 Then
                    betaArray(i, j) = Atn(Abs(m1 - m2 / 1 + m1 * m2)) * (180 / WorksheetFunction.Pi)
                Else
                    betaArray(i, j) = 180 - Atn(Abs(m1 - m2 / 1 + m1 * m2)) * (180 / WorksheetFunction.Pi)
                End If
            End If
        End If
Next
Next
Else
End If
angleBeta = betaArray
End Function
Sub drawtest()

' to plot points of measurement grid. Not being used anywhere in final version.
'This was used to check angle calcualtion of gamma, phi and beta at every grid point.

Dim ngp As Integer
Dim lanewidth As Integer, MedianLength As Integer, FixtureHeight As Integer, NumberOfLanes As Integer, polespacing As Integer
lanewidth = Sheets("Road Geometry").Range("bLaneWidth").Value
MedianLength = Sheets("Road Geometry").Range("bMedianWidth").Value
FixtureHeight = Sheets("Road Geometry").Range("bMountingHeight").Value
NumberOfLanes = Sheets("Road Geometry").Range("bNumLanes").Value
polespacing = Sheets("Road Geometry").Range("bPoleSpacing").Value

Dim polesetback As Integer, ArmLength As Integer
polesetback = Sheets("Road Geometry").Range("bPoleSetback").Value
ArmLength = Sheets("Road Geometry").Range("bArmLength").Value

Dim calcMethod As String, poleconfig As String, dataSheet As String, chartName As String
calcMethod = Sheets("FixtureData").Range("iescieGraphChoice").Value
poleconfig = Sheets("Road Geometry").Range("bFixtureArrangement").Value
dataSheet = "Chart Data Baseline"
chartName = "Baseline"

ngp = TotalGridLength(calcMethod, FixtureHeight, polespacing) / GridSpace(calcMethod, polespacing)
Dim outputXY()
Dim gammaArray()
Dim phi()

outputXY = makeGrid(NumberOfLanes, calcMethod, ngp, poleconfig, MedianLength, polespacing, lanewidth)
'outputY = makeGrid(NumberOfLanes, CalculationMethod, ngp, poleConfig, MedianLength, polespacing, lanewidth)(1)
gammaArray = anglePhi(75, -1, outputXY, polespacing, FixtureHeight, calcMethod)
'gammaArray = angleBeta(phi(), calcMethod, 0, 15, outputXY, polespacing, lanewidth, FixtureHeight)
'gammaArray = angleGamma(75, -1, outputXY, polespacing, FixtureHeight, calcMethod)
Sheets("test1").Activate
For i = LBound(gammaArray(), 1) To UBound(gammaArray(), 1)
For j = 0 To UBound(gammaArray(), 2)
Cells(i + 101, j + 1) = gammaArray(i, j)
Next
Next

End Sub







