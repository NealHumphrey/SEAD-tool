Attribute VB_Name = "FixturePositions"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

Sub drawFixtures(dataSheet As String, lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing, polesetback, ArmLength, calculationmethod As String, PoleConfiguration As String, gridlength)

' to draw fixtures on road geometry page

Dim outputX() As Variant
Dim outputY() As Variant
outputX = FixturePosition(NumberOfLanes, PoleConfiguration, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)(0)
outputY = FixturePosition(NumberOfLanes, PoleConfiguration, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)(1)


For i = 0 To UBound(outputX)

Sheets(dataSheet).Cells(i + 2, 84) = outputX(i)
Sheets(dataSheet).Cells(i + 2, 85) = outputY(i)

Next
End Sub
Sub FixturePosition(NumberOfLanes, PoleConfiguration As String, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength, selectedFixturesPerPole, tiltOnX, tiltOnY, selectedSeparationAngleRadians, fixturePositions)

' to get positions of fixtures for different pole configurations
'Output of this function is an array with the x and y positions of each pole
'For arrangements with poles on both sides of the road, fixtures in the array alternate sides of the road

Dim FPArrayX(), FPArrayY(), FPArrayBackwards(), FPArrayTiltZ()

numberoffixtures = (gridlength / polespacing) + 1                     'default starting point, this is a single-side configuration
numberoffixtures = numberoffixtures * selectedFixturesPerPole       'if there are multiple fixtures per pole
If PoleConfiguration <> "Single-side" Then numberoffixtures = 2 * numberoffixtures          'all scenarios except 'single-side' have double the number of fixtures. FLAG this might be different than before in the staggered configuration because it might include one extra fixture....


'Array with the X and Y coordinates of each pole
ReDim FPArrayX(1 To CInt(numberoffixtures))
ReDim FPArrayY(1 To CInt(numberoffixtures))
ReDim FPArrayBackwards(1 To CInt(numberoffixtures))
ReDim FPArrayTiltZ(1 To CInt(numberoffixtures))

'adjustment factors from where the pole should be to where the fixture is actually located
adjX_arm = ArmLength * Sin(selectedSeparationAngleRadians / 2)                  'assumes symetric distribution when there are two fixtures on a pole. Would also work for 3 fixtures if other logic used adjustment correctly
adjY_arm = ArmLength * Cos(selectedSeparationAngleRadians / 2)

Dim tempX As Single, tempY As Single
Dim poleSide As Integer, fixtureOnPole As Integer, fixtureCount As Integer, polePhase As Integer

'Define how to interpret the poleconfiguration for assigning fixtures
Select Case PoleConfiguration
    Case "Single-side"
        numPoleSides = 1          'Number of poles in each phase, i.e. number of sides of the street
        adjX_sideBPole = 0        'Extra X distance added to the second pole in a phase. if both poles in the same phase are aligned, this is zero.
        adjY_median = 0           'this would be 1 (true) if it were median mounted
    Case "Opposite"
        numPoleSides = 2          'Number of poles in each phase, i.e. number of sides of the street
        adjX_sideBPole = 0        'Extra X distance added to the second pole in a phase. if both poles in the same phase are aligned, this is zero.
        adjY_median = 0           'this would be 1 (true) if it were median mounted
    Case "Median mounted"
        numPoleSides = 2          'Number of poles in each phase, i.e. number of sides of the street
        adjX_sideBPole = 0        'Extra X distance added to the second pole in a phase. if both poles in the same phase are aligned, this is zero.
        adjY_median = 1           'this would be 1 (true) if it were median mounted
    Case "Staggered"
        numPoleSides = 2                    'Number of poles in each phase, i.e. number of sides of the street
        adjX_sideBPole = polespacing / 2    'Extra X distance added to the second pole in a phase. if both poles in the same phase are aligned, this is zero.
        adjY_median = 0                     'this would be 1 (true) if it were median mounted
End Select

Dim facesBackwards As Boolean, tempTiltZ As Single

'Calculate the distances
polePhase = 0   'initialize
fixtureCount = 1
Do While fixtureCount <= UBound(FPArrayX)
    For poleSide = 1 To numPoleSides
        For fixtureOnPole = 1 To selectedFixturesPerPole
            tempX = polePhase * polespacing                         'starting point for the phase is one polespacing length.
            tempX = tempX + (poleSide - 1) * adjX_sideBPole         'Accounts for staggering. If poleSide = 1, add zero; if poleSide = 2, add the adjX_sidBPole amount. Accounts for staggering
            tempX = tempX + ((-1) ^ fixtureOnPole) * adjX_arm       'if fixtureOnPole = 1, subtract adjX_arm. if fixtureOnPole = 2, add the adjX_arm. Does nothing if adjX_arm = 0
            
            tempY = 0                                                                                   'start out on side A
            tempY = tempY - ((-1) ^ poleSide) * (1 - 2 * adjY_median) * adjY_arm                        'Deals with the arm, and whether it should be added or subtracted. if poleSide = 1 and adj_Median = 1, add the amount. if poleside=2 and adj_median = -1, they cancel out (again add the amount). Otherwise subtract.
            tempY = tempY + ((-1) ^ poleSide) * (1 - adjY_median) * polesetback                         'Deals with setback, and whether it should be added or subtracted. If median mounted, setback is zero (ignored)
            tempY = tempY + (poleSide - 1) * (lanewidth * NumberOfLanes + MedianLength) * (1 - adjY_median)     'Adds the road width if on B side. Does nothing if median mounted
            tempY = tempY + (adjY_median) * (lanewidth * NumberOfLanes + MedianLength) / 2              'Adds half the road+median width. Only occurs when median mounted
            
            facesBackwards = False 'initialize
            If poleSide = 1 And adjY_median = 1 Then facesBackwards = True                  'on side A, backwards only if a median
            If poleSide = 2 And adjY_median = 0 Then facesBackwards = True                  'on side B, backwards unless there is a median
            
            tempTiltZ = selectedSeparationAngleRadians
            If fixtureOnPole = 1 Then tempTiltZ = -tempTiltZ    'spins towards the observer     'FLAG need to verify that the sign is correct here... FLAGFLAG
            If fixtureOnPole = 2 Then tempTiltZ = tempTiltZ     'spins away from the observer
            
            'assign to arrays
            FPArrayX(fixtureCount) = tempX
            FPArrayY(fixtureCount) = tempY
            FPArrayBackwards(fixtureCount) = facesBackwards
            FPArrayTiltZ(fixtureCount) = tempTiltZ
            
            fixtureCount = fixtureCount + 1
        Next fixtureOnPole
    Next poleSide
    polePhase = polePhase + 1
Loop

'Write the outputs to the array to pass back to the calling routine.
fixturePositions(0) = FPArrayX
fixturePositions(1) = FPArrayY
fixturePositions(2) = FPArrayBackwards
fixturePositions(3) = tiltOnX   'all the same value
fixturePositions(4) = tiltOnY   'all the same value
fixturePositions(5) = FPArrayTiltZ 'varies whether it is fixture 1 or fixture 2 on the pole


'FLAG this format has been changed to dim the arrays to base 1, instead of 0. Check for compatibility down the line
End Sub


