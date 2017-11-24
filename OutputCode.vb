Option Explicit


Sub InitiateArray()
    
    'declare height and length of hopper
    
    iMax = Sheet1.Range("H136").Value
    jMax = Sheet1.Range("H137").Value + Sheet1.Range("H139").Value
    
    'declare location and range of hopplets / feeders and hopper wall height
    Hopletend1 = Sheet1.Range("H138").Value
    Hopletend2 = Hopletend1 * 2
    Hopletend3 = Hopletend1 * 3
    Hopletend4 = Hopletend1 * 4
    HopperWallHeight = Sheet1.Range("H139").Value
    
    
    'declare array of True/False values
    ReDim HopperArray(1 To iMax, 1 To jMax) As Boolean
        
    Dim intI As Integer, intJ As Integer
    For intI = 1 To iMax
        For intJ = 1 To jMax
            HopperArray(intI, intJ) = False
        Next intJ
    Next intI
    
    
    
End Sub

Sub ParticleFall()
    
    jPos = 1
    iPos = iTrainPos

    
    Do While iPos > 1 And iPos < iMax And jPos < jMax
        
    'SPECIAL DUMPING CASE 1 - hopperlet1 wall to the right and train is to the right of the wall 1
       If iPos = Hopletend1 And iTrainPos < Hopletend1 + 1 Then
            
            '1.1 If the block below is false keep going down till true
            If HopperArray(iPos, jPos + 1) = False Then
                    jPos = jPos + 1
                    
                'colour particle according to door which it left and make it look like its falling
                If door1 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 0, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door2 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 255, 0)
                 Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door3 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 0, 255)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door4 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 140, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                End If
                Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)

            '1.2. If block below is TRUE AND the block to the lower left is FALSE - change co-ordinates of new block to the diagonal left
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = False Then
                iPos = iPos - 1
                jPos = jPos + 1
            
            '1.3 Never stack to the right unless the jpos is greater than the divider height
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = False And jPos < jMax - HopperWallHeight Then
                iPos = iPos + 1
                jPos = jPos + 1
            '1.4. If the blocks on the lower left and directly below are TRUE then exit the loop - and secure the block
            Else 'If HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = True Then
                Exit Do
        
            End If
        End If
        
    'SPECIAL DUMPING CASE 1 - hopperlet2 wall to the right and train is to the right of the wall 2
       If iPos = Hopletend2 And iTrainPos < Hopletend2 + 1 Then
            '1.1 If the block below is false keep going down till true
            If HopperArray(iPos, jPos + 1) = False Then
                    jPos = jPos + 1
                'colour particle according to door which it left and make it look like its falling
                If door1 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 0, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door2 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 255, 0)
                 Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door3 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 0, 255)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                 ElseIf door4 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 140, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                End If
                Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)

            '1.2. If block below is TRUE AND the block to the lower left is FALSE - change co-ordinates of new block to the diagonal left
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = False Then
                iPos = iPos - 1
                jPos = jPos + 1
            
            '1.3 Never stack to the right unless the jpos is greater than the divider height
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = False And jPos < jMax - HopperWallHeight Then
                iPos = iPos + 1
                jPos = jPos + 1
            '1.4. If the blocks on the lower left and directly below are TRUE then exit the loop - and secure the block
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = True Then
                Exit Do
        
            End If
        End If
        
    'SPECIAL DUMPING CASE 1 - hopperlet3 wall to the right and train is to the right of the wall 3
       If iPos = Hopletend3 And iTrainPos < Hopletend3 + 1 Then
            '1.1 If the block below is false keep going down till true
            If HopperArray(iPos, jPos + 1) = False Then
                    jPos = jPos + 1
                'colour particle according to door which it left and make it look like its falling
                If door1 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 0, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door2 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 255, 0)
                 Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door3 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 0, 255)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door4 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 140, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                End If
                Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)

            '1.2. If block below is TRUE AND the block to the lower left is FALSE - change co-ordinates of new block to the diagonal left
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = False Then
                iPos = iPos - 1
                jPos = jPos + 1
            
            '1.3 Never stack to the right unless the jpos is greater than the divider height
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = False And jPos < jMax - HopperWallHeight Then
                iPos = iPos + 1
                jPos = jPos + 1
            '1.4. If the blocks on the lower left and directly below are TRUE then exit the loop - and secure the block
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = True Then
                Exit Do
        
            End If
        End If
        
    'SPECIAL DUMPING CASE 2 - hopperlet wall to the left and train is to the left of the wall 1
        If iPos = Hopletend1 + 1 And iTrainPos > Hopletend1 Then
            '2.1 If the block below is false keep going down till true
            If HopperArray(iPos, jPos + 1) = False Then
                jPos = jPos + 1
       '         'colour particle according to door which it left and make it look like its falling
                If door1 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 0, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door2 = True Then
                   Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 255, 0)
                 Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door3 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 0, 255)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door4 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 140, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                End If
                Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)

            '1.2. If block below is TRUE AND the block to the lower right is FALSE - change co-ordinates of new block to the diagonal left
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = False Then
                iPos = iPos + 1
                jPos = jPos + 1
            
            '1.3 Never stack to the left unless the jpos is greater than the divider height
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = False And jPos < jMax - HopperWallHeight Then
                iPos = iPos - 1
                jPos = jPos + 1
            '1.4. If the blocks on the lower right and directly below are TRUE then exit the loop - and secure the block
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = True Then
                Exit Do
        
            End If
        End If
    
    'SPECIAL DUMPING CASE 2 - hopperlet wall to the left and train is to the left of the wall 2
        If iPos = Hopletend2 + 1 And iTrainPos > Hopletend2 Then
            '2.1 If the block below is false keep going down till true
            If HopperArray(iPos, jPos + 1) = False Then
                jPos = jPos + 1
       '         'colour particle according to door which it left and make it look like its falling
                If door1 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 0, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door2 = True Then
                   Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 255, 0)
                 Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door3 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 0, 255)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door4 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 140, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                    
                End If
                Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)

            '1.2. If block below is TRUE AND the block to the lower right is FALSE - change co-ordinates of new block to the diagonal left
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = False Then
                iPos = iPos + 1
                jPos = jPos + 1
            
            '1.3 Never stack to the left unless the jpos is greater than the divider height
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = False And jPos < jMax - HopperWallHeight Then
                iPos = iPos - 1
                jPos = jPos + 1
            '1.4. If the blocks on the lower right and directly below are TRUE then exit the loop - and secure the block
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = True Then
                Exit Do
        
            End If
        End If
    
    'SPECIAL DUMPING CASE 2 - hopperlet wall to the left and train is to the left of the wall 3
        If iPos = Hopletend3 + 1 And iTrainPos > Hopletend3 Then
            '2.1 If the block below is false keep going down till true
            If HopperArray(iPos, jPos + 1) = False Then
                jPos = jPos + 1
       '         'colour particle according to door which it left and make it look like its falling
                If door1 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 0, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door2 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 255, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door3 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 0, 255)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door4 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 140, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                    
                End If
                Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)

            '1.2. If block below is TRUE AND the block to the lower right is FALSE - change co-ordinates of new block to the diagonal left
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = False Then
                iPos = iPos + 1
                jPos = jPos + 1
            
            '1.3 Never stack to the left unless the jpos is greater than the divider height
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = False And jPos < jMax - HopperWallHeight Then
                iPos = iPos - 1
                jPos = jPos + 1
            '1.4. If the blocks on the lower right and directly below are TRUE then exit the loop - and secure the block
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = True Then
                Exit Do
        
            End If
        End If
    
    'GENERAL DUMPING CASE
            '1.1 If the block below is false keep going down till true
            If HopperArray(iPos, jPos + 1) = False Then
                    jPos = jPos + 1
                'colour particle according to door which it left and make it look like its falling
                If door1 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 0, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door2 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 255, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door3 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 0, 255)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                ElseIf door4 = True Then
                    Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 140, 0)
                    Sheet2.Cells(jPos - 1 + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
                    
                End If
                Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 255, 255)
            '1.2. If block below is TRUE AND the block to the lower left is FALSE - change co-ordinates of new block to the diagonal left
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = False Then
                iPos = iPos - 1
                jPos = jPos + 1
            
            '1.3. If block below is TRUE AND the block to the lower right is FALSE - change co-ords of new block to the lower right
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = False Then
                iPos = iPos + 1
                jPos = jPos + 1
            '1.4. If the blocks on the lower right and lower left and directly below are TRUE then exit the loop - and secure the block
            ElseIf HopperArray(iPos, jPos + 1) = True And HopperArray(iPos + 1, jPos + 1) = True And HopperArray(iPos - 1, jPos + 1) = True Then
                Exit Do
        
            End If
        
       
    Loop
    
    'By now, if block reaches this stage - it has been secured - colour block according to door
    HopperArray(iPos, jPos) = True
    If door1 = True Then
        Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 0, 0)
    ElseIf door2 = True Then
        Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 255, 0)
    ElseIf door3 = True Then
        Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(0, 0, 255)
    ElseIf door4 = True Then
        Sheet2.Cells(jPos + yDisplay, iPos + xDisplay).Interior.Color = RGB(255, 140, 0)
    End If
    
End Sub

Public Sub ClearDisplay()
    Dim intI As Integer, intJ As Integer
    For intI = 1 To 150
       For intJ = 1 To 55 + yDisplay                                                    'this value used to be 1 to 65
            Sheet2.Cells(intJ, intI).Interior.Color = RGB(255, 255, 255)
        Next intJ
    Next intI
    
    'reset outline of hopper
    
    Sheet2.Range("75:1000").Borders(xlDiagonalDown).LineStyle = xlNone
    Sheet2.Range("75:1000").Borders(xlDiagonalUp).LineStyle = xlNone
    Sheet2.Range("75:1000").Borders(xlEdgeLeft).LineStyle = xlNone
    Sheet2.Range("75:1000").Borders(xlEdgeTop).LineStyle = xlNone
    Sheet2.Range("75:1000").Borders(xlEdgeBottom).LineStyle = xlNone
    Sheet2.Range("75:1000").Borders(xlEdgeRight).LineStyle = xlNone
    Sheet2.Range("75:1000").Borders(xlInsideVertical).LineStyle = xlNone
    Sheet2.Range("75:1000").Borders(xlInsideHorizontal).LineStyle = xlNone
    
End Sub


Sub Prepare()

    'initalises display

    If isHPCT Then
        PrepareHPCT
        Exit Sub
    End If

    ClearDisplay
    InitiateArray
    timeInc = 1
    dtgraph
    ElementCount
    elementCount2
    Feedcount1 = 0
    Feedcount2 = 0
    Feedcount3 = 0
    Feedcount4 = 0
    door1 = False
    door2 = False
    door3 = False
    door4 = False
    DrawHopper
    Maxheight = 0
    Maxheight2 = 0
    Maxheight3 = 0
    NewWagon = 0
    f1 = 0
    f2 = 0
    f3 = 0
    f4 = 0
    Sheet2.Range("M131").Value = 0
    Sheet2.Range("M134").Value = 0
    Sheet2.Range("M137").Value = 0
    Sheet2.Range("M140").Value = 0
    Sheet2.Range("AD152").Value = 0
    Sheet2.Range("AV152").Value = 0
    Sheet2.Range("BO152").Value = 0
    Sheet2.Range("CH152").Value = 0
    Sheet2.Range("AD158").Value = 0
    Sheet2.Range("AV158").Value = 0
    Sheet2.Range("BO158").Value = 0
    Sheet2.Range("CH158").Value = 0
    roundingcounter1 = 1
    roundingcounter2 = 1
    
    
    onewagononly = 1
    'clearing graphs and data
    Worksheets("Data").Range("B2:T10000").Clear
End Sub

Sub RunOneSecond()

Dim a As Integer
Dim b As Integer
    
       
    
    dtgraph
    HeightSensor
    Feeder
    
    'scans grid above hopper and calls on particlefall
    
    For a = 1 To 102
        If Sheet4.Cells(timeInc + 3, a + 2).Value > 0 Then
            iTrainPos = a - Offset
            If Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = RGB(255, 0, 0) Then
                door1 = True
            ElseIf Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = RGB(0, 255, 0) Then
                door2 = True
            ElseIf Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = RGB(0, 0, 255) Then
                door3 = True
            ElseIf Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = RGB(255, 140, 0) Then
                door4 = True
            End If
        
            'DisplayMarker
            For b = 1 To Sheet4.Cells(timeInc + 3, a + 2).Value
                ParticleFall
            Next b
    
            door1 = False
            door2 = False
            door3 = False
            door4 = False
        End If
    Next a
    
    If Sheet2.Range("Y74").Value Mod Round(Sheet1.Range("H29").Value, 0) = 0 Then
    HeightHistory
    End If
    
    timeInc = timeInc + 1
    ElementCount
    elementCount2
    'elementCount3
    HeightMarker
    
    Sheet5.Cells(timeInc + 1, 1).Value = timeInc
    Sheet5.Cells(timeInc + 1, 2).Value = height1
    Sheet5.Cells(timeInc + 1, 3).Value = height2
    Sheet5.Cells(timeInc + 1, 4).Value = height3
    Sheet5.Cells(timeInc + 1, 5).Value = height4
    Sheet5.Cells(timeInc + 1, 6).Value = Feedrate1 * 3.6
    Sheet5.Cells(timeInc + 1, 7).Value = Feedrate2 * 3.6
    Sheet5.Cells(timeInc + 1, 8).Value = Feedrate3 * 3.6
    Sheet5.Cells(timeInc + 1, 9).Value = Feedrate4 * 3.6
    
    
End Sub

Sub ElementCount()              'counts and displays number of elements in each hopper - fix this!!


    Dim a As Integer
    Dim b As Integer
    
    h1Count = 0
    h2Count = 0
    h3Count = 0
    h4Count = 0

    For a = 1 To jMax
        For b = 1 To Hopletend1
            If HopperArray(b, a) = True Then
                h1Count = h1Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend1 + 1 To Hopletend2
            If HopperArray(b, a) = True Then
                h2Count = h2Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend2 + 1 To Hopletend3
            If HopperArray(b, a) = True Then
                h3Count = h3Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend3 + 1 To Hopletend4
            If HopperArray(b, a) = True Then
                h4Count = h4Count + 1
            End If
        Next b
    Next a
            
    Sheet2.Range("AD143").Value = h1Count
    Sheet2.Range("AV143").Value = h2Count
    Sheet2.Range("BO143").Value = h3Count
    Sheet2.Range("CH143").Value = h4Count

End Sub

Sub elementCount2()                         'counts how many blocks there are immediately above the feeders

    Dim b As Integer
    
    hb1Count = 0
    hb2Count = 0
    hb3Count = 0
    hb4Count = 0

        For b = 1 To Hopletend1
            If HopperArray(b, jMax) = True Then
                hb1Count = hb1Count + 1
            End If
        Next b
    
        For b = Hopletend1 + 1 To Hopletend2
            If HopperArray(b, jMax) = True Then
                hb2Count = hb2Count + 1
            End If
        Next b
    
        For b = Hopletend2 + 1 To Hopletend3
            If HopperArray(b, jMax) = True Then
                hb3Count = hb3Count + 1
            End If
        Next b
    
        For b = Hopletend3 + 1 To Hopletend4
            If HopperArray(b, jMax) = True Then
                hb4Count = hb4Count + 1
            End If
        Next b
            
    

End Sub

Sub elementCount3()

    Dim a As Integer
    Dim b As Integer
    
    ht1Count = 0
    ht2Count = 0
    ht3Count = 0
    ht4Count = 0

    For a = 1 To jMax
        For b = 1 To Hopletend1
            If HopperArray(b, a) = True Then
                ht1Count = ht1Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend1 + 1 To Hopletend2
            If HopperArray(b, a) = True Then
                ht2Count = ht2Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend2 + 1 To Hopletend3
            If HopperArray(b, a) = True Then
                ht3Count = ht3Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend3 + 1 To Hopletend4
            If HopperArray(b, a) = True Then
                ht4Count = ht4Count + 1
            End If
        Next b
    Next a
            
    Sheet2.Range("AD126").Value = ht1Count
    Sheet2.Range("AV126").Value = ht2Count
    Sheet2.Range("BO126").Value = ht3Count
    Sheet2.Range("CH126").Value = ht4Count

End Sub


Sub dtgraph()

'displays discharge graph on output page - current 15 sec snapshot

Dim a As Integer
Dim b As Integer


Offset = Round(Sheet1.Range("H56").Value / Sheet1.Range("G92").Value, 0)  'calculates of front section of hopper (default = 1.2m)

    For a = 1 To 15
        Sheet2.Cells(75 - a, 25).Value = Sheet4.Cells(timeInc + 2 + a, 2).Value
        For b = 1 To Hopletend4
            Sheet2.Cells(75 - a, b + 24).Interior.Color = Sheet4.Cells(timeInc + 2 + a, b + Offset).Interior.Color
        Next b
    Next a

End Sub

Public Sub Feeder()

    Dim iFeeder As Integer, jFeeder As Integer
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    
    
    MaxFeedRate1 = Sheet1.Range("H53").Value
    MaxFeedRate2 = Sheet1.Range("I53").Value
    MaxFeedRate3 = Sheet1.Range("J53").Value
    MaxFeedRate4 = Sheet1.Range("K53").Value
    
    Feedtotal = Sheet1.Range("L53").Value 'target total feed out rate
    
   'FEEDER RATE BASED ON HOPPER DEPTHS
   'assign control scheme to feeders
    If (height1 + height2 + height3 + height4) > 0 Then
        If Sheet1.Range("H54") = True Then
            Feedrate1 = MaxFeedRate1 'height1 / (height1 + height2 + height3 + height4) * Feedtotal
        Else
            Feedrate1 = height1 / (height1 + height2 + height3 + height4) * Feedtotal
        End If
        
        If Sheet1.Range("I54") = True Then
            Feedrate2 = MaxFeedRate2
        Else
            Feedrate2 = height2 / (height1 + height2 + height3 + height4) * Feedtotal
        End If
        
        If Sheet1.Range("J54") = True Then
            Feedrate3 = MaxFeedRate3
        Else
            Feedrate3 = height3 / (height1 + height2 + height3 + height4) * Feedtotal
        End If
        
        If Sheet1.Range("K54") = True Then
            Feedrate4 = MaxFeedRate4
        Else
        Feedrate4 = height4 / (height1 + height2 + height3 + height4) * Feedtotal
        
        End If
        
    End If
    
    'check feeder rates not greater than installed capacity
    
    If Feedrate1 > MaxFeedRate1 Then
        Feedrate1 = MaxFeedRate1
        Feedtotal = Feedtotal - MaxFeedRate1
        
        If (height2 + height3 + height4) > 0 Then
        Feedrate2 = height2 / (height2 + height3 + height4) * Feedtotal
        Feedrate3 = height3 / (height2 + height3 + height4) * Feedtotal
        Feedrate4 = height4 / (height2 + height3 + height4) * Feedtotal
        End If
        
        If Feedrate2 > MaxFeedRate2 Then
            Feedrate2 = MaxFeedRate2
            Feedtotal = Feedtotal - MaxFeedRate2
        
            If (height3 + height4) > 0 Then
                Feedrate3 = height3 / (height3 + height4) * Feedtotal
                Feedrate4 = height4 / (height3 + height4) * Feedtotal
            End If
        End If
    
        If Feedrate3 > MaxFeedRate3 Then
            Feedrate3 = MaxFeedRate3
            Feedtotal = Feedtotal - MaxFeedRate3
        
            If (height2 + height4) > 0 Then
                Feedrate2 = height2 / (height2 + height4) * Feedtotal
                Feedrate4 = height4 / (height2 + height4) * Feedtotal
            End If
        
        End If
    
        If Feedrate4 > MaxFeedRate4 Then
            Feedrate4 = MaxFeedRate4
            Feedtotal = Feedtotal - MaxFeedRate4
        
            If (height2 + height3) > 0 Then
                Feedrate2 = height2 / (height2 + height3) * Feedtotal
                Feedrate3 = height3 / (height2 + height3) * Feedtotal
            End If
    
        End If
    
    End If
    

    '''''''''''''''''''''''''''''''
    
    
    If Feedrate2 > MaxFeedRate2 Then
        Feedrate2 = MaxFeedRate2
        Feedtotal = Feedtotal - MaxFeedRate2
        
        If (height1 + height3 + height4) > 0 Then
        Feedrate1 = height1 / (height1 + height3 + height4) * Feedtotal
        Feedrate3 = height3 / (height1 + height3 + height4) * Feedtotal
        Feedrate4 = height4 / (height1 + height3 + height4) * Feedtotal
        End If
        
        If Feedrate3 > MaxFeedRate3 Then
            Feedrate3 = MaxFeedRate3
            Feedtotal = Feedtotal - MaxFeedRate3
        
            If (height1 + height4) > 0 Then
                Feedrate1 = height1 / (height1 + height4) * Feedtotal
                Feedrate4 = height4 / (height1 + height4) * Feedtotal
            End If
        End If
    
        If Feedrate4 > MaxFeedRate4 Then
            Feedrate4 = MaxFeedRate4
            Feedtotal = Feedtotal - MaxFeedRate4
        
            If (height1 + height3) > 0 Then
                Feedrate1 = height1 / (height1 + height3) * Feedtotal
                Feedrate3 = height3 / (height1 + height3) * Feedtotal
            End If
        
        End If
    
        If Feedrate1 > MaxFeedRate1 Then
            Feedrate1 = MaxFeedRate1
            Feedtotal = Feedtotal - MaxFeedRate1
        
            If (height3 + height4) > 0 Then
                Feedrate4 = height4 / (height4 + height3) * Feedtotal
                Feedrate3 = height3 / (height4 + height3) * Feedtotal
            End If
    
        End If
        
        
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''
    If Feedrate3 > MaxFeedRate3 Then
        Feedrate3 = MaxFeedRate3
        Feedtotal = Feedtotal - MaxFeedRate3
        
        If (height2 + height1 + height4) > 0 Then
        Feedrate2 = height2 / (height2 + height1 + height4) * Feedtotal
        Feedrate1 = height1 / (height2 + height1 + height4) * Feedtotal
        Feedrate4 = height4 / (height2 + height1 + height4) * Feedtotal
        End If
        
         If Feedrate4 > MaxFeedRate4 Then
            Feedrate4 = MaxFeedRate4
            Feedtotal = Feedtotal - MaxFeedRate4
        
            If (height1 + height2) > 0 Then
                Feedrate1 = height1 / (height1 + height2) * Feedtotal
                Feedrate2 = height2 / (height1 + height2) * Feedtotal
            End If
        End If
    
        If Feedrate1 > MaxFeedRate1 Then
            Feedrate1 = MaxFeedRate1
            Feedtotal = Feedtotal - MaxFeedRate1
        
            If (height2 + height4) > 0 Then
                Feedrate2 = height2 / (height2 + height4) * Feedtotal
                Feedrate4 = height4 / (height2 + height4) * Feedtotal
            End If
        
        End If
    
        If Feedrate2 > MaxFeedRate2 Then
            Feedrate2 = MaxFeedRate2
            Feedtotal = Feedtotal - MaxFeedRate2
        
            If (height1 + height4) > 0 Then
                Feedrate4 = height4 / (height4 + height1) * Feedtotal
                Feedrate1 = height1 / (height4 + height1) * Feedtotal
            End If
    
        End If
        
   
        
    End If
    
    
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''
    If Feedrate4 > MaxFeedRate4 Then
        Feedrate4 = MaxFeedRate4
        Feedtotal = Feedtotal - MaxFeedRate4
        
        If (height2 + height3 + height1) > 0 Then
        Feedrate2 = height2 / (height2 + height3 + height1) * Feedtotal
        Feedrate3 = height3 / (height2 + height3 + height1) * Feedtotal
        Feedrate1 = height1 / (height2 + height3 + height1) * Feedtotal
        End If
        
        If Feedrate1 > MaxFeedRate1 Then
            Feedrate1 = MaxFeedRate1
            Feedtotal = Feedtotal - MaxFeedRate1
        
            If (height2 + height3) > 0 Then
                Feedrate3 = height3 / (height3 + height2) * Feedtotal
                Feedrate2 = height2 / (height3 + height2) * Feedtotal
            End If
        End If
    
        If Feedrate2 > MaxFeedRate2 Then
            Feedrate2 = MaxFeedRate2
            Feedtotal = Feedtotal - MaxFeedRate2
        
            If (height3 + height1) > 0 Then
                Feedrate1 = height1 / (height1 + height3) * Feedtotal
                Feedrate3 = height3 / (height1 + height3) * Feedtotal
            End If
        
        End If
    
        If Feedrate3 > MaxFeedRate3 Then
            Feedrate3 = MaxFeedRate3
            Feedtotal = Feedtotal - MaxFeedRate3
        
            If (height1 + height2) > 0 Then
                Feedrate2 = height2 / (height2 + height1) * Feedtotal
                Feedrate1 = height1 / (height2 + height1) * Feedtotal
            End If
    
        End If
          
        
    End If
 
 
    'FEEDRATE CORRECTION BASED ON FEEDER PAN COVERAGE
    ' number of blocks removed per feeder = (# of blocks on bottom / feeder width)*nominal feedrate )/mass per block
    
    
    f1Count = Round(Feedrate1 / Sheet1.Range("G96").Value, 0)
    f2Count = Round(Feedrate2 / Sheet1.Range("G96").Value, 0)
    f3Count = Round(Feedrate3 / Sheet1.Range("G96").Value, 0)
    f4Count = Round(Feedrate4 / Sheet1.Range("G96").Value, 0)
    bottomrow1 = 0
    bottomrow2 = 0
    bottomrow3 = 0
    bottomrow4 = 0
   
    
    If f1Count > 0 Then
  
    Feedcount1 = Feedcount1 + f1Count
                
                                   
                    'the amount of blocks the code removes is the amount on the bottom row
                    'finding bottome row:
                    For c = 1 To Hopletend1
                        If HopperArray(c, jMax) = True Then
                        bottomrow1 = bottomrow1 + 1
                        End If
                    Next c
                   
                    
                       
                    'removing bottom row if it is less than the current running block excess:
                     If Feedcount1 >= bottomrow1 Then
                     For b = 1 To Hopletend1
                             For a = 1 To jMax
                                 Sheet2.Cells(jMax - a + 1 + yDisplay, b + xDisplay).Interior.Color = Sheet2.Cells(jMax - a + yDisplay, b + xDisplay).Interior.Color
                                 If HopperArray(b, jMax - a) = False Then
                                     HopperArray(b, jMax - a + 1) = False
                                     Exit For
                                 End If
                             Next a
                     Next b
                     blocksremoved1 = blocksremoved1 + bottomrow1
                     Feedcount1 = Feedcount1 - bottomrow1
                     
                     End If
    End If
    
    'feeder 2
    
   If f2Count > 0 Then
        Feedcount2 = Feedcount2 + f2Count
            
                    For c = 1 To Hopletend1
                        If HopperArray(c + Hopletend1, jMax) = True Then
                        bottomrow2 = bottomrow2 + 1
                        End If
                    Next c
             
                    
                 
                    If Feedcount2 >= bottomrow2 Then
                        For b = 1 To Hopletend1
                                For a = 1 To jMax
                                    Sheet2.Cells(jMax - a + 1 + yDisplay, b + Hopletend1 + xDisplay).Interior.Color = Sheet2.Cells(jMax - a + yDisplay, b + Hopletend1 + xDisplay).Interior.Color
                                    If HopperArray(b + Hopletend1, jMax - a) = False Then
                                       HopperArray(b + Hopletend1, jMax - a + 1) = False
                                        Exit For
                                    End If
                                Next a
                        Next b
                        blocksremoved2 = blocksremoved2 + bottomrow2
                        Feedcount2 = Feedcount2 - bottomrow2
                    End If
     
    End If
        
    'feeder 3
    If f3Count > 0 Then
   Feedcount3 = Feedcount3 + f3Count
  
                    For c = 1 To Hopletend1
                        If HopperArray(c + Hopletend2, jMax) = True Then
                        bottomrow3 = bottomrow3 + 1
                        End If
                    Next c
                    
                  
                
                    If Feedcount3 >= bottomrow3 Then
                    For b = 1 To Hopletend1
                           For a = 1 To jMax
                                Sheet2.Cells(jMax - a + 1 + yDisplay, b + Hopletend2 + xDisplay).Interior.Color = Sheet2.Cells(jMax - a + yDisplay, b + Hopletend2 + xDisplay).Interior.Color
                                If HopperArray(b + Hopletend2, jMax - a) = False Then
                                    HopperArray(b + Hopletend2, jMax - a + 1) = False
                                   Exit For
                                End If
                            Next a
                    Next b
                    blocksremoved3 = blocksremoved3 + bottomrow3
                    Feedcount3 = Feedcount3 - bottomrow3
                    End If
                    
    End If
        
    'feeder 4
    
    If f4Count > 0 Then
    Feedcount4 = Feedcount4 + f4Count
  
  
                    For c = 1 To Hopletend1
                        If HopperArray(c + Hopletend3, jMax) = True Then
                        bottomrow4 = bottomrow4 + 1
                        End If
                    Next c
                    
                 
                    
                
                    If Feedcount4 >= bottomrow4 Then
                    For b = 1 To Hopletend1
                           For a = 1 To jMax
                                Sheet2.Cells(jMax - a + 1 + yDisplay, b + Hopletend3 + xDisplay).Interior.Color = Sheet2.Cells(jMax - a + yDisplay, b + Hopletend3 + xDisplay).Interior.Color
                                If HopperArray(b + Hopletend3, jMax - a) = False Then
                                    HopperArray(b + Hopletend3, jMax - a + 1) = False
                                   Exit For
                                End If
                            Next a
                    
                    Next b
                    blocksremoved4 = blocksremoved4 + bottomrow4
                    Feedcount4 = Feedcount4 - bottomrow4
                    End If
                    
    End If
        
    ''''''''''''''''''''
    'DISPLAY FEEDER DATA
  
    Sheet2.Range("AD152").Value = Feedrate1 * 3.6 '1Count * Sheet1.Range("G96").Value * 3.6
    Sheet2.Range("AV152").Value = Feedrate2 * 3.6 '2Count * Sheet1.Range("G96").Value * 3.6
    Sheet2.Range("BO152").Value = Feedrate3 * 3.6 '3Count * Sheet1.Range("G96").Value * 3.6
    Sheet2.Range("CH152").Value = Feedrate4 * 3.6 '4Count * Sheet1.Range("G96").Value * 3.6
    
    
    If f1Count > f1 Then
    f1 = f1Count
    Sheet2.Range("AD158").Value = f1 * Sheet1.Range("G96").Value * 3.6
    Else
    Sheet2.Range("AD158").Value = f1 * Sheet1.Range("G96").Value * 3.6
    End If
    
     If f2Count > f2 Then
    f2 = f2Count
    Sheet2.Range("AV158").Value = f2 * Sheet1.Range("G96").Value * 3.6
    Else
    Sheet2.Range("AV158").Value = f2 * Sheet1.Range("G96").Value * 3.6
    End If
    
     If f3Count > f3 Then
    f3 = f3Count
    Sheet2.Range("BO158").Value = f3 * Sheet1.Range("G96").Value * 3.6
    Else
    Sheet2.Range("BO158").Value = f3 * Sheet1.Range("G96").Value * 3.6
    End If
    
     If f4Count > f4 Then
    f4 = f4Count
    Sheet2.Range("CH158").Value = f4 * Sheet1.Range("G96").Value * 3.6
    Else
    Sheet2.Range("CH158").Value = f4 * Sheet1.Range("G96").Value * 3.6
    End If
             
End Sub


Public Sub HeightSensor()              'counts and displays number of elements in each hopper

    Dim a As Integer
    Dim b As Integer
    
    Dim c As Integer
    Dim d As Integer
    Dim e As Integer
    Dim f As Integer
    
    
    
    

    height1 = 0
    height2 = 0
    height3 = 0
    height4 = 0



    a = Round(Hopletend1 * 0.5, 0)          'a is the lateral location of the height sensor
        For b = 1 To jMax
            If HopperArray(a, b) = True Then
                c = c + 1
            End If
        Next b
        
        If c > height1 Then
            height1 = c
        End If
        c = 0
   
    
    
   a = Round(Hopletend1 * 1.5, 0)
        For b = 1 To jMax
            If HopperArray(a, b) = True Then
                d = d + 1
            End If
        Next b
        If d > height2 Then
        height2 = d
        End If
        d = 0
    
    
a = Round(Hopletend1 * 2.5, 0)
        For b = 1 To jMax
            If HopperArray(a, b) = True Then
                e = e + 1
            End If
        Next b
        If e > height3 Then
        height3 = e
        End If
        e = 0
    
    
  a = Round(Hopletend1 * 3.5, 0)
     For b = 1 To jMax
            If HopperArray(a, b) = True Then
                f = f + 1
            End If
        Next b
        If f > height4 Then
        height4 = f
        End If
        f = 0
    
    Sheet2.Range("AD155").Value = height1 * Sheet1.Range("G93").Value
    Sheet2.Range("AV155").Value = height2 * Sheet1.Range("G93").Value
    Sheet2.Range("BO155").Value = height3 * Sheet1.Range("G93").Value
    Sheet2.Range("CH155").Value = height4 * Sheet1.Range("G93").Value
    
    
   

End Sub

Public Sub CycleOneWagon()

If onewagononly = 1 Then
    'display set of triggers:
    Sheet2.Range("M131").Value = Sheet1.Range("H102").Value
    Sheet2.Range("M134").Value = Sheet1.Range("I102").Value
    Sheet2.Range("M137").Value = Sheet1.Range("J102").Value
    Sheet2.Range("M140").Value = Sheet1.Range("K102").Value
End If
Dim x As Integer
For x = 1 To Round(Sheet1.Range("H29").Value, 0)
    RunOneSecond
    
Next x



End Sub

Public Sub DrawHopper()
    
    'draws the hopper based on information on input sheet
    
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    
    
    For a = 1 To jMax
        Sheet2.Cells(yDisplay + a, xDisplay).Borders(xlEdgeRight).Weight = xlThick
        Sheet2.Cells(yDisplay + a, xDisplay).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    
        Sheet2.Cells(yDisplay + a, xDisplay + Hopletend4).Borders(xlEdgeRight).Weight = xlThick
        Sheet2.Cells(yDisplay + a, xDisplay + Hopletend4).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    
        'For c = 1 To 15
        
        'Sheet2.Cells(yDisplay + a, xDisplay - c - 1).Borders(xlEdgeRight).LineStyle = xlDot
        'Sheet2.Cells(yDisplay + a, xDisplay - c - 1).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
        'Next c
    
    Next a
    
    For b = 1 To iMax
        Sheet2.Cells(yDisplay, xDisplay + b).Borders(xlEdgeBottom).Weight = xlThick
        Sheet2.Cells(yDisplay, xDisplay + b).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
        Sheet2.Cells(yDisplay + jMax, xDisplay + b).Borders(xlEdgeBottom).Weight = xlThick
        Sheet2.Cells(yDisplay + jMax, xDisplay + b).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        
      
    
    Next b
    
     For b = 1 To 23
        Sheet2.Cells(yDisplay, xDisplay - b).Borders(xlEdgeBottom).Weight = xlThin
        Sheet2.Cells(yDisplay, xDisplay - b).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    
     '   Sheet2.Cells(yDisplay + jMax, xDisplay - b).Borders(xlEdgeBottom).Weight = xlThin
    '    Sheet2.Cells(yDisplay + jMax, xDisplay - b).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    Next b
    
    
    For c = 1 To 3
        For d = 1 To HopperWallHeight
            Sheet2.Cells(yDisplay + jMax - d + 1, xDisplay + c * Hopletend1).Borders(xlEdgeRight).Weight = xlThick
            Sheet2.Cells(yDisplay + jMax - d + 1, xDisplay + c * Hopletend1).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
        Next d
    Next c

    
End Sub

Public Sub CycleAll()
onewagononly = 0 'this line stops the display of the triggers in the cycleonewagon sub
'cycle through all wagons
Dim x As Integer
For x = 1 To Sheet1.Range("H30").Value      'Number of wagons to go through
        Sheet2.Range("M131").Value = Sheet1.Cells(x + 101, 8).Value
        Sheet2.Range("M134").Value = Sheet1.Cells(x + 101, 9).Value
        Sheet2.Range("M137").Value = Sheet1.Cells(x + 101, 10).Value
        Sheet2.Range("M140").Value = Sheet1.Cells(x + 101, 11).Value
        CycleOneWagon
    
    
Next x


End Sub

Public Sub HeightMarker()
    
    Sheet2.Cells(yDisplay + jMax - Maxheight, xDisplay - 1).Interior.Color = RGB(255, 255, 255)
    Sheet2.Cells(yDisplay + jMax - Maxheight2, xDisplay - 3).Interior.Color = RGB(255, 255, 255)
    
    
    'Find the maximum of the 3 heights
    Maxheight = 0
    If height4 > Maxheight Then
    Maxheight = height4
    End If
    If height3 > Maxheight Then
    Maxheight = height3
    End If
    If height2 > Maxheight Then
    Maxheight = height2
    End If
    If height1 > Maxheight Then
    Maxheight = height1
    End If
    
    Sheet2.Cells(yDisplay + jMax - Maxheight, xDisplay - 1).Interior.Color = RGB(139, 0, 139)
    
   
    
    ' find max height for that simulation
    
    If Maxheight > Maxheight2 Then
    Maxheight2 = Maxheight
    Sheet2.Cells(yDisplay + jMax - Maxheight2, xDisplay - 3).Interior.Color = RGB(0, 238, 0)
    Else
    Sheet2.Cells(yDisplay + jMax - Maxheight2, xDisplay - 3).Interior.Color = RGB(0, 238, 0)
    End If
    
    

End Sub

Public Sub HeightHistory()

Dim m As Integer
Dim n As Integer

'scan max height in column before and copy into new column and save it as a history of max heights

Sheet2.Cells(yDisplay + jMax - Maxheight2, xDisplay - 3).Interior.Color = RGB(105, 105, 105)

For m = 1 To jMax
Sheet2.Cells(yDisplay + jMax - m, xDisplay - 13).Interior.Color = Sheet2.Cells(yDisplay + jMax - m, xDisplay - 11).Interior.Color
Sheet2.Cells(yDisplay + jMax - m, xDisplay - 11).Interior.Color = Sheet2.Cells(yDisplay + jMax - m, xDisplay - 9).Interior.Color
Sheet2.Cells(yDisplay + jMax - m, xDisplay - 9).Interior.Color = Sheet2.Cells(yDisplay + jMax - m, xDisplay - 7).Interior.Color
Sheet2.Cells(yDisplay + jMax - m, xDisplay - 7).Interior.Color = Sheet2.Cells(yDisplay + jMax - m, xDisplay - 5).Interior.Color
Sheet2.Cells(yDisplay + jMax - m, xDisplay - 5).Interior.Color = Sheet2.Cells(yDisplay + jMax - m, xDisplay - 3).Interior.Color

Next m

Sheet2.Cells(yDisplay + jMax - Maxheight2, xDisplay - 3).Interior.Color = RGB(225, 225, 225)
Maxheight2 = 0

End Sub


Sub TrainFlow2()
'this  sub is used when the wagons are set to auto adjusting triggers, it is just a slight modification
'of sheet4. It calculates only the current wagons mass flows

Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
a = currentwagon


'mark out trigger locations
TrigDist1 = Round(Sheet1.Range("H70").Value / Sheet1.Range("G92").Value, 0)
TrigDist2 = Round(Sheet1.Range("H71").Value / Sheet1.Range("G92").Value, 0)
TrigDist3 = Round(Sheet1.Range("H72").Value / Sheet1.Range("G92").Value, 0)
        Sheet4.Cells(3, TrigDist1 + 2).Value = "T1"
        Sheet4.Cells(3, TrigDist2 + 2).Value = "T2"
        Sheet4.Cells(3, TrigDist3 + 2).Value = "T3"

    For d = 3 To Round(Sheet1.Range("H30").Value * Sheet1.Range("H29").Value, 0) + 100
    
        Sheet4.Cells(d, TrigDist1 + 2).Interior.Color = RGB(222, 222, 222)
        Sheet4.Cells(d, TrigDist2 + 2).Interior.Color = RGB(222, 222, 222)
        Sheet4.Cells(d, TrigDist3 + 2).Interior.Color = RGB(222, 222, 222)
    Next d
    
    For e = 1 To 90
        Sheet4.Cells(2, e + 2).Value = Sheet1.Range("G92").Value * e
    Next e
        



    
        For b = 1 To Sheet1.Range("H35").Value  'Number of doors per wagon
            
            
            
            'Determine time and location at which each door opens
            'Inital time = (time to reach trigger + number of door lengths + wagon lengths travel) / train speed
            'time zero is when door 1 passes leading edge of hopper
            
            If Sheet1.Cells(a + 101, b + 7).Value = 1 Then
                initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H70").Value) / Sheet1.Range("H27").Value, 0)
                initialDistance = TrigDist1
            ElseIf Sheet1.Cells(a + 101, b + 7).Value = 2 Then
                initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H71").Value) / Sheet1.Range("H27").Value, 0)
                initialDistance = TrigDist2
            ElseIf Sheet1.Cells(a + 101, b + 7).Value = 3 Then
                initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H72").Value) / Sheet1.Range("H27").Value, 0)
                initialDistance = TrigDist3
            End If
                
                dooropentime(b) = initialTime   'this creates an array of all the door opening times
                                                'to be used in the doortiming sub
                
                
            'Determine number of blocks discharge per door based on calculated packet size
            If b = 1 Then
                Time = Sheet1.Range("H42").Value
                Blocks = Sheet1.Range("H45").Value / Sheet1.Range("G96").Value
            ElseIf b = 2 Then
                Time = Sheet1.Range("I42").Value
                Blocks = Sheet1.Range("I45").Value / Sheet1.Range("G96").Value
            ElseIf b = 3 Then
                Time = Sheet1.Range("J42").Value
                Blocks = Sheet1.Range("J45").Value / Sheet1.Range("G96").Value
            ElseIf b = 4 Then
                Time = Sheet1.Range("K42").Value
                Blocks = Sheet1.Range("K45").Value / Sheet1.Range("G96").Value
            End If

           

            'Display number of blocks to be discharge on the worksheet, colour blocks accordingly ,
            'then insert marker on timescale
            For c = 1 To Time
                Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Value = Round(Blocks, 0)
                
                If b = 1 Then
                    Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Interior.Color = RGB(255, 0, 0)
                    Sheet4.Cells(initialTime + 2 + c, 2).Interior.Color = RGB(255, 0, 0)
                ElseIf b = 2 Then
                    Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Interior.Color = RGB(0, 255, 0)
                    Sheet4.Cells(initialTime + 2 + c, 2).Interior.Color = RGB(0, 255, 0)
                ElseIf b = 3 Then
                    Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Interior.Color = RGB(0, 0, 255)
                    Sheet4.Cells(initialTime + 2 + c, 2).Interior.Color = RGB(0, 0, 255)
                ElseIf b = 4 Then
                    Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Interior.Color = RGB(255, 140, 0)
                    Sheet4.Cells(initialTime + 2 + c, 2).Interior.Color = RGB(255, 140, 0)
                End If
                
                
            Next c
        Next b
    
End Sub


Public Sub triggeradjust()
'cycles through all wagons, adjusting the triggering points after each wagon
'currently only applies to 1 type of wagon with 4 doors
    
    Sheet4.Range("C3:CO5750").ClearContents              'clear diplay
    Sheet4.Range("B3:CO5750").Interior.ColorIndex = 0    'clear display
    currentwagon = 1
    
    Dim a As Integer
    Dim b As Integer
    Dim Sorted As Boolean
    
  
    
    For a = 1 To Sheet1.Range("H30").Value      'Number of wagons to go through
        
        onewagononly = 0 'this line stops the display of the triggers in the cycleonewagon sub
        
        'display of current triggers
        Sheet2.Range("M131").Value = Sheet1.Cells(a + 101, 8).Value
        Sheet2.Range("M134").Value = Sheet1.Cells(a + 101, 9).Value
        Sheet2.Range("M137").Value = Sheet1.Cells(a + 101, 10).Value
        Sheet2.Range("M140").Value = Sheet1.Cells(a + 101, 11).Value
        
        
        
        TrainFlow2
        
        
        CycleOneWagon
        
        HeightSensor 'retrieving height information
        
        If Sheet1.Range("H35").Value = 4 Then
                        'Situation1:
                        'Sensor 2 AND 3 read much higher then one, send two big loads to trigger 1 and the rest to trigger 3
                        If (height2 * Sheet1.Range("G93").Value) > (0.65 * (Sheet1.Range("H69").Value)) And (height3 * Sheet1.Range("G93").Value) > (0.65 * (Sheet1.Range("H69").Value)) And (height1 * Sheet1.Range("G93").Value) < (0.65 * (Sheet1.Range("H69").Value)) Then
                        Sheet1.Cells(a + 102, 8).Value = 1
                        Sheet1.Cells(a + 102, 9).Value = 3
                        Sheet1.Cells(a + 102, 10).Value = 1
                        Sheet1.Cells(a + 102, 11).Value = 3
                    
                        'sensor 1 reads too high, diverts to trigger 2 and 3
                        ElseIf (height1 * Sheet1.Range("G93").Value) > (0.4 * Sheet1.Range("H69").Value) Then
                        Sheet1.Cells(a + 102, 8).Value = 2
                        Sheet1.Cells(a + 102, 9).Value = 3
                        Sheet1.Cells(a + 102, 10).Value = 2
                        Sheet1.Cells(a + 102, 11).Value = 3
                        
                        'back to original
                        ElseIf (height1 * Sheet1.Range("G93").Value) < (0.3 * Sheet1.Range("H69").Value) And (height2 * Sheet1.Range("G93").Value) < (0.5 * Sheet1.Range("H69").Value) And (height3 * Sheet1.Range("G93").Value) < (0.5 * Sheet1.Range("H69").Value) Then
                        Sheet1.Cells(a + 102, 8).Value = 1
                        Sheet1.Cells(a + 102, 9).Value = 3
                        Sheet1.Cells(a + 102, 10).Value = 1
                        Sheet1.Cells(a + 102, 11).Value = 2
            
                        Else
                        Sheet1.Cells(a + 102, 8).Value = 1
                        Sheet1.Cells(a + 102, 9).Value = 3
                        Sheet1.Cells(a + 102, 10).Value = 2
                        Sheet1.Cells(a + 102, 11).Value = 3
                        End If
                    
                    
                        
                        For b = 1 To Sheet1.Range("H35").Value  'Number of doors per wagon
                
                            If Sheet1.Cells(a + 102, b + 7).Value = 1 Then
                            initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H70").Value) / Sheet1.Range("H27").Value, 0)
                            
                            ElseIf Sheet1.Cells(a + 102, b + 7).Value = 2 Then
                            initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H71").Value) / Sheet1.Range("H27").Value, 0)
                            
                            ElseIf Sheet1.Cells(a + 102, b + 7).Value = 3 Then
                            initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H72").Value) / Sheet1.Range("H27").Value, 0)
                            End If
                            
                        
                        
                            dooropentime(b) = initialTime   'this creates an array of all the door opening times
                                                            'to be used in the doortiming sub
                        Next b
                    
                        
                        
                        'this next section of code first creates an array called sorteddooropentime, which becomes
                        'an array of door opening times from quickest to slowest. This array is then compared
                        'with the array of unsorted door open times (dooropentime) to allow insertion of correct mass
                        'discharges for each door.
                        
                        
                        
                                    Dim sorteddooropentime()
                                    Dim temp As Single
                                    Dim x As Integer
                        
                        sorteddooropentime = dooropentime
                        
                        Sorted = False
                        
                        Do While Not Sorted
                                Sorted = True
                        
                                For x = 1 To UBound(sorteddooropentime) - 1
                                    If sorteddooropentime(x) > sorteddooropentime(x + 1) Then
                                    temp = sorteddooropentime(x + 1)
                                    sorteddooropentime(x + 1) = sorteddooropentime(x)
                                    sorteddooropentime(x) = temp
                                    Sorted = False
                                    End If
                                Next x
                        Loop
                        
                        
                        
                        'This section assigns the mass discharge and times for each door
                        Dim firstdoor As Integer
                        Dim seconddoor As Integer
                        Dim thirddoor As Integer
                        Dim fourthdoor As Integer
                        
                        For firstdoor = 1 To Sheet1.Range("H35").Value
                        
                                If dooropentime(firstdoor) = sorteddooropentime(1) Then
                                Sheet1.Cells(37, firstdoor + 7).Value = 42
                                Sheet1.Cells(42, firstdoor + 7).Value = 20
                                End If
                        Next firstdoor
                        
                        For seconddoor = 1 To Sheet1.Range("H35").Value
                        
                                If dooropentime(seconddoor) = sorteddooropentime(2) Then
                                Sheet1.Cells(37, seconddoor + 7).Value = 26
                                Sheet1.Cells(42, seconddoor + 7).Value = 16
                                End If
                        Next seconddoor
                                
                        For thirddoor = 1 To Sheet1.Range("H35").Value
                        
                                If dooropentime(thirddoor) = sorteddooropentime(3) Then
                                Sheet1.Cells(37, thirddoor + 7).Value = 16
                                Sheet1.Cells(42, thirddoor + 7).Value = 10
                                End If
                        Next thirddoor
                                        
                        For fourthdoor = 1 To Sheet1.Range("H35").Value
                        
                                If dooropentime(fourthdoor) = sorteddooropentime(4) Then
                                Sheet1.Cells(37, fourthdoor + 7).Value = 16
                                Sheet1.Cells(42, fourthdoor + 7).Value = 10
                                End If
                        Next fourthdoor
                        
                            
                                        
                    
        End If
      
      
      
      
    currentwagon = currentwagon + 1
    Next a




End Sub




Sub beltfeed1()

    Sheet4.Range("C3:CO5750").ClearContents              'clear diplay
    Sheet4.Range("B3:CO5750").Interior.ColorIndex = 0    'clear display
    
    'setting trigger values to 1:
    Sheet1.Range("H102:K131").Value = 1
    'display of current triggers:
    Sheet2.Range("M131").Value = 1
    Sheet2.Range("M134").Value = 1
    Sheet2.Range("M137").Value = 1
    Sheet2.Range("M140").Value = 1
    
    
    TrainFlowBelt

    
    Dim x As Integer
    
    

    For x = 1 To Round(Sheet1.Range("H29").Value, 0) * Sheet1.Range("H30").Value
    
        
        RunOneSecondBelt
    
    
    Next x



End Sub


Sub RunOneSecondBelt()

Dim a As Integer
Dim b As Integer
    
       
    
    dtgraph
    HeightSensor
    FeederBelt
    
    'scans grid above hopper and calls on particlefall
    
    For a = 1 To 102
        If Sheet4.Cells(timeInc + 3, a + 2).Value > 0 Then
            iTrainPos = a - Offset
            If Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = RGB(255, 0, 0) Then
                door1 = True
            ElseIf Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = RGB(0, 255, 0) Then
                door2 = True
            ElseIf Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = RGB(0, 0, 255) Then
                door3 = True
            ElseIf Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = RGB(255, 140, 0) Then
                door4 = True
            End If
        
            'DisplayMarker
            For b = 1 To Sheet4.Cells(timeInc + 3, a + 2).Value
                ParticleFall
            Next b
    
            door1 = False
            door2 = False
            door3 = False
            door4 = True
        End If
    Next a
    
    If Sheet2.Range("Y74").Value Mod Round(Sheet1.Range("H29").Value, 0) = 0 Then
    HeightHistory
    End If
    
    timeInc = timeInc + 1
    
    
 
    
    
    
    HeightMarker
    
    Sheet5.Cells(timeInc + 1, 1).Value = timeInc
    Sheet5.Cells(timeInc + 1, 2).Value = height1
    Sheet5.Cells(timeInc + 1, 3).Value = height2
    Sheet5.Cells(timeInc + 1, 4).Value = height3
    Sheet5.Cells(timeInc + 1, 5).Value = height4
    Sheet5.Cells(timeInc + 1, 6).Value = massremoval(1) * 3.6
    Sheet5.Cells(timeInc + 1, 7).Value = massremoval(2) * 3.6
    Sheet5.Cells(timeInc + 1, 8).Value = massremoval(3) * 3.6
    Sheet5.Cells(timeInc + 1, 9).Value = massremoval(4) * 3.6
    Sheet5.Cells(timeInc + 1, 10).Value = Feedcount1
    Sheet5.Cells(timeInc + 1, 11).Value = Feedcount2
    Sheet5.Cells(timeInc + 1, 12).Value = Feedcount3
    Sheet5.Cells(timeInc + 1, 13).Value = Feedcount4
    
End Sub


Public Sub FeederBelt()
    
    ElementCountBelt
    
    'convert h1Count h2Count etc to arrays:
    hcount(1) = h1Count
    hcount(2) = h2Count
    hcount(3) = h3Count
    hcount(4) = h4Count
    
    
    Dim w As Integer
    Dim x As Integer
    Dim y As Integer
    Dim z As Integer
    
    'hCount is in block terms => convert to mass terms:
    For w = 1 To 4
    massinfeeder(w) = hcount(w) * Sheet1.Range("G96").Value
    Next w
    
    'calculate massremovals (in kg):
    
       
    For x = 1 To 4
    massremoval(x) = 0
    Next x
        
    For y = 1 To 4
    z = 5 - y
        If massinfeeder(z) = 0 Then
        massremoval(z) = 0
        ElseIf massinfeeder(z) > 0 Then
        massremoval(z) = (((y * 0.25) * Sheet1.Range("H16").Value / 3.6) - massremoval(2) - massremoval(3) - massremoval(4))
        End If
        If massremoval(z) > massinfeeder(z) Then
        massremoval(z) = massinfeeder(z)
        End If
        
    Next y
           
    'converting massremoval rates to blocks removed:
    f1Count = Round(massremoval(1) / Sheet1.Range("G96").Value, 0)
    f2Count = Round(massremoval(2) / Sheet1.Range("G96").Value, 0)
    f3Count = Round(massremoval(3) / Sheet1.Range("G96").Value, 0)
    f4Count = Round(massremoval(4) / Sheet1.Range("G96").Value, 0)
    
   
   
   Dim a As Integer
   Dim b As Integer
   Dim c As Integer
   Dim v As Integer
   bottomrow1 = 0
   bottomrow2 = 0
   bottomrow3 = 0
   bottomrow4 = 0





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'EXTRACTIONS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'FEEDER1
   
   
    If f1Count > 0 Then
  
    Feedcount1 = Feedcount1 + f1Count
                
                'if the current block excess is greater than the amount of blocks in the Hopperlet:
                If Feedcount1 >= h1Count Then
                      'need to clear the hopperlet
                      For w = 1 To Hopletend1
                      For v = 1 To jMax
        
                      HopperArray(w, v) = False
                      Sheet2.Cells(yDisplay + v, xDisplay + w + 1).Interior.Color = RGB(255, 255, 255)
                      Next v
                      Next w
                blocksremoved1 = blocksremoved1 + h1Count + Feedcount1
                Feedcount1 = Feedcount1 - h1Count
                
                Else
                    
                    
                    'the amount of blocks the code removes is the amount on the bottom row
                    'finding bottome row:
                    For c = 1 To Hopletend1
                        If HopperArray(c, jMax) = True Then
                        bottomrow1 = bottomrow1 + 1
                        End If
                    Next c
                   
                    
                       
                    'removing bottom row if it is less than the current running block excess:
                     If Feedcount1 >= bottomrow1 Then
                     For b = 1 To Hopletend1
                             For a = 1 To jMax
                                 Sheet2.Cells(jMax - a + 1 + yDisplay, b + xDisplay).Interior.Color = Sheet2.Cells(jMax - a + yDisplay, b + xDisplay).Interior.Color
                                 If HopperArray(b, jMax - a) = False Then
                                     HopperArray(b, jMax - a + 1) = False
                                     Exit For
                                 End If
                             Next a
                     Next b
                     blocksremoved1 = blocksremoved1 + bottomrow1
                     Feedcount1 = Feedcount1 - bottomrow1
                     
                     End If
                     
                     
                End If
                      
                   
     
     
  
    
    'FEEDER2
   
  
    If f2Count > 0 Then
    Feedcount2 = Feedcount2 + f2Count
               
                If Feedcount2 >= h2Count Then
                     'need to clear the hopperlet
                     For w = 1 To Hopletend1
                     For v = 1 To jMax
        
                         HopperArray(w + Hopletend1, v) = False
                         Sheet2.Cells(yDisplay + v, xDisplay + Hopletend1 + w).Interior.Color = RGB(255, 255, 255)
                     
                     Next v
                     Next w
                    Feedcount2 = Feedcount2 - h2Count
                    
                    blocksremoved2 = blocksremoved2 + h2Count + Feedcount2
                
                Else
                    
                    
                    For c = 1 To Hopletend1
                        If HopperArray(c + Hopletend1, jMax) = True Then
                        bottomrow2 = bottomrow2 + 1
                        End If
                    Next c
             
                    
                 
                            If Feedcount2 >= bottomrow2 Then
                                For b = 1 To Hopletend1
                                       For a = 1 To jMax
                                            Sheet2.Cells(jMax - a + 1 + yDisplay, b + Hopletend1 + xDisplay).Interior.Color = Sheet2.Cells(jMax - a + yDisplay, b + Hopletend1 + xDisplay).Interior.Color
                                            If HopperArray(b + Hopletend1, jMax - a) = False Then
                                                HopperArray(b + Hopletend1, jMax - a + 1) = False
                                               Exit For
                                            End If
                                        Next a
                                Next b
                                blocksremoved2 = blocksremoved2 + bottomrow2
                                Feedcount2 = Feedcount2 - bottomrow2
                    End If
            
                  
                End If
            
         
          
          
    End If
    
    

    'FEEDER3
   
  
    If f3Count > 0 Then
   Feedcount3 = Feedcount3 + f3Count
  
  
                If Feedcount3 >= h3Count Then
                     'need to clear the hopperlet
                     For w = 1 To Hopletend1
                     For v = 1 To jMax
        
                         HopperArray(w + Hopletend2, v) = False
                         Sheet2.Cells(yDisplay + v, xDisplay + Hopletend2 + w).Interior.Color = RGB(255, 255, 255)
                     
                     Next v
                     Next w
                        blocksremoved3 = blocksremoved3 + h3Count + Feedcount3
                        Feedcount3 = Feedcount3 - h3Count
                        
                Else
                    For c = 1 To Hopletend1
                        If HopperArray(c + Hopletend2, jMax) = True Then
                        bottomrow3 = bottomrow3 + 1
                        End If
                    Next c
                    
                  
                
                    If Feedcount3 >= bottomrow3 Then
                    For b = 1 To Hopletend1
                           For a = 1 To jMax
                                Sheet2.Cells(jMax - a + 1 + yDisplay, b + Hopletend2 + xDisplay).Interior.Color = Sheet2.Cells(jMax - a + yDisplay, b + Hopletend2 + xDisplay).Interior.Color
                                If HopperArray(b + Hopletend2, jMax - a) = False Then
                                    HopperArray(b + Hopletend2, jMax - a + 1) = False
                                   Exit For
                                End If
                            Next a
                    Next b
                    blocksremoved3 = blocksremoved3 + bottomrow3
                    Feedcount3 = Feedcount3 - bottomrow3
                    End If
                    
                    
    
    
    
    
                End If
                
                
    End If
    
    
    'FEEDER4
   
  
    If f4Count > 0 Then
    Feedcount4 = Feedcount4 + f4Count
  
  
                If Feedcount4 >= h4Count Then
                     'need to clear the hopperlet
                     For w = 1 To Hopletend1
                     For v = 1 To jMax
        
                         HopperArray(w + Hopletend3, v) = False
                         Sheet2.Cells(yDisplay + v, xDisplay + Hopletend3 + w).Interior.Color = RGB(255, 255, 255)
                     
                     Next v
                     Next w
                    blocksremoved4 = blocksremoved4 + h4Count + Feedcount4
                    Feedcount4 = Feedcount4 - h4Count
                    
                Else
                    For c = 1 To Hopletend1
                        If HopperArray(c + Hopletend3, jMax) = True Then
                        bottomrow4 = bottomrow4 + 1
                        End If
                    Next c
                    
                 
                    
                
                    If Feedcount4 >= bottomrow4 Then
                    For b = 1 To Hopletend1
                           For a = 1 To jMax
                                Sheet2.Cells(jMax - a + 1 + yDisplay, b + Hopletend3 + xDisplay).Interior.Color = Sheet2.Cells(jMax - a + yDisplay, b + Hopletend3 + xDisplay).Interior.Color
                                If HopperArray(b + Hopletend3, jMax - a) = False Then
                                    HopperArray(b + Hopletend3, jMax - a + 1) = False
                                   Exit For
                                End If
                            Next a
                    
                    Next b
                    blocksremoved4 = blocksremoved4 + bottomrow4
                    Feedcount4 = Feedcount4 - bottomrow4
                    End If
                    
                    
                    
                    
                End If
          End If
                
    End If
    ''''''''''''''''''''
    'DISPLAY FEEDER DATA
  
    Sheet2.Range("AD152").Value = massremoval(1) * 3.6
    Sheet2.Range("AV152").Value = massremoval(2) * 3.6
    Sheet2.Range("BO152").Value = massremoval(3) * 3.6
    Sheet2.Range("CH152").Value = massremoval(4) * 3.6
    Sheet5.Cells(timeInc + 1, 10).Value = blocksremoved1
    Sheet5.Cells(timeInc + 1, 11).Value = blocksremoved2
    Sheet5.Cells(timeInc + 1, 12).Value = blocksremoved3
    Sheet5.Cells(timeInc + 1, 13).Value = blocksremoved4
    Sheet5.Cells(timeInc + 1, 14).Value = f1Count
    Sheet5.Cells(timeInc + 1, 15).Value = f2Count
    Sheet5.Cells(timeInc + 1, 16).Value = f3Count
    Sheet5.Cells(timeInc + 1, 17).Value = f4Count
    Sheet5.Cells(timeInc + 1, 18).Value = Feedcount1
    Sheet5.Cells(timeInc + 1, 19).Value = Feedcount2
    Sheet5.Cells(timeInc + 1, 20).Value = Feedcount3
    Sheet5.Cells(timeInc + 1, 21).Value = Feedcount4
label1:
  
  End Sub
  




Public Sub ElementCountBelt()

'counts and displays number of elements in hopperlets.


Dim a As Integer
    Dim b As Integer
    
    h1Count = 0
    h2Count = 0
    h3Count = 0
    h4Count = 0

    For a = 1 To jMax
        For b = 1 To Hopletend1
            If HopperArray(b, a) = True Then
                h1Count = h1Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend1 + 1 To Hopletend2
            If HopperArray(b, a) = True Then
                h2Count = h2Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend2 + 1 To Hopletend3
            If HopperArray(b, a) = True Then
                h3Count = h3Count + 1
            End If
        Next b
    Next a
    
    For a = 1 To jMax
        For b = Hopletend3 + 1 To Hopletend4
            If HopperArray(b, a) = True Then
                h4Count = h4Count + 1
            End If
        Next b
    Next a
            
    Sheet2.Range("AD143").Value = h1Count
    Sheet2.Range("AV143").Value = h2Count
    Sheet2.Range("BO143").Value = h3Count
    Sheet2.Range("CH143").Value = h4Count



End Sub


Sub TrainFlowBelt()
'this sub is the same as the sub on sheet 4, but it has to be postioned here so that it can be redone since all the triggers
'were set to 1 at the beggining of sub "beltfeed1"


Dim initialTime As Integer                  'used to calculate the time when each door opens
Dim initalDistance As Integer               'used to door triggering positions and wagon location

Dim TrigDist1 As Integer                    'location of triggers
Dim TrigDist2 As Integer
Dim TrigDist3 As Integer

Dim Time As Integer
Dim Blocks As Integer

Sheet4.Range("C3:CO5750").ClearContents              'clear diplay
Sheet4.Range("B3:CO5750").Interior.ColorIndex = 0    'clear display

'mark out trigger locations
TrigDist1 = Round(Sheet1.Range("H70").Value / Sheet1.Range("G92").Value, 0)
TrigDist2 = Round(Sheet1.Range("H71").Value / Sheet1.Range("G92").Value, 0)
TrigDist3 = Round(Sheet1.Range("H72").Value / Sheet1.Range("G92").Value, 0)
        Sheet4.Cells(3, TrigDist1 + 2).Value = "T1"
        Sheet4.Cells(3, TrigDist2 + 2).Value = "T2"
        Sheet4.Cells(3, TrigDist3 + 2).Value = "T3"

   Dim d As Integer
   Dim e As Integer
   Dim a As Integer
   Dim b As Integer
   Dim c As Integer
   
    
    For d = 3 To Round(Sheet1.Range("H30").Value * Sheet1.Range("H29").Value, 0) + 100
    
        Sheet4.Cells(d, TrigDist1 + 2).Interior.Color = RGB(222, 222, 222)
        Sheet4.Cells(d, TrigDist2 + 2).Interior.Color = RGB(222, 222, 222)
        Sheet4.Cells(d, TrigDist3 + 2).Interior.Color = RGB(222, 222, 222)
    Next d
    
    For e = 1 To 90
        Sheet4.Cells(2, e + 2).Value = Sheet1.Range("G92").Value * e
    Next e
        



    For a = 1 To Sheet1.Range("H30").Value      'Number of wagons to be modelled
        For b = 1 To Sheet1.Range("H35").Value  'Number of doors per wagon
            
            
            
            'Determine time and location at which each door opens
            'Inital time = (time to reach trigger + number of door lengths + wagon lengths travel) / train speed
            'time zero is when door 1 passes leading edge of hopper
            
            If Sheet1.Cells(a + 101, b + 7).Value = 1 Then
                initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H70").Value) / Sheet1.Range("H27").Value, 0)
                initialDistance = TrigDist1
            ElseIf Sheet1.Cells(a + 101, b + 7).Value = 2 Then
                initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H71").Value) / Sheet1.Range("H27").Value, 0)
                initialDistance = TrigDist2
            ElseIf Sheet1.Cells(a + 101, b + 7).Value = 3 Then
                initialTime = Round(((a - 1) * Sheet1.Range("H21").Value + (b - 1) / 1000 * Sheet1.Range("H32").Value + Sheet1.Range("H72").Value) / Sheet1.Range("H27").Value, 0)
                initialDistance = TrigDist3
            End If
            
            'Determine number of blocks discharge per door based on calculated packet size
            If b = 1 Then
                Time = Sheet1.Range("H42").Value
                Blocks = Sheet1.Range("H45").Value / Sheet1.Range("G96").Value
            ElseIf b = 2 Then
                Time = Sheet1.Range("I42").Value
                Blocks = Sheet1.Range("I45").Value / Sheet1.Range("G96").Value
            ElseIf b = 3 Then
                Time = Sheet1.Range("J42").Value
                Blocks = Sheet1.Range("J45").Value / Sheet1.Range("G96").Value
            ElseIf b = 4 Then
                Time = Sheet1.Range("K42").Value
                Blocks = Sheet1.Range("K45").Value / Sheet1.Range("G96").Value
            End If

           

            'Display number of blocks to be discharge on the worksheet, colour blocks accordingly ,
            'then insert marker on timescale
            For c = 1 To Time
                Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Value = Round(Blocks, 0)
                
                If b = 1 Then
                    Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Interior.Color = RGB(255, 0, 0)
                    Sheet4.Cells(initialTime + 2 + c, 2).Interior.Color = RGB(255, 0, 0)
                ElseIf b = 2 Then
                    Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Interior.Color = RGB(0, 255, 0)
                    Sheet4.Cells(initialTime + 2 + c, 2).Interior.Color = RGB(0, 255, 0)
                ElseIf b = 3 Then
                    Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Interior.Color = RGB(0, 0, 255)
                    Sheet4.Cells(initialTime + 2 + c, 2).Interior.Color = RGB(0, 0, 255)
                ElseIf b = 4 Then
                    Sheet4.Cells(initialTime + 2 + c, initialDistance + 1 + c).Interior.Color = RGB(255, 140, 0)
                    Sheet4.Cells(initialTime + 2 + c, 2).Interior.Color = RGB(255, 140, 0)
                End If
                
                
            Next c
        Next b
    Next a
End Sub

Sub doordependantbelt()









End Sub

