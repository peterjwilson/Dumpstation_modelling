Option Explicit

Public Const isHPCT = True

Public gi_hoplet_len As Integer
Public gd_pkt_kg As Double

Public HopLetEndHP(1 To 2, 1 To 4) As Integer
Public doorColors(1 To 4) As Long

Public gi_hoplet_count(1 To 2, 1 To 4) As Integer

Public gd_max_rates(1 To 2, 1 To 4) As Double       'kg/s (rate achievable accounting for possible upstream hopperlets that are empty)
Public gd_base_rates(1 To 2, 1 To 4) As Double      'kg/s (rate achievable when all hopperlets are outfeeding)
Public gd_outfeed_rates(1 To 2, 1 To 4) As Double   'tonnes/hr (current observed actual outfeed rate)
Public gd_feed_count(1 To 2, 1 To 4) As Double      'keeps a tally of the number of packets waiting to be fed out of the hoperlet

'stats collection
'
Public gd_tot_infeed_rate(2) As Double
Public gd_tot_outfeed_rate(2) As Double
Public gd_infeed_tally(2) As Double
Public gd_outfeed_tally(2) As Double
Public gd_hopper_tally(2) As Double
Public gd_infeed_start As Double
Public gb_topout(1 To 2, 1 To 4) As Boolean
Public gd_depth(1 To 2, 1 To 4) As Double
Public ga_data_cols As Collection

Public Sub FixScreenUpdating()
    Application.ScreenUpdating = True
End Sub

Public Sub CycleAllHPCT()
    Dim i As Integer
    
    For i = 1 To Sheet1.Range("H30").Value
        CycleOneWagonHPCT
    Next i
End Sub

Public Sub CycleOneWagonHPCT()
    Dim i As Integer
    
    If Range("ShowAnim") = "N" Then Application.ScreenUpdating = False
    
    For i = 1 To Round(Sheet1.Range("H29").Value, 0)
        RunOneSecondHPCT
    Next i
    
    If Range("ShowAnim") = "N" Then
        Application.ScreenUpdating = True
        'force screen to redraw
        ActiveSheet.Calculate
        DoEvents
    End If
End Sub

Public Sub PrepareHPCT()
    Dim i As Integer, j As Integer
    
    If Range("ShowAnim") = "N" Then Application.ScreenUpdating = False
    
    'initalises display

    ClearDisplay
    InitiateArrayHPCT
    timeInc = 1
    
    Hopletend4 = HopLetEndHP(2, 4)      'this is just to trick dtgraph to account for 2 hoppers
    dtgraph
    ElementCountHPCT
     
    For i = 1 To 2
        For j = 1 To 4
            gd_feed_count(i, j) = 0
            gb_topout(i, j) = False
        Next j
    Next i
    doorColors(1) = RGB(255, 0, 0)
    doorColors(2) = RGB(0, 255, 0)
    doorColors(3) = RGB(0, 0, 255)
    doorColors(4) = RGB(255, 140, 0)
    
    DrawHopperHPCT
    PrepareDashboard
        
    onewagononly = 1
    
    If Range("ShowAnim") = "N" Then Application.ScreenUpdating = True
End Sub

Public Sub ElementCountHPCT()              'counts and displays number of elements in each hopper - fix this!!
    Dim i As Integer, j As Integer, a As Integer, b As Integer, i_bstart As Integer
    Dim d_depth As Double, d_pkt_ht As Double
    
    d_pkt_ht = Sheets("Input page").Range("G93")
    For i = 1 To 2
        gd_hopper_tally(i) = 0
        For j = 1 To 4
            gi_hoplet_count(i, j) = 0
            gd_depth(i, j) = 0
        Next j
    Next i
    
    i_bstart = 1
    For i = 1 To 2                                      'dumpstation
        For j = 1 To 4                                  'hoperlet
            For a = 1 To jMax                           'depth
                d_depth = (jMax + 1 - a) * d_pkt_ht
                For b = i_bstart To HopLetEndHP(i, j)   'longitudinal position (within bounds of dumpstation-hoperlet)
                    If HopperArray(b, a) Then
                        gi_hoplet_count(i, j) = gi_hoplet_count(i, j) + 1
                        gd_hopper_tally(i) = gd_hopper_tally(i) + 0.001 * gd_pkt_kg
                        gd_depth(i, j) = WorksheetFunction.Max(gd_depth(i, j), d_depth)
                    End If
                Next b
            Next a
        i_bstart = HopLetEndHP(i, j) + 1
        Next j
    Next i
End Sub

Public Sub InitiateArrayHPCT()
    Dim i As Integer, j As Integer
    
    'declare height and length of hopper
    
    iMax = 2 * Sheet1.Range("H136").Value       'double for second dumpstation for HPCT
    jMax = Sheet1.Range("H137").Value
    
    'declare location and range of hopplets / feeders and hopper wall height
    gi_hoplet_len = Sheet1.Range("H138").Value
    For i = 1 To 2
        For j = 1 To 4
            HopLetEndHP(i, j) = (4 * (i - 1) + j) * gi_hoplet_len
        Next j
    Next i
    HopperWallHeight = Sheet1.Range("H139").Value
    
    'declare array of True/False values
    ReDim HopperArray(1 To iMax, 1 To jMax) As Boolean
        
    For i = 1 To iMax
        For j = 1 To jMax
            HopperArray(i, j) = False
        Next j
    Next i
      
End Sub


Public Sub DrawHopperHPCT()
    
    'draws the hopper based on information on input sheet
    
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    
    Dim xOffset As Integer, i As Integer
    Dim s As Worksheet
    
    Set s = Sheet2
    s.Range(s.Cells(yDisplay + jMax + 1, 1), s.Cells(500, 1)).Rows.RowHeight = 18
    s.Range(s.Cells(yDisplay + 1, 1), s.Cells(yDisplay + jMax, 1)).Rows.RowHeight = 6
    
    xOffset = xDisplay
    For i = 1 To 2          'updated for HPCT to show two separate hoppers in sequence
        s.Range(s.Cells(yDisplay + 1, xOffset), s.Cells(yDisplay + jMax, xOffset)).Borders(xlEdgeRight).Weight = xlThick
        s.Range(s.Cells(yDisplay + 1, xOffset), s.Cells(yDisplay + jMax, xOffset)).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
        s.Range(s.Cells(yDisplay + 1, xOffset), s.Cells(yDisplay + jMax, xOffset + HopLetEndHP(1, 4))).Borders(xlEdgeRight).Weight = xlThick
        s.Range(s.Cells(yDisplay + 1, xOffset), s.Cells(yDisplay + jMax, xOffset + HopLetEndHP(1, 4))).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
        
        s.Range(s.Cells(yDisplay, xOffset + 1), s.Cells(yDisplay, xOffset + HopLetEndHP(1, 4))).Borders(xlEdgeBottom).Weight = xlThick
        s.Range(s.Cells(yDisplay, xOffset + 1), s.Cells(yDisplay, xOffset + HopLetEndHP(1, 4))).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        s.Range(s.Cells(yDisplay + jMax, xOffset + 1), s.Cells(yDisplay, xOffset + HopLetEndHP(1, 4))).Borders(xlEdgeBottom).Weight = xlThick
        s.Range(s.Cells(yDisplay + jMax, xOffset + 1), s.Cells(yDisplay, xOffset + HopLetEndHP(1, 4))).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        
        For c = 1 To 3
            s.Range(s.Cells(yDisplay + jMax - HopperWallHeight + 1, xOffset + c * HopLetEndHP(1, 1)), Cells(yDisplay + jMax, xOffset + c * HopLetEndHP(1, 1))).Borders(xlEdgeRight).Weight = xlThick
            s.Range(s.Cells(yDisplay + jMax - HopperWallHeight + 1, xOffset + c * HopLetEndHP(1, 1)), Cells(yDisplay + jMax, xOffset + c * HopLetEndHP(1, 1))).Borders(xlEdgeRight).Color = RGB(0, 0, 0)
        Next c
        xOffset = xOffset + HopLetEndHP(1, 4)
    Next i
    
End Sub

Sub RunOneSecondHPCT()

Dim a As Integer
Dim b As Integer
Dim i As Integer, i_train_pos As Integer, i_door As Integer, h As Integer
           
    gd_pkt_kg = Sheet1.Range("G96")     'mass of one material block
    Hopletend4 = HopLetEndHP(2, 4)      'this is just to trick dtgraph to account for 2 hoppers
    dtgraph
    FeederHPCT
    
    For h = 1 To 2
        gd_tot_infeed_rate(h) = 0
    Next h
    
    'scans grid above hopper and calls on particlefall
    For a = 1 To 8 * gi_hoplet_len
        i_train_pos = a - Offset      'NOTE: Offset pre-calculated in dtgraph subroutine

        If Sheet4.Cells(timeInc + 3, a + 2).Value > 0 Then
            For i = 1 To 4
                If Sheet4.Cells(timeInc + 3, a + 2).Interior.Color = doorColors(i) Then
                    i_door = i
                    Exit For
                End If
            Next i
        
            For b = 1 To Sheet4.Cells(timeInc + 3, a + 2).Value
                ParticleFallHPCT i_train_pos, i_door
            Next b
        End If
    Next a
    
    timeInc = timeInc + 1
    ElementCountHPCT
    UpdateDashboard
End Sub

Function gi_get_h(i As Integer) As Integer
    If i <= 4 * gi_hoplet_len Then gi_get_h = 1 Else gi_get_h = 2
End Function

Function gi_get_k(i As Integer) As Integer
        gi_get_k = WorksheetFunction.RoundDown((i - 1) / gi_hoplet_len, 0) + 1
        If gi_get_k > 4 Then gi_get_k = gi_get_k - 4
End Function


Public Sub PrepareDashboard()
    Dim s As Worksheet, s_d As Worksheet
    Dim h As Integer, k As Integer, i_rw As Integer, i_lbl_cl As Integer, i_cl As Integer
    
    Set s = Sheets("Output")
    Set s_d = Sheets("DATA")
    s.Rows("90:500").ClearContents
    i_rw = yDisplay + jMax + 2
    For i_lbl_cl = xDisplay To xDisplay + 8 * gi_hoplet_len + 1 Step 8 * gi_hoplet_len + 1
        s.Cells(i_rw + 1, i_lbl_cl).Value = "Elements"
        s.Cells(i_rw + 2, i_lbl_cl).Value = "Mass-t"
        s.Cells(i_rw + 3, i_lbl_cl).Value = "Fill-%"
        s.Cells(i_rw + 4, i_lbl_cl).Value = "FeedOut-tph"
        s.Cells(i_rw + 5, i_lbl_cl).Value = "Depth-m"
        s.Cells(i_rw + 6, i_lbl_cl).Value = "MaxDepth-m"
    
        s.Cells(i_rw + 8, i_lbl_cl) = "InstFeedIn-tph"
        s.Cells(i_rw + 9, i_lbl_cl) = "InstFeedOut-tph"
        s.Cells(i_rw + 10, i_lbl_cl) = "AvgFeedIn-tph"
        s.Cells(i_rw + 11, i_lbl_cl) = "AvgFeedOut-tph"
        s.Cells(i_rw + 12, i_lbl_cl) = "InfeedTally-t"
        s.Cells(i_rw + 13, i_lbl_cl) = "OutfeedTally-t"
        s.Cells(i_rw + 14, i_lbl_cl) = "HopperMass-t"
    Next i_lbl_cl
    s.Range(s.Cells(i_rw, xDisplay + 8 * gi_hoplet_len + 1), s.Cells(i_rw + 14, xDisplay + 8 * gi_hoplet_len + 1)).HorizontalAlignment = xlLeft
    s.Range(s.Cells(i_rw, xDisplay), s.Cells(i_rw + 14, xDisplay)).HorizontalAlignment = xlRight
    For h = 1 To 2
        For k = 1 To 4
            i_lbl_cl = xDisplay + (4 * (h - 1) + k) * gi_hoplet_len
            s.Cells(i_rw, i_lbl_cl).Value = "Hopperlet" & h & "-" & k
            s.Range(s.Cells(i_rw, i_lbl_cl), s.Cells(i_rw + 14, i_lbl_cl)).HorizontalAlignment = xlRight
        Next k
        gd_outfeed_tally(h) = 0
        gd_infeed_tally(h) = 0
    Next h
    gd_infeed_start = 0
    
    'prepare the data sheet
    '
    Set ga_data_cols = New Collection
    s_d.Cells.ClearContents
    s_d.Range("A1") = "time"
    i_cl = 2
    For h = 1 To 2
         s_d.Cells(1, i_cl) = "tonsH" & h
         ga_data_cols.Add Item:=i_cl, Key:=s_d.Cells(1, i_cl).Value
         i_cl = i_cl + 1
    Next h
    For h = 1 To 2
         s_d.Cells(1, i_cl) = "rateH" & h
         ga_data_cols.Add Item:=i_cl, Key:=s_d.Cells(1, i_cl).Value
         i_cl = i_cl + 1
    Next h
    For h = 1 To 2
        For k = 1 To 4
            s_d.Cells(1, i_cl) = "depthH" & h & "F" & k
            ga_data_cols.Add Item:=i_cl, Key:=s_d.Cells(1, i_cl).Value
            i_cl = i_cl + 1
        Next k
    Next h
    For h = 1 To 2
        For k = 1 To 4
            s_d.Cells(1, i_cl) = "tonsH" & h & "F" & k
            ga_data_cols.Add Item:=i_cl, Key:=s_d.Cells(1, i_cl).Value
            i_cl = i_cl + 1
        Next k
    Next h
    For h = 1 To 2
        For k = 1 To 4
            s_d.Cells(1, i_cl) = "rateH" & h & "F" & k
            ga_data_cols.Add Item:=i_cl, Key:=s_d.Cells(1, i_cl).Value
            i_cl = i_cl + 1
        Next k
    Next h
End Sub

Function gi_data_cl(ac_stat As String, ai_hopper As Integer, ai_feeder As Integer) As Integer
    Dim c_key As String
    
    c_key = ac_stat & "H" & ai_hopper
    If ai_feeder > 0 Then c_key = c_key & "F" & ai_feeder
    gi_data_cl = ga_data_cols(c_key)
End Function

Public Sub UpdateDashboard()
    Dim s As Worksheet, s_d As Worksheet
    Dim h As Integer, k As Integer, i_lbl_cl As Integer, i_rw As Integer
    Dim d_pkt_kg As Double, t As Double, d_in_rate(1 To 2), d_out_rate(1 To 2)
    
    Set s = Sheets("Output")
    Set s_d = Sheets("DATA")
    d_pkt_kg = 0.001 * Sheet1.Range("G96")
    
    t = s.Range("Y74")
    s_d.Cells(t + 1, 1) = t
    i_rw = yDisplay + jMax + 2
    For h = 1 To 2
        If gd_infeed_start = 0 And gd_infeed_tally(h) > 0 Then gd_infeed_start = t
         For k = 1 To 4
            i_lbl_cl = xDisplay + (4 * (h - 1) + k) * gi_hoplet_len
            s.Cells(i_rw + 1, i_lbl_cl).Value = gi_hoplet_count(h, k) & " pkt"
            s.Cells(i_rw + 2, i_lbl_cl).Value = Round(gi_hoplet_count(h, k) * d_pkt_kg, 2) & " t"
            's.Cells(i_rw + 3, i_lbl_cl).Value = "TODO" & " %"
            s.Cells(i_rw + 4, i_lbl_cl).Value = Round(gd_outfeed_rates(h, k), 0) & " tph"
            s.Cells(i_rw + 5, i_lbl_cl).Value = gd_depth(h, k) & " m"
            's.Cells(i_rw + 6, i_lbl_cl).Value = "TODO" & " m"
            
            If gb_topout(h, k) Then s.Range(s.Cells(i_rw, i_lbl_cl + 1 - gi_hoplet_len), s.Cells(i_rw, i_lbl_cl)).Interior.Color = RGB(255, 0, 0)
            
            s_d.Cells(t + 1, gi_data_cl("tons", h, k)) = gi_hoplet_count(h, k) * d_pkt_kg
            s_d.Cells(t + 1, gi_data_cl("rate", h, k)) = gd_outfeed_rates(h, k)
            s_d.Cells(t + 1, gi_data_cl("depth", h, k)) = gd_depth(h, k)
        Next k
        
        i_lbl_cl = xDisplay + (4 * (h - 1) + 1) * gi_hoplet_len
        s.Cells(i_rw + 8, i_lbl_cl) = Round(gd_tot_infeed_rate(h), 0) & " tph"
        s.Cells(i_rw + 9, i_lbl_cl) = Round(gd_tot_outfeed_rate(h), 0) & " tph"
        
        If gd_infeed_start > 0 And t > gd_infeed_start Then
            d_in_rate(h) = 3600 * gd_infeed_tally(h) / (t - gd_infeed_start)
            s.Cells(i_rw + 10, i_lbl_cl) = Round(d_in_rate(h), 0) & " tph"
        End If
        If gd_infeed_start > 0 And t > gd_infeed_start Then
            d_out_rate(h) = 3600 * gd_outfeed_tally(h) / (t - gd_infeed_start)
            s.Cells(i_rw + 11, i_lbl_cl) = Round(d_out_rate(h), 0) & " tph"
        End If
        
        s.Cells(i_rw + 12, i_lbl_cl) = Round(gd_infeed_tally(h), 1) & " t"
        s.Cells(i_rw + 13, i_lbl_cl) = Round(gd_outfeed_tally(h), 1) & " t"
        s.Cells(i_rw + 14, i_lbl_cl) = Round(gd_hopper_tally(h), 1) & " t"
        
        s_d.Cells(t + 1, gi_data_cl("tons", h, 0)) = gd_hopper_tally(h)
        s_d.Cells(t + 1, gi_data_cl("rate", h, 0)) = gd_tot_outfeed_rate(h)
     Next h
    
    i_lbl_cl = xDisplay + 8 * gi_hoplet_len
    s.Cells(i_rw + 8, i_lbl_cl) = Round(gd_tot_infeed_rate(1) + gd_tot_infeed_rate(2), 0) & " tph"
    s.Cells(i_rw + 9, i_lbl_cl) = Round(gd_tot_outfeed_rate(1) + gd_tot_outfeed_rate(2), 0) & " tph"
    
    s.Cells(i_rw + 10, i_lbl_cl) = Round(d_in_rate(1) + d_in_rate(2), 0) & " tph"
    s.Cells(i_rw + 11, i_lbl_cl) = Round(d_out_rate(1) + d_out_rate(2), 0) & " tph"
    
    s.Cells(i_rw + 12, i_lbl_cl) = Round(gd_infeed_tally(1) + gd_infeed_tally(2), 1) & " t"
    s.Cells(i_rw + 13, i_lbl_cl) = Round(gd_outfeed_tally(1) + gd_outfeed_tally(2), 1) & " t"
    s.Cells(i_rw + 14, i_lbl_cl) = Round(gd_hopper_tally(1) + gd_hopper_tally(2), 1) & " t"
End Sub

Public Sub ParticleFallHPCT(ai_train_pos As Integer, ai_door As Integer)
    Dim i As Integer, j As Integer, h As Integer, k As Integer
    Dim b_falling As Boolean        'flag to indicate if particle is falling (False indicates the particle is rolling down the slope)
    Dim b_left_wall As Boolean, b_right_wall As Boolean
    Dim s As Worksheet
    
    Set s = Sheet2
    j = 1
    i = ai_train_pos
    
    'check if hopper has topped out
    '
    While HopperArray(i, 1)   'move to right into we find some room (this simulates plowing)
        h = gi_get_h(i)
        k = gi_get_k(i)
        gb_topout(h, k) = True
        i = i + 1
        If i > iMax Then End  'there is nowhere else to fit anymore product so it's all gone to shit
    Wend
    '
    'stats collection
    '
    h = gi_get_h(i)
    gd_infeed_tally(h) = gd_infeed_tally(h) + gd_pkt_kg * 0.001
    gd_tot_infeed_rate(h) = gd_tot_infeed_rate(h) + gd_pkt_kg * 3.6
    
    Do While i >= 1 And i <= iMax And j < jMax
        If HopperArray(i, j + 1) Then
            'the particle has landed on product, need to determine if it sticks or falls down angle of repose
            '
            If b_falling Then
                RenderPacket i, j, RGB(255, 255, 255)       'set last known position back to white
                b_falling = False
            End If
            
            'check if there is are walls either side of this position
            '
            If i Mod gi_hoplet_len = 1 Then         'wall to left
                b_left_wall = i Mod (4 * gi_hoplet_len) = 1 Or j >= jMax - HopperWallHeight
            ElseIf i Mod gi_hoplet_len = 0 Then     'wall to right
                b_right_wall = i Mod (4 * gi_hoplet_len) = 0 Or j >= jMax - HopperWallHeight
            End If
            
            'check if the block should fall to one side
            '
            If Not b_right_wall And HopperArray(i, j + 1) And Not HopperArray(WorksheetFunction.Min(iMax, i + 1), j + 1) Then        'check whether to fall right
                i = i + 1
                j = j + 1
            ElseIf Not b_left_wall And HopperArray(i, j + 1) And Not HopperArray(WorksheetFunction.Max(1, i - 1), j + 1) Then     'check whether to fall left
                i = i - 1
                j = j + 1
            Else
                Exit Do     'the particle has landed somewhere it can stick
            End If
        Else 'HopperArray(i,j) is False
            b_falling = True
            j = j + 1
            RenderPacket i, j, doorColors(ai_door)
            RenderPacket i, j - 1, RGB(255, 255, 255)
        End If
    Loop
    HopperArray(i, j) = True
    RenderPacket i, j, doorColors(ai_door)
End Sub

Sub RenderPacket(i As Integer, j As Integer, ai_color As Long)
    Sheet2.Cells(j + yDisplay, i + xDisplay).Interior.Color = ai_color
End Sub

Function gi_query_color(i As Integer, j As Integer) As Long
    gi_query_color = Sheet2.Cells(j + yDisplay, i + xDisplay).Interior.Color
End Function


'-------------------------------------------------------------------------------
'Function FeederHPCT
'-------------------------------------------------------------------------------
'DESCRIPTION:
'	Outloading function drawing material from the hopper.
'
'ARGUMENTS:
'	NONE
'
'NOTE: 
'	this draws down the entire hopper uniformally, ie, not hoperlets 
'	individually, this is to prevent angle of repose violations
'
Public Sub FeederHPCT()
    Dim n As Integer, i_hopper_number As Integer, i_feeding_direction As Integer, i_hopperlet_number As Integer, i_hopper_bot_layer_mat_count As Integer, i_hopperlet_bot_layer_mat_count As Integer
    Dim i As Integer, j As Integer
    Dim d_hopper_outfeed As Double, d_hopper_outfeed_pkt As Double
    
    For i_hopper_number = 1 To 2          'for each of the 2 hoppers
        gd_tot_outfeed_rate(i_hopper_number) = 0
        'start by initialising data for this hopper
        '
        i_feeding_direction = 3 - 2 * i_hopper_number   'set i_feeding_direction to +1 for the first hopper and -1 for the second
        For i_hopperlet_number = 1 To 4
            'hopperlet rates are assumed symettric from the point between the 2 dump stations
            '
            gd_base_rates(i_hopper_number, i_hopperlet_number) = Sheet1.Range("G51").Offset(i_hopper_number, i_hopperlet_number)
        Next i_hopperlet_number
        
		
        'determine the outfeed rates for each hoperlet (r'[n] = {r[n] if s[n+1]>0; r'[n+1]+r[n] otherwise; where i+1 hoperlet is upstream)
        '
        d_hopper_outfeed_pkt = 0
		
		'SET FEEDING DIRECTION (1ST HOPPER = FORWARD, 2ND HOPPER = BACKWARD)
        If i_feeding_direction > 0 Then i_hopperlet_number = 1 Else i_hopperlet_number = 4      '2nd dumpstation has hopperlet 1 upstream, 2nd dumpstation has hopperlet 4 upstream
        n = 4
		
        While 1 <= i_hopperlet_number And i_hopperlet_number <= 4
            If n = 4 Then
                gd_max_rates(i_hopper_number, i_hopperlet_number) = gd_base_rates(i_hopper_number, i_hopperlet_number)
            Else
                gd_max_rates(i_hopper_number, i_hopperlet_number) = gd_base_rates(i_hopper_number, i_hopperlet_number) + gd_max_rates(i_hopper_number, i_hopperlet_number - i_feeding_direction) - gd_outfeed_rates(i_hopper_number, i_hopperlet_number - i_feeding_direction) / 3.6 'account for upstream hopperlet operating at less than full rate, so it's base rate can transfer to this hoperlet
            End If
                        
            'determine how many packets are to be fed
            '
            d_hopper_outfeed = WorksheetFunction.Min(gd_max_rates(i_hopper_number, i_hopperlet_number), gi_hoplet_count(i_hopper_number, i_hopperlet_number) * gd_pkt_kg)
            gd_tot_outfeed_rate(i_hopper_number) = gd_tot_outfeed_rate(i_hopper_number) + d_hopper_outfeed * 3.6
            gd_feed_count(i_hopper_number, i_hopperlet_number) = gd_feed_count(i_hopper_number, i_hopperlet_number) + d_hopper_outfeed / gd_pkt_kg
            If d_hopper_outfeed > 0 Then gd_outfeed_rates(i_hopper_number, i_hopperlet_number) = d_hopper_outfeed * 3.6 Else gd_outfeed_rates(i_hopper_number, i_hopperlet_number) = 0
            d_hopper_outfeed_pkt = d_hopper_outfeed_pkt + gd_feed_count(i_hopper_number, i_hopperlet_number)
            
            'update counters to next downstream hopperlet
            '
            i_hopperlet_number = i_hopperlet_number + i_feeding_direction
            n = n - 1
        Wend 'loop over all hopperlets
		
		
		' count how many blocks on the bottom layer of the current hopper have material
        i_hopper_bot_layer_mat_count = fnc_i_bot_layer_mat_count(4 * (i_hopper_number - 1) * gi_hoplet_len + 1, 4 * i_hopper_number * gi_hoplet_len)
		
		
		'outload if: 	- we have material at the bottom of our hopper, and
		'				- there's space in the hopper outload rate for the bottom layer of our hopper
        While i_hopper_bot_layer_mat_count > 0 And Round(d_hopper_outfeed_pkt, 0) >= i_hopper_bot_layer_mat_count
		
            'update packet counters
            gd_outfeed_tally(i_hopper_number) = gd_outfeed_tally(i_hopper_number) + i_hopper_bot_layer_mat_count * gd_pkt_kg * 0.001 	'(tonnes)
            d_hopper_outfeed_pkt = d_hopper_outfeed_pkt - i_hopper_bot_layer_mat_count
			
			'loop over hopperlets, reduce their feeding capacity because we just took their bottom layer
            For i_hopperlet_number = 1 To 4
                i_hopperlet_bot_layer_mat_count = fnc_i_bot_layer_mat_count((4 * (i_hopper_number - 1) + (i_hopperlet_number - 1)) * gi_hoplet_len + 1, (4 * (i_hopper_number - 1) + i_hopperlet_number) * gi_hoplet_len)
                gd_feed_count(i_hopper_number, i_hopperlet_number) = gd_feed_count(i_hopper_number, i_hopperlet_number) - i_hopperlet_bot_layer_mat_count
            Next i_hopperlet_number
             
            'need to remove the bottom row and index the material towards the bottom of the hopperlet
			'loop over entire hopper
            For j = jMax To 2 Step -1	'work vertically from bottom to top
                For i = 4 * (i_hopper_number - 1) * gi_hoplet_len + 1 To 4 * i_hopper_number * gi_hoplet_len	'work horizontally over the range of the hopper
                    HopperArray(i, j) = HopperArray(i, j - 1)		'shift everything vertically down one block (ie. falling material)
                    RenderPacket i, j, gi_query_color(i, j - 1)		'update colours
                Next i
            Next j
			
			'because we just moved everything down, there is now space along the entire top row of the hopper
            For i = 4 * (i_hopper_number - 1) * gi_hoplet_len + 1 To 4 * i_hopper_number * gi_hoplet_len
                HopperArray(i, 1) = False
                RenderPacket i, 1, RGB(255, 255, 255)
            Next i
            
            i_hopper_bot_layer_mat_count = fnc_i_bot_layer_mat_count(4 * (i_hopper_number - 1) * gi_hoplet_len + 1, 4 * i_hopper_number * gi_hoplet_len)
        Wend
    Next i_hopper_number
	
End Sub

'-------------------------------------------------------------------------------
'Function fnc_i_bot_layer_mat_count
'-------------------------------------------------------------------------------
'DESCRIPTION:
'	Given a horizontal range of blocks from 'ai_hopplet_block_index_start' to 
'	'ai_hopplet_block_index_end', return how many have material
'
'ARGUMENTS:
'	ai_hopplet_block_index_start = Index of hopperlet horizontal starting block 
'	(referred to hopper)
'
'	ai_hopplet_block_index_start = Index of hopperlet horizontal end block 
'	(referred to hopper)
'

Private Function fnc_i_bot_layer_mat_count(ai_hopplet_block_index_start As Integer, ai_hopplet_block_index_end As Integer) As Integer
    Dim i As Integer
    
    fnc_i_bot_layer_mat_count = 0
	
    For i = ai_hopplet_block_index_start To ai_hopplet_block_index_end
		'count the number of blocks at the bottom of the hopperlet that have material
        If HopperArray(i, jMax) Then fnc_i_bot_layer_mat_count = fnc_i_bot_layer_mat_count + 1
    Next i
	
End Function
