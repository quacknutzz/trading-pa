Attribute VB_Name = "New_Logic"
Option Explicit
Sub New_logic_Daily()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim srcRow As Long, destRow As Long
    Dim start_row As Long, end_row As Long
    Dim colNum As Long ', TotalRows As Long
    Dim val, val1 As Variant
    Dim cellColor, c As Long
    Dim hasData1030, hasData1050 As Boolean
    
 
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsDest = ThisWorkbook.Sheets("New_Logic_Daily")

    'ThisWorkbook.ForceFullCalculation = True
    start_row = wsSource.Range("P1").value
    end_row = wsSource.Range("R1").value

    ' Clear old data
    With wsDest
        .Range("A3:U" & .Rows.count).ClearContents
        .Range("A3:U" & .Rows.count).Interior.ColorIndex = xlNone
        .Range("A3:U" & .Rows.count).Font.Color = 0
        .Range("A3:U" & .Rows.count).Font.Bold = False

    End With

    Dim sourceData As Variant
    sourceData = wsSource.Range("A" & start_row & ":BU" & end_row).value
    
    Dim TotalRows As Long
    TotalRows = end_row - start_row + 1
    
    ReDim color32Arr(1 To TotalRows) As String
    ReDim color36Arr(1 To TotalRows) As String
    Dim i As Long

    
    For i = 1 To TotalRows
        srcRow = start_row + i - 1
        color32Arr(i) = wsSource.cells(srcRow, 32).value
        color36Arr(i) = wsSource.cells(srcRow, 36).value
    Next i

    ReDim destData(1 To TotalRows, 1 To 21) As Variant
    ReDim col20ColorArr(1 To TotalRows) As Long
    Dim rowCount As Long
    rowCount = 0
    
    Dim overUnderCol As Long
    For i = 1 To TotalRows
        hasData1030 = True
        hasData1050 = True

        If hasData1030 Or hasData1050 Then
            rowCount = rowCount + 1
            destData(rowCount, 1) = sourceData(i, 1)
            For colNum = 18 To 19
                If hasData1030 Then
                    cellColor = LCase(color32Arr(i))
                    overUnderCol = colNum
                    hasData1030 = False
                    Select Case cellColor
                        Case "bullish"
                            destData(rowCount, overUnderCol) = "Over"
                            destData(rowCount, 2) = "BUY"
                        Case "bearish"
                            destData(rowCount, overUnderCol) = "Under"
                            destData(rowCount, 2) = "SELL"
                    End Select
                ElseIf hasData1050 Then
                    cellColor = LCase(color36Arr(i))
                    overUnderCol = colNum
                    hasData1050 = False
                    Select Case cellColor
                        Case "bullish"
                            destData(rowCount, overUnderCol) = "Over"
                            destData(rowCount, 2) = "BUY"
                        Case "bearish"
                            destData(rowCount, overUnderCol) = "Under"
                            destData(rowCount, 2) = "SELL"
                    End Select
                Else
                    Exit For
                
                End If
                
            Next colNum
            
            destData(rowCount, 3) = Round(sourceData(i, 6), 2)
            destData(rowCount, 4) = Round(sourceData(i, 60), 2)
            destData(rowCount, 6) = Round(sourceData(i, 11), 2)
            destData(rowCount, 7) = Round(sourceData(i, 12), 2)
            destData(rowCount, 8) = Round(sourceData(i, 13), 2)
            destData(rowCount, 9) = Round(sourceData(i, 14), 2)
            destData(rowCount, 10) = Round(sourceData(i, 15), 2)
            destData(rowCount, 11) = Round(sourceData(i, 16), 2)
            destData(rowCount, 12) = Round(sourceData(i, 17), 2)
            destData(rowCount, 13) = Round(sourceData(i, 21), 2)
            destData(rowCount, 14) = Round(sourceData(i, 22), 2)
            destData(rowCount, 15) = Round(sourceData(i, 23), 2)
            destData(rowCount, 16) = Round(sourceData(i, 24), 2)
            'destData(rowCount, 17) = sourceData(i, 72)
            
            'Use Supertrend from Daily in Column BU (73)
            
'            If sourceData(i, 72) > destData(rowCount, 3) Then
'                destData(rowCount, 17) = "Bearish"
'            ElseIf sourceData(i, 72) < destData(rowCount, 3) Then
'                destData(rowCount, 17) = "Bullish"
'            Else
'                destData(rowCount, 17) = "Flat"
'            End If

            If sourceData(i, 73) > destData(rowCount, 3) Then
                destData(rowCount, 17) = "Bearish"
            ElseIf sourceData(i, 73) < destData(rowCount, 3) Then
                destData(rowCount, 17) = "Bullish"
            Else
                destData(rowCount, 17) = "Flat"
            End If
            
            
            Select Case destData(rowCount, 3) - destData(rowCount, 4)
            Case Is > 0
                destData(rowCount, 21) = "Up"
            Case Is < 0
                destData(rowCount, 21) = "Down"
            Case Else
                destData(rowCount, 21) = "Flat"
            End Select
            
            
            If sourceData(i, 8) <> "" Then
                destData(rowCount, 20) = sourceData(i, 8)
                col20ColorArr(rowCount) = RGB(204, 255, 204)
            ElseIf sourceData(i, 9) <> "" Then
                destData(rowCount, 20) = sourceData(i, 9)
                col20ColorArr(rowCount) = RGB(255, 204, 204)
            End If
            
            
            
            Select Case Trim(LCase(destData(rowCount, 2)))
                Case "buy"
                    If destData(rowCount, 3) > destData(rowCount, 4) * 1.05 Then
                        destData(rowCount, 5) = "Buy on Dip"
                    Else
                        destData(rowCount, 5) = "Buy"
                    End If
                Case "sell"
                    If destData(rowCount, 3) < destData(rowCount, 4) * 0.95 Then
                        destData(rowCount, 5) = "Sell on Rise"
                    Else
                        destData(rowCount, 5) = "Sell"
                    End If
                
            End Select
        End If
    Next i
    wsDest.Activate
    If rowCount > 0 Then
        wsDest.Range("A3:U" & (2 + rowCount)).value = destData
    End If
    
    Dim r As Long
    For r = 1 To rowCount
        destRow = 2 + r
        For colNum = 18 To 19
            val = destData(r, colNum)
            If val = "Over" Then
                With wsDest.cells(destRow, colNum)
                    .Interior.Color = RGB(204, 255, 204)
                    .Font.Bold = True
                End With
            ElseIf val = "Under" Then
                With wsDest.cells(destRow, colNum)
                    .Interior.Color = RGB(255, 204, 204)
                    .Font.Bold = True
                End With
            End If
        Next colNum
        
        val = destData(r, 2)
        If val = "BUY" Then
            With wsDest.cells(destRow, 2)
                .Interior.Color = RGB(204, 255, 204)
                .Font.Bold = True
            End With
            With wsDest.cells(destRow, 1)
                .Font.Color = RGB(0, 150, 0)
                .Font.Bold = True
            End With
        ElseIf val = "SELL" Then
            With wsDest.cells(destRow, 2)
                .Interior.Color = RGB(255, 204, 204)
                .Font.Bold = True
            End With
            With wsDest.cells(destRow, 1)
                .Font.Color = RGB(180, 0, 0)
                .Font.Bold = True
            End With
        End If
        
        If col20ColorArr(r) <> 0 Then
            wsDest.cells(destRow, 20).Interior.Color = col20ColorArr(r)
        End If
        
        If wsDest.cells(destRow, 21).value = "Up" Then
            With wsDest.cells(destRow, 21)
                .Interior.Color = RGB(204, 255, 204)
                .Font.Bold = True
            End With
        ElseIf wsDest.cells(destRow, 21).value = "Down" Then
            With wsDest.cells(destRow, 21)
                .Interior.Color = RGB(255, 204, 204)
                .Font.Bold = True
            End With
        End If
            
        If wsDest.cells(destRow, 17).value = "Bullish" Then
            With wsDest.cells(destRow, 17)
                .Interior.Color = RGB(204, 255, 204)
                .Font.Bold = True
            End With
        ElseIf wsDest.cells(destRow, 17).value = "Bearish" Then
            With wsDest.cells(destRow, 17)
                .Interior.Color = RGB(255, 204, 204)
                .Font.Bold = True
            End With
        End If
        
        
        ' === NEW LOGIC: Highlight Pivot Levels ===
        
        If LCase(Trim(wsDest.cells(destRow, 2).value)) = "buy" Then
            For c = 6 To 12
                If wsDest.cells(destRow, 3).value > wsDest.cells(destRow, c).value Then
                    GoTo skipcol
                Else
                    wsDest.cells(destRow, c - 1).Interior.Color = RGB(204, 255, 204)
                    wsDest.cells(destRow, c - 1).Font.Bold = True
                    wsDest.cells(destRow, c).Interior.Color = RGB(255, 255, 160)
                    wsDest.cells(destRow, c).Font.Bold = True
                    wsDest.cells(destRow, c - 2).Interior.Color = RGB(255, 204, 204)
                    wsDest.cells(destRow, c - 2).Font.Bold = True
                    Exit For
                End If
skipcol:
            Next c
        End If
        
        If LCase(Trim(wsDest.cells(destRow, 2).value)) = "sell" Then
            For c = 12 To 6 Step -1
                If wsDest.cells(destRow, 3).value < wsDest.cells(destRow, c).value Then
                    GoTo skipCol1
                Else
                    wsDest.cells(destRow, c + 1).Interior.Color = RGB(204, 255, 204)
                    wsDest.cells(destRow, c + 1).Font.Bold = True
                    wsDest.cells(destRow, c).Interior.Color = RGB(255, 255, 160)
                    wsDest.cells(destRow, c).Font.Bold = True
                    wsDest.cells(destRow, c + 2).Interior.Color = RGB(255, 204, 204)
                    wsDest.cells(destRow, c + 2).Font.Bold = True
                    Exit For
                End If
skipCol1:
            Next c
        End If
        ' === END OF NEW LOGIC ===
         
         
         If wsDest.cells(destRow, 3).value < wsDest.cells(destRow, 4).value Then
            wsDest.cells(destRow, 3).Interior.Color = RGB(160, 0, 0)
            wsDest.cells(destRow, 4).Interior.Color = RGB(0, 160, 0)
        Else
            wsDest.cells(destRow, 4).Interior.Color = RGB(160, 0, 0)
            wsDest.cells(destRow, 3).Interior.Color = RGB(0, 160, 0)
        End If
        wsDest.cells(destRow, 3).Font.Color = RGB(255, 255, 255)
        wsDest.cells(destRow, 4).Font.Color = RGB(255, 255, 255)
        wsDest.cells(destRow, 3).Font.Bold = True
        wsDest.cells(destRow, 4).Font.Bold = True
    Next r
End Sub
Sub New_logic_Intraday()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim start_row As Long, end_row As Long
    Dim srcData As Variant
    Dim srcRows As Long, numRows As Long
    Dim destData() As Variant
    Dim destColor() As Long
    Dim color29() As String, color30() As String, color33() As String, color34() As String
    Dim r As Long, d As Long, c As Long
    Dim ColCount As Long, colNum As Long, srcCol As Long
    Dim colorVal As String
    Dim outData() As Variant
    Dim rowNum As Long
    Dim val, val1 As Variant
    
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsDest = ThisWorkbook.Sheets("New_Logic_Intraday")
    
    start_row = CLng(wsSource.Range("P1").value)
    end_row = CLng(wsSource.Range("R1").value)
    If end_row < start_row Then Exit Sub
    
    ' Load source values (columns 1..72)
    'srcData = wsSource.Range(wsSource.cells(start_row, 1), wsSource.cells(end_row, 72)).value
    ' Adding new range
    srcData = wsSource.Range(wsSource.cells(start_row, 1), wsSource.cells(end_row, 80)).value
    srcRows = UBound(srcData, 1)
    
    ' Prefetch the 4 source column colors used for Over/Under decisions:
    ReDim color29(1 To srcRows), color30(1 To srcRows), color33(1 To srcRows), color34(1 To srcRows)
    
    For r = 1 To srcRows
        color29(r) = wsSource.cells(start_row + r - 1, 29).value
        color30(r) = wsSource.cells(start_row + r - 1, 30).value
        color33(r) = wsSource.cells(start_row + r - 1, 33).value
        color34(r) = wsSource.cells(start_row + r - 1, 34).value
    Next r
    ' Prepare destination in-memory structures:
    
    ReDim destData(1 To srcRows, 1 To 26)        ' columns A..W
    ReDim destColor(1 To srcRows, 1 To 26)       ' 0 = none, 1 = green (good), 2 = red (bad)
    
    d = 0
    For r = 1 To srcRows
'        If Is_Index_Ticker = True Then
'            GoTo nxtrow
'        End If
        Dim hasData1030 As Boolean, hasData1050 As Boolean
        hasData1030 = (Len(CStr(srcData(r, 29))) > 0 Or Len(CStr(srcData(r, 30))) > 0)
        hasData1050 = (Len(CStr(srcData(r, 33))) > 0 Or Len(CStr(srcData(r, 34))) > 0)
        
        If hasData1030 Or hasData1050 Then
            d = d + 1
            ' Symbol
            destData(d, 1) = srcData(r, 1)
            
            ' Over/Under detection: destination cols 18..21 (R..U)
            ColCount = 11
            For colNum = 18 To 21
                srcCol = colNum + ColCount ' yields 29,30,33,34 as in original
                If Len(CStr(srcData(r, srcCol))) > 0 Then
                    ' get prefetched color
                    Select Case srcCol
                        Case 29: colorVal = LCase(color29(r))
                        Case 30: colorVal = LCase(color30(r))
                        Case 33: colorVal = LCase(color33(r))
                        Case 34: colorVal = LCase(color34(r))
                        Case Else: colorVal = ""
                    End Select
                    
                    Select Case colorVal
                        Case "bullish"
                            destData(d, 2) = "BUY"
                            destData(d, colNum) = "Over"
                            destColor(d, colNum) = 1 ' green
                        Case "bearish"
                            destData(d, 2) = "SELL"
                            destData(d, colNum) = "Under"
                            destColor(d, colNum) = 2 ' red
                        Case Else
                            ' If unexpected color, just place the text (no special formatting)
                            destData(d, colNum) = srcData(r, srcCol)
                            destColor(d, colNum) = 0
                    End Select
                End If
                If colNum = 19 Then ColCount = ColCount + 2
            Next colNum
            
            ' Copy other value columns (matching your original mapping)
            destData(d, 3) = Round(srcData(r, 20), 2)
            destData(d, 4) = Round(srcData(r, 45), 2)
            destData(d, 6) = Round(srcData(r, 11), 2)
            destData(d, 7) = Round(srcData(r, 12), 2)
            destData(d, 8) = Round(srcData(r, 13), 2)
            destData(d, 9) = Round(srcData(r, 14), 2)
            destData(d, 10) = Round(srcData(r, 15), 2)
            destData(d, 11) = Round(srcData(r, 16), 2)
            destData(d, 12) = Round(srcData(r, 17), 2)
            destData(d, 13) = Round(srcData(r, 21), 2)
            destData(d, 14) = Round(srcData(r, 22), 2)
            destData(d, 15) = Round(srcData(r, 23), 2)
            destData(d, 16) = Round(srcData(r, 24), 2)
            destData(d, 25) = Round(srcData(r, 80), 2)
            'destData(d, 17) = srcData(r, 72)
            
            
            If srcData(r, 72) > destData(d, 3) Then
                destData(d, 17) = "Bearish"
            ElseIf srcData(r, 72) < destData(d, 3) Then
                destData(d, 17) = "Bullish"
            Else
                destData(d, 17) = "Flat"
            End If
            
            ' Columns 22 & 23 (V & W) with color flags
            If Len(CStr(srcData(r, 37))) > 0 Then
                destData(d, 22) = srcData(r, 37)
                destColor(d, 22) = 1
            ElseIf Len(CStr(srcData(r, 41))) > 0 Then
                destData(d, 22) = srcData(r, 41)
                destColor(d, 22) = 2
            End If
            
            If Len(CStr(srcData(r, 38))) > 0 Then
                destData(d, 23) = srcData(r, 38)
                destColor(d, 23) = 1
            ElseIf Len(CStr(srcData(r, 42))) > 0 Then
                ' preserved original behavior (you used 43 in your original assignment; keep it if intentional)
                destData(d, 23) = srcData(r, 43)
                destColor(d, 23) = 2
            End If
            
            
            val = "=IF(ISERROR(VLOOKUP(""" & destData(d, 1) & """,New_Logic_Daily!$A:$U, 21, FALSE)), """", VLOOKUP(""" & destData(d, 1) & """, New_Logic_Daily!$A:$U, 21, FALSE))"
            val1 = Application.Evaluate(val)
            destData(d, 24) = val1

            ' Buy/Sell label logic -> col 5
            Select Case LCase(Trim(CStr(destData(d, 2))))
                Case "buy"
                    If IsNumeric(destData(d, 3)) And IsNumeric(destData(d, 4)) Then
                        If CDbl(destData(d, 3)) > CDbl(destData(d, 4)) * 1.05 Then
                            destData(d, 5) = "Buy on Dip"
                        Else
                            destData(d, 5) = "Buy"
                        End If
                    Else
                        destData(d, 5) = "Buy"
                    End If
                Case "sell"
                    If IsNumeric(destData(d, 3)) And IsNumeric(destData(d, 4)) Then
                        If CDbl(destData(d, 3)) < CDbl(destData(d, 4)) * 0.95 Then
                            destData(d, 5) = "Sell on Rise"
                        Else
                            destData(d, 5) = "Sell"
                        End If
                    Else
                        destData(d, 5) = "Sell"
                    End If
            End Select
        End If

    Next r
    
    numRows = d
    ' Clear old data
    With wsDest
        .Range("A3:Y" & .Rows.count).ClearContents
        .Range("A3:Y" & .Rows.count).Interior.ColorIndex = xlNone
        .Range("A3:Y" & .Rows.count).Font.Color = 0
        .Range("A3:Y" & .Rows.count).Font.Bold = False
    End With
    
    If numRows = 0 Then Exit Sub
    
    ' Prepare trimmed output array (only filled rows)
    ReDim outData(1 To numRows, 1 To 26)
    For r = 1 To numRows
        For c = 1 To 26
            outData(r, c) = destData(r, c)
        Next c
    Next r
    
    ' Bulk write values
    wsDest.Range("A3").Resize(numRows, 26).value = outData
    
    ' Re-apply formatting ONLY to the specific cells that need it (per-row)
    For r = 1 To numRows
        rowNum = 3 + r - 1
        
        ' Format header / BUY-SELL cell & symbol color based on final classification in col 2
        Select Case LCase(Trim(CStr(outData(r, 2))))
            Case "buy"
                wsDest.cells(rowNum, 2).Font.Bold = True
                wsDest.cells(rowNum, 2).Interior.Color = RGB(204, 255, 204)
                wsDest.cells(rowNum, 1).Font.Color = RGB(0, 150, 0)
                wsDest.cells(rowNum, 1).Font.Bold = True
            Case "sell"
                wsDest.cells(rowNum, 2).Font.Bold = True
                wsDest.cells(rowNum, 2).Interior.Color = RGB(255, 204, 204)
                wsDest.cells(rowNum, 1).Font.Color = RGB(180, 0, 0)
                wsDest.cells(rowNum, 1).Font.Bold = True
            Case Else
                ' nothing
        End Select
        
        If wsDest.cells(rowNum, 3).value < wsDest.cells(rowNum, 4).value Then
            wsDest.cells(rowNum, 3).Interior.Color = RGB(160, 0, 0)
            wsDest.cells(rowNum, 4).Interior.Color = RGB(0, 160, 0)
        Else
            wsDest.cells(rowNum, 4).Interior.Color = RGB(160, 0, 0)
            wsDest.cells(rowNum, 3).Interior.Color = RGB(0, 160, 0)
        End If
        wsDest.cells(rowNum, 3).Font.Color = RGB(255, 255, 255)
        wsDest.cells(rowNum, 4).Font.Color = RGB(255, 255, 255)
        wsDest.cells(rowNum, 3).Font.Bold = True
        wsDest.cells(rowNum, 4).Font.Bold = True
        ' Format Over/Under cells individually (cols R..U = 18..21)
        For c = 18 To 21
            If Len(CStr(outData(r, c))) > 0 Then
                Select Case destColor(r, c)
                    Case 1
                        wsDest.cells(rowNum, c).Font.Bold = True
                        wsDest.cells(rowNum, c).Interior.Color = RGB(204, 255, 204)
                    Case 2
                        wsDest.cells(rowNum, c).Font.Bold = True
                        wsDest.cells(rowNum, c).Interior.Color = RGB(255, 204, 204)
                    Case Else
                        ' no fill
                End Select
            End If
        Next c
        
        ' Format cols 22 & 23 individually if present
        If Len(CStr(outData(r, 22))) > 0 Then
            Select Case destColor(r, 22)
                Case 1: wsDest.cells(rowNum, 22).Interior.Color = RGB(204, 255, 204)
                Case 2: wsDest.cells(rowNum, 22).Interior.Color = RGB(255, 204, 204)
            End Select
        End If
        If Len(CStr(outData(r, 23))) > 0 Then
            Select Case destColor(r, 23)
                Case 1: wsDest.cells(rowNum, 23).Interior.Color = RGB(204, 255, 204)
                Case 2: wsDest.cells(rowNum, 23).Interior.Color = RGB(255, 204, 204)
            End Select
        End If
        
        
        
        If wsDest.cells(rowNum, 17).value = "Bullish" Then
            With wsDest.cells(rowNum, 17)
                .Interior.Color = RGB(204, 255, 204)
                .Font.Bold = True
            End With
        ElseIf wsDest.cells(rowNum, 17).value = "Bearish" Then
            With wsDest.cells(rowNum, 17)
                .Interior.Color = RGB(255, 204, 204)
                .Font.Bold = True
            End With
        End If
        
        If wsDest.cells(rowNum, 24).value = "Up" Then
            With wsDest.cells(rowNum, 24)
                .Interior.Color = RGB(204, 255, 204)
                .Font.Bold = True
            End With
        ElseIf wsDest.cells(rowNum, 24).value = "Down" Then
            With wsDest.cells(rowNum, 24)
                .Interior.Color = RGB(255, 204, 204)
                .Font.Bold = True
            End With
        End If
        
        For c = 13 To 16
            Select Case wsDest.cells(rowNum, c).value
                Case Is >= 80
                    wsDest.cells(rowNum, c).Interior.Color = RGB(210, 180, 140) 'RGB(204, 255, 204)
                Case 70 To 79, 21 To 30
                    wsDest.cells(rowNum, c).Interior.Color = RGB(255, 255, 160)
                Case Is <= 20
                    wsDest.cells(rowNum, c - 2).Interior.Color = RGB(0, 200, 0)
            End Select
            
       Next c
        

        If LCase(Trim(wsDest.cells(rowNum, 2).value)) = "buy" Then
            For c = 8 To 12   ' shifted start
                If wsDest.cells(rowNum, 3).value <= wsDest.cells(rowNum, c).value Then
                    wsDest.cells(rowNum, c).Interior.Color = RGB(255, 255, 160)
                    wsDest.cells(rowNum, c).Font.Bold = True

                    wsDest.cells(rowNum, c - 1).Interior.Color = RGB(204, 255, 204)
                    wsDest.cells(rowNum, c - 1).Font.Bold = True

                    wsDest.cells(rowNum, c - 2).Interior.Color = RGB(255, 204, 204)
                    wsDest.cells(rowNum, c - 2).Font.Bold = True

                    Exit For
                End If
            Next c
        End If

        If LCase(Trim(wsDest.cells(rowNum, 2).value)) = "sell" Then
            For c = 10 To 6 Step -1   ' shifted end
                If wsDest.cells(rowNum, 3).value >= wsDest.cells(rowNum, c).value Then
                    wsDest.cells(rowNum, c).Interior.Color = RGB(255, 255, 160)
                    wsDest.cells(rowNum, c).Font.Bold = True

                    wsDest.cells(rowNum, c + 1).Interior.Color = RGB(204, 255, 204)
                    wsDest.cells(rowNum, c + 1).Font.Bold = True

                    wsDest.cells(rowNum, c + 2).Interior.Color = RGB(255, 204, 204)
                    wsDest.cells(rowNum, c + 2).Font.Bold = True

                    Exit For
                End If
            Next c
        End If
        
        ' === END OF NEW LOGIC ===
    Next r
    'Call AppendNewLogicData
End Sub
Sub New_logic_RSI()

    ' 1. Declare Variables
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim start_row As Long, end_row As Long
    Dim srcData As Variant
    Dim srcRows As Long
    Dim destData() As Variant
    Dim r As Long, d As Long ' r for source row, d for destination row
    Dim outData() As Variant
    Dim numRows As Long
    Dim valFormula As String
    Dim valResult As Variant
    Dim c As Long ' column counter for loops
    Dim colNum As Long ' for checking source columns 21, 22, 23, 24

    ' 2. Set Worksheet References
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsDest = ThisWorkbook.Sheets("New_Logic_RSI")

    ' 3. Define Data Range (Assuming P1 and R1 contain row numbers)
    ' Read start and end row from P1 and R1 of Sheet1
    On Error Resume Next ' Handle case where P1 or R1 might not be numbers
    start_row = CLng(wsSource.Range("P1").value)
    end_row = CLng(wsSource.Range("R1").value)
    On Error GoTo 0 ' Turn error handling off

    If end_row < start_row Then Exit Sub ' Exit if range is invalid or empty

    ' 4. Load Source Values (Columns 1..72)
    ' We load all necessary columns to preserve original data assignments
    srcData = wsSource.Range(wsSource.cells(start_row, 1), wsSource.cells(end_row, 72)).value
    srcRows = UBound(srcData, 1)

    ' 5. Prepare Destination Array (Max size needed, columns 1 to 24)
    ReDim destData(1 To srcRows, 1 To 25)
    d = 0 ' Initialize destination row counter

    ' 6. Loop and Apply New Filter Condition
    For r = 1 To srcRows
        ' The new condition: Check if any of the values in source columns U (21), V (22), W (23), or X (24)
        ' is >= 80 OR <= 20.

        Dim meetsCondition As Boolean
        meetsCondition = False
        
        ' Check columns 21, 22, 23, 24
        For colNum = 21 To 24
            Dim cellValue As Variant
            cellValue = srcData(r, colNum) ' Get value from source column U, V, W, or X

            If IsNumeric(cellValue) Then ' Only check numeric values
                If CDbl(cellValue) >= 80 Or CDbl(cellValue) <= 20 Then
                    meetsCondition = True
                    Exit For ' Exit the inner column loop as soon as one condition is met
                End If
            End If
        Next colNum

        ' Only process and copy the row if the condition is met
        If meetsCondition Then
            d = d + 1
            
            ' --- Copy Data (Preserving original assignments) ---
            
            ' Symbol (Col A)
            destData(d, 1) = srcData(r, 1)
            
            ' The original code populated Col B (Buy/Sell) later in the loop.
            ' Since the Over/Under logic is removed, we must ensure Col B gets a value.
            ' We'll leave it blank for now, as the subsequent logic (Col E) and formatting depends on it.
            ' The original code's logic for Over/Under setting Col 2 (B) is removed.
            ' You may need to adjust the logic for Col 2 (Buy/Sell) if it's crucial.
            ' For now, we'll keep Col B and the related logic intact but with no value assigned here.
            
            ' Copy value columns (B, C, D, F..Q) based on original mapping
            ' Col C (3) from src Col 20
            destData(d, 3) = Round(srcData(r, 20), 2)
            ' Col D (4) from src Col 60
            destData(d, 4) = Round(srcData(r, 60), 2)
            
            ' Col F (6) to Col L (12) from src Cols 11 to 17
            destData(d, 6) = Round(srcData(r, 11), 2)
            destData(d, 7) = Round(srcData(r, 12), 2)
            destData(d, 8) = Round(srcData(r, 13), 2)
            destData(d, 9) = Round(srcData(r, 14), 2)
            destData(d, 10) = Round(srcData(r, 15), 2)
            destData(d, 11) = Round(srcData(r, 16), 2)
            destData(d, 12) = Round(srcData(r, 17), 2)
            
            ' Col M (13) to Col P (16) from src Cols 21 to 24 (the checked columns)
            destData(d, 13) = Round(srcData(r, 21), 2)
            destData(d, 14) = Round(srcData(r, 22), 2)
            destData(d, 15) = Round(srcData(r, 23), 2)
            destData(d, 16) = Round(srcData(r, 24), 2)
            destData(d, 25) = srcData(r, 65)

            ' Col Q (17) Trend logic from src Col 72 vs Col 20
            If IsNumeric(srcData(r, 72)) And IsNumeric(destData(d, 3)) Then
                If CDbl(srcData(r, 72)) > CDbl(destData(d, 3)) Then
                    destData(d, 17) = "Bearish"
                ElseIf CDbl(srcData(r, 72)) < CDbl(destData(d, 3)) Then
                    destData(d, 17) = "Bullish"
                Else
                    destData(d, 17) = "Flat"
                End If
            End If
            
            ' Col R (18) to Col U (21) - Values from the now-removed color logic columns (29,30,33,34)
            ' These are now left blank as the source data logic for them is removed.
            
            ' Col V (22) and W (23) - Values from the now-removed color flag columns (37, 41, 38, 42/43)
            ' These are now left blank as the source data logic for them is removed.
            
            ' Col X (24) VLOOKUP formula evaluation
            valFormula = "=IF(ISERROR(VLOOKUP(""" & destData(d, 1) & """,New_Logic_Daily!$A:$U, 21, FALSE)), """", VLOOKUP(""" & destData(d, 1) & """, New_Logic_Daily!$A:$U, 21, FALSE))"
            valResult = Application.Evaluate(valFormula)
            destData(d, 24) = valResult
            
            ' Buy/Sell label logic (Col B (2) determined by original Over/Under) -> Col E (5)
            ' Since Col B (2) is unassigned without the original color logic,
            ' the following logic is **disabled** and Col E (5) is left blank.
            ' If you need the Buy/Sell logic for Col E, you MUST define Col B (2) first.
            ' The previous logic for Col 2 relied on srcCols 29, 30, 33, 34.
            
            ' Re-enable a simplified Col B assignment to allow Col E logic/formatting to work:
            ' Assume "Buy" if the triggering condition was <= 20 and "Sell" if >= 80 for the first
            ' triggering column. This is an assumption based on typical RSI interpretation.
            For colNum = 21 To 24
                If IsNumeric(srcData(r, colNum)) Then
                    If CDbl(srcData(r, colNum)) <= 20 Then
                        destData(d, 2) = "BUY"
                        Exit For
                    ElseIf CDbl(srcData(r, colNum)) >= 80 Then
                        destData(d, 2) = "SELL"
                        Exit For
                    End If
                End If
            Next colNum
            
            ' Re-run Col E (5) logic now that Col B (2) is populated
            Select Case LCase(Trim(CStr(destData(d, 2))))
                Case "buy"
                    If IsNumeric(destData(d, 3)) And IsNumeric(destData(d, 4)) Then
                        If CDbl(destData(d, 3)) > CDbl(destData(d, 4)) * 1.05 Then
                            destData(d, 5) = "Buy on Dip"
                        Else
                            destData(d, 5) = "Buy"
                        End If
                    Else
                        destData(d, 5) = "Buy"
                    End If
                Case "sell"
                    If IsNumeric(destData(d, 3)) And IsNumeric(destData(d, 4)) Then
                        If CDbl(destData(d, 3)) < CDbl(destData(d, 4)) * 0.95 Then
                            destData(d, 5) = "Sell on Rise"
                        Else
                            destData(d, 5) = "Sell"
                        End If
                    Else
                        destData(d, 5) = "Sell"
                    End If
            End Select
            
        End If ' End If meetsCondition
    Next r

    numRows = d ' Final count of copied rows

    ' 7. Clear Destination Sheet
    With wsDest
        .Range("A3:Y" & .Rows.count).ClearContents
        .Range("A3:Y" & .Rows.count).Interior.ColorIndex = xlNone
        .Range("A3:Y" & .Rows.count).Font.Color = 0
        .Range("A3:Y" & .Rows.count).Font.Bold = False
    End With

    If numRows = 0 Then Exit Sub

    ' 8. Prepare Trimmed Output Array
    ReDim outData(1 To numRows, 1 To 25)
    For r = 1 To numRows
        For c = 1 To 25
            outData(r, c) = destData(r, c)
        Next c
    Next r

    ' 9. Bulk Write Values
    wsDest.Range("A3").Resize(numRows, 25).value = outData

    ' 10. Re-apply Formatting (Only simplified formatting is kept)
    For r = 1 To numRows
        Dim rowNum As Long
        rowNum = 3 + r - 1

        ' Format Col B (2) and Symbol Col A (1) based on final classification in Col B
        Select Case LCase(Trim(CStr(outData(r, 2))))
            Case "buy"
                wsDest.cells(rowNum, 2).Font.Bold = True
                wsDest.cells(rowNum, 2).Interior.Color = RGB(204, 255, 204) ' Light Green
                wsDest.cells(rowNum, 1).Font.Color = RGB(0, 150, 0)
                wsDest.cells(rowNum, 1).Font.Bold = True
            Case "sell"
                wsDest.cells(rowNum, 2).Font.Bold = True
                wsDest.cells(rowNum, 2).Interior.Color = RGB(255, 204, 204) ' Light Red
                wsDest.cells(rowNum, 1).Font.Color = RGB(180, 0, 0)
                wsDest.cells(rowNum, 1).Font.Bold = True
            'Case Else: no formatting
        End Select

        ' Format Cols C (3) and D (4) - Price/Reference Price
        If wsDest.cells(rowNum, 3).value < wsDest.cells(rowNum, 4).value Then
            wsDest.cells(rowNum, 3).Interior.Color = RGB(160, 0, 0) ' Red
            wsDest.cells(rowNum, 4).Interior.Color = RGB(0, 160, 0) ' Green
        Else
            wsDest.cells(rowNum, 4).Interior.Color = RGB(160, 0, 0) ' Red
            wsDest.cells(rowNum, 3).Interior.Color = RGB(0, 160, 0) ' Green
        End If
        wsDest.cells(rowNum, 3).Font.Color = RGB(255, 255, 255)
        wsDest.cells(rowNum, 4).Font.Color = RGB(255, 255, 255)
        wsDest.cells(rowNum, 3).Font.Bold = True
        wsDest.cells(rowNum, 4).Font.Bold = True

        ' Format Col Q (17) - Trend
        If wsDest.cells(rowNum, 17).value = "Bullish" Then
            With wsDest.cells(rowNum, 17)
                .Interior.Color = RGB(204, 255, 204) ' Light Green
                .Font.Bold = True
            End With
        ElseIf wsDest.cells(rowNum, 17).value = "Bearish" Then
            With wsDest.cells(rowNum, 17)
                .Interior.Color = RGB(255, 204, 204) ' Light Red
                .Font.Bold = True
            End With
        End If

        ' Format Col X (24) - VLOOKUP Result
        If wsDest.cells(rowNum, 24).value = "Up" Then
            With wsDest.cells(rowNum, 24)
                .Interior.Color = RGB(204, 255, 204) ' Light Green
                .Font.Bold = True
            End With
        ElseIf wsDest.cells(rowNum, 24).value = "Down" Then
            With wsDest.cells(rowNum, 24)
                .Interior.Color = RGB(255, 204, 204) ' Light Red
                .Font.Bold = True
            End With
        End If
        
        ' Format Cols M (13) to P (16) - The checked RSI columns
        For c = 13 To 16
            Select Case wsDest.cells(rowNum, c).value
                Case Is >= 80
                    wsDest.cells(rowNum, c).Interior.Color = RGB(210, 180, 140) ' Tan/Brownish
                Case 70 To 79, 21 To 30
                    wsDest.cells(rowNum, c).Interior.Color = RGB(255, 255, 160) ' Yellow
                Case Is <= 20
                    wsDest.cells(rowNum, c).Interior.Color = RGB(0, 200, 0) ' Green
            End Select
        Next c
        
        ' Note: The complex Pivot Level highlighting (Cols 6-12) logic is kept intact
        ' as it uses existing data (Col C vs Cols F-L) and Col B (which we now assign).

        ' === Pivot Levels: Buy Logic ===
'        If LCase(Trim(wsDest.Cells(rowNum, 2).value)) = "buy" Then
'            For c = 6 To 12
'                If wsDest.Cells(rowNum, 3).value > wsDest.Cells(rowNum, c).value Then
'                    ' GoTo skipcol ' GoTo is generally discouraged, replacing with conditional logic
'                Else
'                    ' c is the first value >= Col C (current price), so highlight c-1, c, c-2
'                    wsDest.Cells(rowNum, c - 1).Interior.Color = RGB(204, 255, 204) ' Light Green
'                    wsDest.Cells(rowNum, c - 1).Font.Bold = True
'                    wsDest.Cells(rowNum, c).Interior.Color = RGB(255, 255, 160) ' Yellow
'                    wsDest.Cells(rowNum, c).Font.Bold = True
'                    wsDest.Cells(rowNum, c - 2).Interior.Color = RGB(255, 204, 204) ' Light Red
'                    wsDest.Cells(rowNum, c - 2).Font.Bold = True
'                    Exit For
'                End If
'            Next c
'        End If
'
'        ' === Pivot Levels: Sell Logic ===
'        If LCase(Trim(wsDest.Cells(rowNum, 2).value)) = "sell" Then
'            For c = 12 To 6 Step -1
'                If wsDest.Cells(rowNum, 3).value < wsDest.Cells(rowNum, c).value Then
'                    ' GoTo skipCol1 ' GoTo is generally discouraged, replacing with conditional logic
'                Else
'                    ' c is the first value <= Col C (current price) when iterating backwards, so highlight c+1, c, c+2
'                    wsDest.Cells(rowNum, c + 1).Interior.Color = RGB(204, 255, 204) ' Light Green
'                    wsDest.Cells(rowNum, c + 1).Font.Bold = True
'                    wsDest.Cells(rowNum, c).Interior.Color = RGB(255, 255, 160) ' Yellow
'                    wsDest.Cells(rowNum, c).Font.Bold = True
'                    wsDest.Cells(rowNum, c + 2).Interior.Color = RGB(255, 204, 204) ' Light Red
'                    wsDest.Cells(rowNum, c + 2).Font.Bold = True
'                    Exit For
'                End If
'            Next c
'        End If
        If LCase(Trim(wsDest.cells(rowNum, 2).value)) = "buy" Then
            For c = 8 To 12   ' shifted start
                If wsDest.cells(rowNum, 3).value <= wsDest.cells(rowNum, c).value Then
                    wsDest.cells(rowNum, c).Interior.Color = RGB(255, 255, 160)
                    wsDest.cells(rowNum, c).Font.Bold = True
        
                    wsDest.cells(rowNum, c - 1).Interior.Color = RGB(204, 255, 204)
                    wsDest.cells(rowNum, c - 1).Font.Bold = True
        
                    wsDest.cells(rowNum, c - 2).Interior.Color = RGB(255, 204, 204)
                    wsDest.cells(rowNum, c - 2).Font.Bold = True
        
                    Exit For
                End If
            Next c
        End If

        If LCase(Trim(wsDest.cells(rowNum, 2).value)) = "sell" Then
            For c = 10 To 6 Step -1   ' shifted end
                If wsDest.cells(rowNum, 3).value >= wsDest.cells(rowNum, c).value Then
                    wsDest.cells(rowNum, c).Interior.Color = RGB(255, 255, 160)
                    wsDest.cells(rowNum, c).Font.Bold = True
        
                    wsDest.cells(rowNum, c + 1).Interior.Color = RGB(204, 255, 204)
                    wsDest.cells(rowNum, c + 1).Font.Bold = True
        
                    wsDest.cells(rowNum, c + 2).Interior.Color = RGB(255, 204, 204)
                    wsDest.cells(rowNum, c + 2).Font.Bold = True
        
                    Exit For
                End If
            Next c
        End If
    Next r
End Sub

Sub AppendNewLogicData()
'
' AppendNewLogicData Macro
' Copies data from New_Logic_Intraday and appends to NL_ID_Cumulative.
'

' Declare variables
Dim sourceSheet As Worksheet
Dim destSheet As Worksheet
Dim lastRowSource As Long
Dim lastRowDest As Long
Dim copyRange As Range
Dim newRows As Long
    
    ' Set error handling to gracefully exit if sheets are not found
    On Error GoTo SheetNotFound
    
    ' Define the source and destination worksheets
    Set sourceSheet = ThisWorkbook.Worksheets("New_Logic_Intraday")
    Set destSheet = ThisWorkbook.Worksheets("NL_ID_Cumulative")
    
    ' Disable screen updating for faster execution
    Application.ScreenUpdating = False
    
    ' Find the last used row in the source sheet in column A
    ' This ensures we copy all data, not just a fixed range
    lastRowSource = sourceSheet.cells(sourceSheet.Rows.count, "A").End(xlUp).row
    
    ' Check if there is data to copy (starting from row 3)
    If lastRowSource < 3 Then
        MsgBox "No data to copy in 'New_Logic_Intraday' starting from row 3.", vbInformation, "No Data"
        GoTo ExitSub
    End If
    
    ' Find the last used row in the destination sheet in column A
    lastRowDest = destSheet.cells(destSheet.Rows.count, "A").End(xlUp).row
    
    ' Define the range to copy (from row 3 to the last used row, columns A to W)
    Set copyRange = sourceSheet.Range("A3:X" & lastRowSource)
    
    ' Calculate the number of rows to copy
    newRows = lastRowSource - 2
    ' Assign data directly from source sheet (bypassing OS clipboard)
    destSheet.Range("A" & lastRowDest + 1).Resize(newRows, copyRange.Columns.count).value = copyRange.value
    
    ' Add a datetime stamp to the newly appended rows in column Y
    destSheet.Range("Y" & lastRowDest + 1).Resize(newRows, 1).value = Now()
    
    ' Write COCKPIT status for each NI row to column Z
    Dim statusArr As Variant
    On Error Resume Next
    statusArr = Module_Cockpit.mStatusArray
    On Error GoTo 0
    If Not IsEmpty(statusArr) Then
        Dim i As Long
        For i = 1 To newRows
            If i <= UBound(statusArr) Then
                destSheet.cells(lastRowDest + i, 26).value = statusArr(i)
            End If
        Next i
    End If
    destSheet.Range("Z1").value = "COCKPIT_Status"
    
    ' Clear the clipboard to avoid the "marching ants" border
    Application.CutCopyMode = False
    
    ' Inform the user that the operation was successful
    'MsgBox "Data successfully copied from 'New_Logic_Intraday' and appended to 'NL_ID_Cumulative'.", vbInformation, "Success"
    
ExitSub:
    ' Re-enable screen updating
    ' Note: ScreenUpdating managed by caller (RunAll)
    Exit Sub
    
SheetNotFound:
    ' Handle the error if one of the sheets is not found
    MsgBox "Error: One or both of the required worksheets ('New_Logic_Intraday' or 'NL_ID_Cumulative') were not found.", vbCritical, "Sheet Not Found"
    GoTo ExitSub

End Sub
Sub UpdateSignalsSheet()

    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Dim srcData As Variant, outData As Variant
    Dim lastRow As Long, r As Long, outRow As Long
    Dim score As Double, absScore As Double
    Dim ltp As Double
    Dim swingHigh As Double, swingLow As Double
    Dim risk As Double
    Dim narrative As String

    ' --- CONFIGURATION ---
    Set wsSrc = ThisWorkbook.Sheets("Sheet1")

    On Error Resume Next
    Set wsTgt = ThisWorkbook.Sheets("Signals")
    On Error GoTo 0
    If wsTgt Is Nothing Then
        Set wsTgt = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsTgt.Name = "Signals"
    End If

    ' 1. READ SOURCE DATA
    lastRow = wsSrc.cells(wsSrc.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub

    srcData = wsSrc.Range("A2:DG" & lastRow).value

    ' 2. PREPARE OUTPUT
    ReDim outData(1 To UBound(srcData, 1), 1 To 21)
    outRow = 0

    ' 3. LOOP
    For r = 1 To UBound(srcData, 1)

        score = val(srcData(r, 109))
        absScore = Abs(score)

        If absScore >= 60 Then
            outRow = outRow + 1

            ' --- A. Ticker ---
            outData(outRow, 1) = srcData(r, 1)

            ' --- B. Action ---
            If score > 0 Then
                outData(outRow, 2) = "BUY"
            Else
                outData(outRow, 2) = "SELL"
            End If

            ' --- C. Daily Direction ---
            outData(outRow, 3) = srcData(r, 111)

            ' --- D. LTP ---
            ltp = val(srcData(r, 20))
            outData(outRow, 4) = ltp

            ' --- E. Entry / Target / SL (UNCHANGED) ---
            swingHigh = val(srcData(r, 93))
            swingLow = val(srcData(r, 94))

            If score > 0 Then
                outData(outRow, 5) = swingHigh
                outData(outRow, 7) = swingLow
                If swingHigh > 0 And swingLow > 0 Then
                    risk = swingHigh - swingLow
                    outData(outRow, 6) = swingHigh + (risk * 1.5)
                End If
            Else
                outData(outRow, 5) = swingLow
                outData(outRow, 7) = swingHigh
                If swingHigh > 0 And swingLow > 0 Then
                    risk = swingHigh - swingLow
                    outData(outRow, 6) = swingLow - (risk * 1.5)
                End If
            End If

            ' --- H. Conf Score ---
            outData(outRow, 8) = score

            ' =========================
            ' I. NARRATIVE (NEW, HUMAN)
            ' =========================

'            If score > 0 Then
'                If ltp < swingHigh Then
'                    narrative = _
'                        "WAIT — Breakout not triggered yet" & vbLf & _
'                        "Bullish trend, but resistance at " & Format(swingHigh, "0.00") & _
'                        " is holding. Wait for a clean move above this level."
'                Else
'                    narrative = _
'                        "BUY — Breakout confirmed" & vbLf & _
'                        "Price has broken resistance. Buy near " & Format(swingHigh, "0.00") & _
'                        "; exit if price falls below " & Format(swingLow, "0.00") & "."
'                End If
'            Else
'                If ltp > swingLow Then
'                    narrative = _
'                        "WAIT — Breakdown not triggered yet" & vbLf & _
'                        "Bearish trend, but support at " & Format(swingLow, "0.00") & _
'                        " is holding. Wait for a clean break below this level."
'                Else
'                    narrative = _
'                        "SELL — Breakdown confirmed" & vbLf & _
'                        "Price has broken support. Sell near " & Format(swingLow, "0.00") & _
'                        "; exit if price moves above " & Format(swingHigh, "0.00") & "."
'                End If
'            End If
            If score > 0 Then
                If ltp < swingHigh Then
                    narrative = _
                        "WAIT — Breakout not triggered yet!..  " & "Bullish trend, but resistance at " & Format(swingHigh, "0.00") & _
                        " is holding. Wait for a clean move above this level."
                Else
                    narrative = _
                        "BUY — Breakout confirmed!..  " & vbLf & "Price has broken resistance. Buy near " & Format(swingHigh, "0.00") & _
                        "; exit if price falls below " & Format(swingLow, "0.00") & "."
                End If
            Else
                If ltp > swingLow Then
                    narrative = _
                        "WAIT — Breakdown not triggered yet!..  " & "Bearish trend, but support at " & Format(swingLow, "0.00") & _
                        " is holding. Wait for a clean break below this level."
                Else
                    narrative = _
                        "SELL — Breakdown confirmed!..  " & "Price has broken support. Sell near " & Format(swingLow, "0.00") & _
                        "; exit if price moves above " & Format(swingHigh, "0.00") & "."
                End If
            End If


            outData(outRow, 9) = narrative

            ' --- PIVOTS ---
            outData(outRow, 10) = srcData(r, 11)
            outData(outRow, 11) = srcData(r, 12)
            outData(outRow, 12) = srcData(r, 13)
            outData(outRow, 13) = srcData(r, 14)
            outData(outRow, 14) = srcData(r, 15)
            outData(outRow, 15) = srcData(r, 16)
            outData(outRow, 16) = srcData(r, 17)

            ' --- RSI & SUPERTREND ---
            outData(outRow, 17) = srcData(r, 21)
            outData(outRow, 18) = srcData(r, 22)
            outData(outRow, 19) = srcData(r, 23)
            outData(outRow, 20) = srcData(r, 24)
            outData(outRow, 21) = srcData(r, 72)

        End If
    Next r

    ' 4. WRITE TO DESTINATION (ORIGINAL LOGIC)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    With wsTgt
        If .cells(.Rows.count, "A").End(xlUp).row >= 4 Then
            .Range("A4:U" & .cells(.Rows.count, "A").End(xlUp).row).ClearContents
            .Range("A4:U" & .Rows.count).Interior.ColorIndex = xlNone
            .Range("A4:U" & .Rows.count).Font.Color = 0
            .Range("A4:U" & .Rows.count).Font.Bold = False
        End If

        If outRow > 0 Then
            .Range("A4").Resize(outRow, 21).value = outData
            .Range("A4").Resize(outRow, 21).NumberFormat = "0.00"
            .Columns("A:U").AutoFit
        End If
    End With

    ' 5. FORMATTING (UNCHANGED — YOUR ORIGINAL)
    For r = 4 To outRow + 3
        Select Case LCase(Trim(wsTgt.cells(r, 2).value))
            Case "buy"
                wsTgt.cells(r, 2).Interior.Color = RGB(204, 255, 204)
                wsTgt.cells(r, 8).Interior.Color = RGB(204, 255, 204)
            Case "sell"
                wsTgt.cells(r, 2).Interior.Color = RGB(160, 0, 0)
                wsTgt.cells(r, 2).Font.Color = RGB(255, 255, 255)
                wsTgt.cells(r, 8).Interior.Color = RGB(160, 0, 0)
                wsTgt.cells(r, 8).Font.Color = RGB(255, 255, 255)
            Case "watch", "watch (s)"
                wsTgt.cells(r, 2).Interior.Color = RGB(255, 255, 160)
                wsTgt.cells(r, 8).Interior.Color = RGB(255, 255, 160)
        End Select

        If LCase(wsTgt.cells(r, 3).value) = "bullish" Then
            wsTgt.cells(r, 3).Interior.Color = RGB(204, 255, 204)
        ElseIf LCase(wsTgt.cells(r, 3).value) = "bearish" Then
            wsTgt.cells(r, 3).Interior.Color = RGB(160, 0, 0)
            wsTgt.cells(r, 3).Font.Color = RGB(255, 255, 255)
        End If
    Next r

    ' --- YOUR PIVOT & NARRATIVE BANDS (UNCHANGED) ---
    outRow = outRow + 3

    Dim rng As Range
    Set rng = wsTgt.Range("J4:J" & outRow & ",L4:L" & outRow & _
                          ",N4:N" & outRow & ",P4:P" & outRow & _
                          ",R4:R" & outRow & ",T4:T" & outRow)

    With rng.Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.35
    End With

    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True

    Set rng = wsTgt.Range("I4:I" & outRow)
    With rng.Interior
        .Pattern = xlSolid
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = -0.35
    End With

    rng.Font.Color = RGB(255, 255, 255)
    rng.Font.Bold = True
    
    Dim txt As String
    Dim actionWord As String
    
    For r = 4 To outRow + 3
    
        txt = wsTgt.cells(r, 9).value
        If Len(txt) = 0 Then GoTo NextNarrative
    
        ' Detect action keyword at start
        If Left(txt, 4) = "WAIT" Then
            actionWord = "WAIT"
            With wsTgt.cells(r, 9).Characters(1, Len(actionWord)).Font
                .Color = RGB(255, 192, 0)   ' Yellow
                .Bold = True
            End With
    
        ElseIf Left(txt, 3) = "BUY" Then
            actionWord = "BUY"
            With wsTgt.cells(r, 9).Characters(1, Len(actionWord)).Font
                .Color = RGB(0, 176, 80)    ' Green
                .Bold = True
            End With
    
        ElseIf Left(txt, 4) = "SELL" Then
            actionWord = "SELL"
            With wsTgt.cells(r, 9).Characters(1, Len(actionWord)).Font
                .Color = RGB(192, 0, 0)     ' Red
                .Bold = True
            End With
        End If
    
NextNarrative:
    Next r


    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


Sub UpdateSignalsSheet_Gemini()
    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Dim srcData As Variant, outData As Variant
    Dim lastRow As Long, r As Long, outRow As Long
    Dim score As Double, absScore As Double
    Dim ltp As Double
    Dim swingHigh As Double, swingLow As Double
    Dim risk As Double
    
    ' --- CONFIGURATION ---
    Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    
    ' Check if Signals sheet exists, create if not
    On Error Resume Next
    Set wsTgt = ThisWorkbook.Sheets("Signals_Gemini")
    On Error GoTo 0
    If wsTgt Is Nothing Then
        Set wsTgt = ThisWorkbook.Sheets.Add(After:=wsSrc)
        wsTgt.Name = "Signals"
    End If
    
    ' 1. READ SOURCE DATA (FAST)
    lastRow = wsSrc.cells(wsSrc.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub
    
    ' Load block. NOTE: Reading up to DG (111) to capture Daily Direction if needed
    srcData = wsSrc.Range("A2:DG" & lastRow).value
    
    ' 2. PREPARE OUTPUT ARRAY
    ReDim outData(1 To UBound(srcData, 1), 1 To 21)
    outRow = 0
    
    ' 3. LOOP AND PROCESS
    For r = 1 To UBound(srcData, 1)
        
        ' Column 109 is DE (Conf_Score)
        score = val(srcData(r, 109))
        absScore = Abs(score) ' Use Absolute value for filtering
        
        ' FILTER: Process if Absolute Score > 60 (Handles Buy > 60 and Sell < -60)
        If absScore >= 80 Then
            outRow = outRow + 1
            
            ' --- A. Ticker ---
            outData(outRow, 1) = srcData(r, 1)
            
            ' --- B. Action (Derived) - DUAL POLARITY ---
            If score > 0 Then
                ' --- BUY LOGIC ---
                If score >= 80 Then
                    outData(outRow, 2) = "BUY"
                Else
                    outData(outRow, 2) = "WATCH"
                End If
            Else
                ' --- SELL LOGIC ---
                If score <= -80 Then
                    outData(outRow, 2) = "SELL"
                Else
                    outData(outRow, 2) = "WATCH (S)"
                End If
            End If
            
            ' --- C. Daily Direction ---
            ' Reading Column 111 (DG) if Python is fixed, else 36 (AJ)
            ' Assuming you want the Python "State" logic here eventually.
            ' For now, keeping your mapping (AJ = 36)
            'outData(outRow, 3) = srcData(r, 36)
            outData(outRow, 3) = srcData(r, 111)
            
            ' --- D. LTP ---
            ltp = val(srcData(r, 20))
            outData(outRow, 4) = ltp
            
            ' --- E. Entry & SL Logic (Inverted for Sell?) ---
            swingHigh = val(srcData(r, 93))
            swingLow = val(srcData(r, 94))
            
            If score > 0 Then
                ' BUY SETUP
                outData(outRow, 5) = swingHigh ' Entry = Break High
                outData(outRow, 7) = swingLow  ' SL = Low
                
                ' Target (1.5 RR)
                If swingHigh > 0 And swingLow > 0 Then
                    risk = swingHigh - swingLow
                    outData(outRow, 6) = swingHigh + (risk * 1.5)
                End If
            Else
                ' SELL SETUP (Inverse)
                outData(outRow, 5) = swingLow  ' Entry = Break Low
                outData(outRow, 7) = swingHigh ' SL = High
                
                ' Target (1.5 RR Downside)
                If swingHigh > 0 And swingLow > 0 Then
                    risk = swingHigh - swingLow
                    outData(outRow, 6) = swingLow - (risk * 1.5)
                End If
            End If
            
            ' --- H. Conf Score ---
            outData(outRow, 8) = score
            
            ' --- I. Narrative ---
            outData(outRow, 9) = srcData(r, 110)
            
            ' --- PIVOTS ---
            outData(outRow, 10) = srcData(r, 11) ' S3
            outData(outRow, 11) = srcData(r, 12) ' S2
            outData(outRow, 12) = srcData(r, 13) ' S1
            outData(outRow, 13) = srcData(r, 14) ' PP
            outData(outRow, 14) = srcData(r, 15) ' R1
            outData(outRow, 15) = srcData(r, 16) ' R2
            outData(outRow, 16) = srcData(r, 17) ' R3
            
            ' --- RSI & SUPERTREND ---
            outData(outRow, 17) = srcData(r, 21)
            outData(outRow, 18) = srcData(r, 22)
            outData(outRow, 19) = srcData(r, 23)
            outData(outRow, 20) = srcData(r, 24)
            outData(outRow, 21) = srcData(r, 72)
            
        End If
    Next r
    
    ' 4. WRITE TO DESTINATION
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    With wsTgt
        If .cells(.Rows.count, "A").End(xlUp).row >= 4 Then
            .Range("A4:U" & .cells(.Rows.count, "A").End(xlUp).row).ClearContents
            .Range("A4:U" & .Rows.count).Interior.ColorIndex = xlNone
            .Range("A4:U" & .Rows.count).Font.Color = 0
            .Range("A4:U" & .Rows.count).Font.Bold = False
        End If
        
        If outRow > 0 Then
            .Range("A4").Resize(outRow, 21).value = outData
            .Range("A4").Resize(outRow, 21).NumberFormat = "0.00"
            .Columns("A:U").AutoFit
        End If
    End With
    
    ' 5. FORMATTING (Now handles SELL colors)
    For r = 4 To outRow + 3
        Select Case LCase(Trim(CStr(wsTgt.cells(r, 2).value)))
            Case "buy"

                
                wsTgt.cells(r, 2).Interior.Color = RGB(204, 255, 204)
                wsTgt.cells(r, 2).Font.Color = RGB(0, 100, 0)
                
                wsTgt.cells(r, 8).Interior.Color = RGB(204, 255, 204)
                wsTgt.cells(r, 8).Font.Color = RGB(0, 100, 0)
                
            Case "sell"

                
                wsTgt.cells(r, 2).Interior.Color = RGB(160, 0, 0)
                wsTgt.cells(r, 2).Font.Color = RGB(255, 255, 255)
                
                wsTgt.cells(r, 8).Interior.Color = RGB(160, 0, 0)
                wsTgt.cells(r, 8).Font.Color = RGB(255, 255, 255)
                
            Case "watch", "watch (s)"
                wsTgt.cells(r, 2).Interior.Color = RGB(255, 255, 160) ' Yellow
                wsTgt.cells(r, 2).Font.Color = RGB(0, 0, 0)
                
                wsTgt.cells(r, 8).Interior.Color = RGB(255, 255, 160) ' Yellow
                wsTgt.cells(r, 8).Font.Color = RGB(0, 0, 0)
                
            Case "weak"
                wsTgt.cells(r, 2).Font.Color = RGB(128, 128, 128)
        End Select

        If LCase(wsTgt.cells(r, 3).value) = "bullish" Then
            wsTgt.cells(r, 3).Interior.Color = RGB(204, 255, 204)
            wsTgt.cells(r, 3).Font.Color = RGB(0, 150, 0)

        ElseIf LCase(wsTgt.cells(r, 3).value) = "bearish" Then
            wsTgt.cells(r, 3).Interior.Color = RGB(160, 0, 0)
            wsTgt.cells(r, 3).Font.Color = RGB(255, 255, 255)
        End If
    Next r
    
        outRow = outRow + 3
        
        Dim rng As Range

        
        
        Set rng = wsTgt.Range("J4:J" & outRow & ",L4:L" & outRow & ",N4:N" & outRow & ",P4:P" & outRow & ",R4:R" & outRow & ",T4:T" & outRow)
        
        With rng.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.349986266670736
        End With
        
        rng.Font.Color = RGB(255, 255, 255)
        rng.Font.Bold = True
        
        
        Set rng = wsTgt.Range("I4:I" & outRow)
        With rng.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight2
            .TintAndShade = -0.349986266670736
        End With
        rng.Font.Color = RGB(255, 255, 255)
        rng.Font.Bold = True

    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub


Function Is_Index_Ticker(t As String) As Boolean
    Select Case UCase(t)
        Case "NSEI", "NSEBANK", "CNXFIN", "CNXAUTO", "CNXIT", _
             "CNXPHARMA", "CNXFMCG", "CNXMETAL", "CNXENERGY", _
             "CNXINFRA", "CNXREALTY", "CNXMEDIA", "CNXCMDT", _
             "CNXCONSUM", "CNXPSE", "CNXPSUBANK"
            Is_Index_Ticker = True
    End Select
End Function

