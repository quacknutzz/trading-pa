Attribute VB_Name = "Module2"
' ---- SAFE NUMERIC HELPERS ----
Private Function NF_IsNum(v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Trim$(CStr(v)) = "" Then Exit Function
        NF_IsNum = IsNumeric(v)
    Else
        NF_IsNum = IsNumeric(v)
    End If
End Function

Private Function NF_GetNum(v As Variant, Optional d As Double = 0#) As Double
    If IsError(v) Or IsEmpty(v) Then NF_GetNum = d: Exit Function
    If NF_IsNum(v) Then NF_GetNum = CDbl(v) Else NF_GetNum = d
End Function

Public Function NF_RowHasLiveOHLCV(arr As Variant, rowIdx As Long) As Boolean
    On Error Resume Next
    If arr(rowIdx, 3) <> "" And arr(rowIdx, 4) <> "" And _
       arr(rowIdx, 5) <> "" And arr(rowIdx, 6) <> "" And _
       arr(rowIdx, 7) <> "" Then
        NF_RowHasLiveOHLCV = True
    Else
        NF_RowHasLiveOHLCV = False
    End If
    On Error GoTo 0
End Function

Public Function NF_RowHasDailyOHLCV(arr As Variant, rowIdx As Long) As Boolean
    On Error Resume Next
    If arr(rowIdx, 10) <> "" And arr(rowIdx, 11) <> "" And _
       arr(rowIdx, 12) <> "" And arr(rowIdx, 13) <> "" And _
       arr(rowIdx, 14) <> "" Then
        NF_RowHasDailyOHLCV = True
    Else
        NF_RowHasDailyOHLCV = False
    End If
    On Error GoTo 0
End Function

'-----------------------------------------------------------------
' Builds Step1 outputs from Sheet1 + Trades
' Returns a 2D array outArr = {Symbol, Date, Open, High, Low, Close, Gate, Score, Narrative}
'-----------------------------------------------------------------
'=============================================================
' Build NF_Outputs sheet from Sheet1 + Trades arrays
'=============================================================
Public Sub NF_BuildOutputs_Step1()
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim srcArr As Variant, tradeArr As Variant, whichTrades As String
    Dim startRow As Long, endRow As Long, lastRowSrc As Long
    Dim outArr() As Variant
    Dim i As Long, outRow As Long
    Dim b5 As Variant, b15 As Variant, b1h As Variant
    Dim gate As Boolean, tradeRow As Long
    Dim res As Variant, score As Long, narrative As String
    
    Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Set wsOut = ThisWorkbook.Sheets("NF_Outputs")
    
    startRow = CLng(wsSrc.Range("P1").value)
    endRow = CLng(wsSrc.Range("R1").value)
    If startRow <= 1 Or endRow <= 1 Or endRow < startRow Then Exit Sub
    
    lastRowSrc = wsSrc.cells(wsSrc.Rows.count, "A").End(xlUp).row
    If endRow > lastRowSrc Then endRow = lastRowSrc
    
    srcArr = wsSrc.Range("A1:DD" & endRow).value
    tradeArr = NF_LoadTradeArray(whichTrades)
    
    wsOut.cells.Clear
    'wsOut.Range("A1:E" & Rows.count).Clear
    
    wsOut.Range("A1:E1").value = Array("Symbol", "Score", "Narrative", "BiasGate", "RowIdx")
    ReDim outArr(1 To (endRow - startRow + 1), 1 To 5)
    outRow = 0
    
    For i = startRow To endRow
        If Trim$(CStr(srcArr(i, 1))) = "" Then GoTo SkipRow
        
        NF_GetBiasTripletFromArr wsSrc, srcArr, i, b5, b15, b1h
        
        gate = NF_DecideGate(b5, b15, b1h)
        If Not gate Then GoTo SkipRow
        
        tradeRow = NF_FindTradeRow(tradeArr, srcArr(i, 1))
        
        res = NF_ScoreRow(srcArr, tradeArr, i, tradeRow, b5, b15, b1h)
        score = CLng(res(1))
        narrative = CStr(res(2))
        
        outRow = outRow + 1
        outArr(outRow, 1) = srcArr(i, 1)
        outArr(outRow, 2) = score
        outArr(outRow, 3) = narrative
        outArr(outRow, 4) = True
        outArr(outRow, 5) = i
        'Debug.Print "Skipping " & srcArr(i, 1) & " gate=" & gate & " b5=" & b5 & " b15=" & b15 & " b1h=" & b1h

SkipRow:
    Next i
    
    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, 5).value = outArr
    End If
    
    ThisWorkbook.Sheets("NF_Outputs").Select
    'ActiveWindow.SmallScroll Down:=-6
    ThisWorkbook.Sheets("NF_Outputs").Range("F2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISERROR(VLOOKUP(RC[-5], New_Trades!C1, 1, FALSE)), """", ""Match"")"
    'Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F108")
    
    
    ThisWorkbook.Sheets("NF_Outputs").Select
    ActiveSheet.Range("$A$1:$F$200").AutoFilter Field:=2, Criteria1:=">=6", _
        Operator:=xlOr, Criteria2:="<=-6"
        
    ThisWorkbook.Sheets("NF_Outputs").Select
    ActiveSheet.Range("$A$1:$F$200").AutoFilter Field:=6, Criteria1:="=Match"
End Sub
Sub NF_CopyIndicators()
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Set wsOut = ThisWorkbook.Sheets("NF_Outputs")
    
    Dim lastRow As Long
    lastRow = wsSrc.cells(wsSrc.Rows.count, "A").End(xlUp).row
    
    ' === ATRs ===
    wsOut.Range("H1").value = "ATR5m"
    wsOut.Range("H2:H" & lastRow).value = wsSrc.Range("CC2:CC" & lastRow).value
    
    wsOut.Range("I1").value = "ATR15m"
    wsOut.Range("I2:I" & lastRow).value = wsSrc.Range("CD2:CD" & lastRow).value
    
    ' === VWAP ===
    wsOut.Range("J1").value = "VWAP"
    wsOut.Range("J2:J" & lastRow).value = wsSrc.Range("BR2:BR" & lastRow).value
    
    ' === Volume Spike ===
    wsOut.Range("K1").value = "VolSpike"
    wsOut.Range("K2:K" & lastRow).value = wsSrc.Range("BS2:BS" & lastRow).value
    
    MsgBox "ATR, VWAP, VolSpike copied to NF_Outputs", vbInformation
End Sub



'=============================================================
' Build NF_Outputs sheet from Sheet1 + Trades arrays (full build)
'=============================================================
Public Sub NF_BuildOutputs()
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim srcArr As Variant, tradeArr As Variant, whichTrades As String
    Dim lastRowSrc As Long, startRow As Long, endRow As Long
    Dim outArr() As Variant
    Dim i As Long, outRow As Long, j As Long
    Dim b5 As Variant, b15 As Variant, b1h As Variant
    Dim gate As Boolean, tradeRow As Long
    Dim res As Variant, score As Long, narrative As String
    Dim val, val1 As Variant
    
    Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Set wsOut = ThisWorkbook.Sheets("NF_Outputs")
    
    ' Row boundaries from P1/R1
    startRow = CLng(wsSrc.Range("P1").value)
    endRow = CLng(wsSrc.Range("R1").value)
    If startRow <= 1 Or endRow <= 1 Or endRow < startRow Then
        MsgBox "Invalid boundaries in Sheet1!P1/R1", vbExclamation
        Exit Sub
    End If
    
    ' Load Sheet1 data A:DD up to endRow (row 1 = headers)
    lastRowSrc = wsSrc.cells(wsSrc.Rows.count, "A").End(xlUp).row
    If endRow > lastRowSrc Then endRow = lastRowSrc
    srcArr = wsSrc.Range("A1:DD" & endRow).value
    
    ' Load trades array (New_Trades row2 preferred, else Trades row3)
    tradeArr = NF_LoadTradeArray(whichTrades)
    
    ' Prep output
    wsOut.cells.Clear
    wsOut.Range("A1:Q1").value = Array("Symbol", "Score", "Narrative", "BiasGate", "RowIdx", "Match", "Bol", "Entry", "Exit", "SL", "S3", "S2", "S1", "PP", "R1", "R2", "R3")
    
    ReDim outArr(1 To (endRow - startRow + 1), 1 To 17)
    outRow = 0
    
    ' Loop through bounded rows
    For i = startRow To endRow
        If Trim$(CStr(srcArr(i, 1))) = "" Then GoTo SkipRow
        
        ' Biases
        NF_GetBiasTripletFromArr wsSrc, srcArr, i, b5, b15, b1h
        
        ' Gate
        gate = NF_DecideGate(b5, b15, b1h)
        If Not gate Then GoTo SkipRow
        
        ' Find matching trade row (symbol in col 1)
        tradeRow = NF_FindTradeRow(tradeArr, srcArr(i, 1))
        
        ' Score + narrative (NF_ScoreRow returns Variant(1..2): score, narrative)
        res = NF_ScoreRow(srcArr, tradeArr, i, tradeRow, b5, b15, b1h)
        score = CLng(res(1))
        narrative = CStr(res(2))
        
        ' Collect
        outRow = outRow + 1
       
        outArr(outRow, 1) = srcArr(i, 1)   ' Symbol
        outArr(outRow, 2) = score
        outArr(outRow, 3) = narrative
        outArr(outRow, 4) = gate
        outArr(outRow, 5) = i

'
        val = "=IF(ISERROR(VLOOKUP(""" & outArr(outRow, 1) & """, New_Trades!$A:$B, 2, FALSE)), """", VLOOKUP(""" & outArr(outRow, 1) & """, New_Trades!$A:$B, 2, FALSE))"
        val1 = Application.Evaluate(val)
        
        outArr(outRow, 6) = val1
        
        If score >= 6 Then
            If LCase(LTrim(val1)) = "sell" Then
                outArr(outRow, 7) = "0"
            ElseIf LCase(LTrim(val1)) = "buy" Then
                outArr(outRow, 7) = "1"
            Else
                outArr(outRow, 7) = ""
            End If

        ElseIf score <= -6 Then
            If LCase(LTrim(val1)) = "buy" Then
                outArr(outRow, 7) = "0"
            ElseIf LCase(LTrim(val1)) = "sell" Then
                outArr(outRow, 7) = "1"
            Else
                outArr(outRow, 7) = ""
            End If
        End If
        
        For j = 8 To 17
            val = "=IF(ISERROR(VLOOKUP(""" & outArr(outRow, 1) & """, New_Trades!$A:$L," & j - 5 & ", FALSE)), """", VLOOKUP(""" & outArr(outRow, 1) & """, New_Trades!$A:$L," & j - 5 & ", FALSE))"
            val1 = Application.Evaluate(val)
            'If LCase(Trim(val1)) <> "buy" Or LCase(Trim(val1)) <> "sell" Then
            If val1 <> "" Then
                'Debug.Print CLng(val1)
                outArr(outRow, j) = Round(val1, 2)
            Else
                outArr(outRow, j) = val1
            End If
        Next j
        
        
SkipRow:
    Next i
    
    If outRow > 0 Then
        wsOut.Range("A2").Resize(outRow, 17).value = outArr
    End If
    
    For i = 2 To outRow + 1
        If LCase(LTrim(ThisWorkbook.Sheets("NF_Outputs").Range("B" & i).value)) >= 6 Then
            ThisWorkbook.Sheets("NF_Outputs").Range("A" & i).Font.Color = RGB(0, 160, 0)
            ThisWorkbook.Sheets("NF_Outputs").Range("B" & i).Font.Color = RGB(0, 160, 0)
            ThisWorkbook.Sheets("NF_Outputs").Range("A" & i).Font.Bold = True
            ThisWorkbook.Sheets("NF_Outputs").Range("B" & i).Font.Bold = True
            ThisWorkbook.Sheets("NF_Outputs").Range("F" & i).Font.Color = RGB(0, 160, 0)
            ThisWorkbook.Sheets("NF_Outputs").Range("F" & i).Font.Bold = True

            
        ElseIf LCase(LTrim(ThisWorkbook.Sheets("NF_Outputs").Range("B" & i).value)) <= -6 Then
            ThisWorkbook.Sheets("NF_Outputs").Range("A" & i).Font.Color = RGB(200, 0, 0)
            ThisWorkbook.Sheets("NF_Outputs").Range("B" & i).Font.Color = RGB(200, 0, 0)
            ThisWorkbook.Sheets("NF_Outputs").Range("B" & i).Font.Bold = True
            ThisWorkbook.Sheets("New_Trades").Range("A" & i).Font.Bold = True
            ThisWorkbook.Sheets("NF_Outputs").Range("F" & i).Font.Color = RGB(200, 0, 0)
            ThisWorkbook.Sheets("NF_Outputs").Range("F" & i).Font.Bold = True

            
        Else
            ThisWorkbook.Sheets("NF_Outputs").Range("A" & i).Font.Color = RGB(0, 0, 0)
            ThisWorkbook.Sheets("NF_Outputs").Range("A" & i).Font.Bold = False
        End If

        If LCase(LTrim(ThisWorkbook.Sheets("NF_Outputs").Range("G" & i).value)) = 1 Then
            ThisWorkbook.Sheets("NF_Outputs").Range("G" & i).Interior.Color = RGB(200, 255, 200)
        End If
    Next i
    
    ThisWorkbook.Sheets("NF_Outputs").Range("N2:N" & i).Font.Color = RGB(180, 0, 0)
    'ThisWorkbook.Sheets("NF_Outputs").Range("N2:N" & i).Font.Bold = True
    ThisWorkbook.Sheets("NF_Outputs").Range("K2:K" & i).Interior.Color = RGB(200, 200, 200)
    'ThisWorkbook.Sheets("NF_Outputs").Range("K2:K" & i).Font.Bold = True
    ThisWorkbook.Sheets("NF_Outputs").Range("M2:M" & i).Interior.Color = RGB(200, 200, 200)
    'ThisWorkbook.Sheets("NF_Outputs").Range("M2:M" & i).Font.Bold = True
    ThisWorkbook.Sheets("NF_Outputs").Range("O2:O" & i).Interior.Color = RGB(200, 200, 200)
    'ThisWorkbook.Sheets("NF_Outputs").Range("O2:O" & i).Font.Bold = True
    ThisWorkbook.Sheets("NF_Outputs").Range("Q2:Q" & i).Interior.Color = RGB(200, 200, 200)
    'ThisWorkbook.Sheets("NF_Outputs").Range("Q2:Q" & i).Font.Bold = True
    ThisWorkbook.Sheets("NF_Outputs").Range("K2:Q" & i).Font.Bold = True
    
    
    Sheets("NF_Outputs").Select
    wsOut.Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$E$500").AutoFilter Field:=2, Criteria1:=">=6", _
        Operator:=xlOr, Criteria2:="<=-6"
        

    ActiveSheet.Range("$A$1:$F$500").AutoFilter Field:=6, Criteria1:="<>"
    ActiveSheet.Range("$A$1:$F$500").AutoFilter Field:=7, Criteria1:="1"
    ThisWorkbook.Sheets("NF_Outputs").Range("A2").Select
    'Sheets("New_Trades").Activate
    
End Sub


