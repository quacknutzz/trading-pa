Attribute VB_Name = "Module1"
Public Type IndicatorWeight
    macd As Double
    EMA_10_30 As Double
    EMA_10_50 As Double
    bullishCandle As Double
    bearishCandle As Double
    rsi As Double
    adx As Double
    Stochastic As Double
    williamsR As Double
    roc As Double
    cmf As Double
    superTrend As Double
    psar As Double
    VolumeSpike As Double
End Type


Sub RunAll()
Dim NowTime As Double
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

    'Call Trades

    'Call UpdateSmartNextSteps
    'Debug.Print Now()
    

    Call New_logic_Intraday
    'Debug.Print Now()

    Call New_logic_RSI
    'Debug.Print Now()
    
    'Call Run_Cockpit
    'Debug.Print Now()
    
    'Call AppendNewLogicData
    'Debug.Print Now()
    
    'Call Module_Cockpit_V3.Run_Cockpit_V3     ' V2 Experimental Engine
    
    
    'Call Module_Cockpit_V2.Run_Cockpit_V2     ' V2 Experimental Engine
    
    'Debug.Print Now()
    'Call Module_Cockpit_V2.Run_Cockpit_Compare ' Automated Side-by-Side Comparison
'    Call Run_Paper_Trader
    'Debug.Print Now()
    'Call Module_Cockpit_V4.Run_Cockpit_V4
    
    'Call Module_Cockpit_V5.Run_Cockpit_V5
    
    Call Module_Cockpit_V6.Run_Cockpit_V6
    
    
    ThisWorkbook.Save
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub


Sub FetchOHLCV()
    Dim batchPath As String
    Dim command As String
    
    ' Set batch file path
    batchPath = ThisWorkbook.path ' Removed extra quotes; they're not needed here
    batchPath = batchPath & "\run_data_fetcher.bat"
    
    ' Build command
    command = batchPath
    
    ' Run batch file
    On Error GoTo ErrorHandler
    Shell command, vbNormalFocus
    Exit Sub

ErrorHandler:
    MsgBox "Error running batch file: " & Err.Description, vbCritical
End Sub


Sub FetchRTP_RSI()
    Dim batchPath As String
    Dim command As String

    ' Set batch file path
    'batchPath = """C:\Personal\Trading_100_700\run_fetch_rsi.bat"""
    batchPath = ThisWorkbook.path ' Removed extra quotes; they're not needed here
    batchPath = batchPath & "\run_ai_btb_engine.bat"

    ' Build command
    command = batchPath

    ' Run batch file
    On Error GoTo ErrorHandler
    Shell command, vbNormalFocus

    Exit Sub

ErrorHandler:
    MsgBox "Error running batch file: " & Err.Description, vbCritical
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                              Added to generate scripts for trading                    '
'                              This cript from here on generates Daily                  '
'                              And Intraday scripts for trades                          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Trades()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim srcRow As Long, destRow As Long
    Dim start_row As Long, end_row As Long
    Dim colNum As Long, SupRes As Long
    Dim sourceCell As Range, destCell As Range, overUnderCell As Range, firstValueCell As Range
    Dim cellColor As Long
    Dim hasData, hasData1 As Boolean
    Dim reason As String

    ' Indicator variables
    Dim Price As Double, spike As Double, ATR_5m As Double, superTrend As Double, vwap As Double, atr_15m As Double
    Dim fibHigh As Variant, fibLow As Variant
    Dim bbU_5m, bbL_5m, bbU_15m, bbL_15m
    Dim RSI_5m, rsi_15m, RSI_1H, RSI_1D
    Dim MACD_5m, macd_15m, EMA_5m, EMA_15m, ema_5m_1050, ema_15m_1050
    Dim stochK_5m, stochD_5m, stochK_15m, stochD_15m

    'Debug.Print Now()
    
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsDest = ThisWorkbook.Sheets("Trades")
    destRow = 3

    'ThisWorkbook.ForceFullCalculation = True
    start_row = wsSource.Range("P1").value
    end_row = wsSource.Range("R1").value

    ' Clear old data
    With wsDest
        .Range("A3:Q" & .Rows.count).ClearContents
        .Range("A3:Q" & .Rows.count).Interior.ColorIndex = 0
        .Range("AD3:AR" & .Rows.count).ClearContents
        .Range("A3:Q" & .Rows.count).Interior.ColorIndex = 0
    End With

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
   
    For srcRow = start_row To end_row
        hasData = Application.WorksheetFunction.CountA(wsSource.Range(wsSource.cells(srcRow, 25), wsSource.cells(srcRow, 36))) > 0
        hasData1 = False
        
        hasData1 = Not (wsSource.cells(srcRow, 10).value = "")
        If hasData Or hasData1 Then
            wsDest.cells(destRow, 1).value = wsSource.cells(srcRow, 1).value
            Set firstValueCell = Nothing
            If hasData Then
                For colNum = 0 To 11
                    Set sourceCell = wsSource.cells(srcRow, 25 + colNum)
                    Set destCell = wsDest.cells(destRow, 6 + colNum)
                    Set overUnderCell = wsDest.cells(destRow, 32 + colNum)
    
                    destCell.value = sourceCell.value
                    destCell.Interior.Color = sourceCell.Interior.Color
                    destCell.Font.Color = sourceCell.Font.Color
                    destCell.Font.Bold = sourceCell.Font.Bold
                    destCell.NumberFormat = sourceCell.NumberFormat
    
                    If firstValueCell Is Nothing And Not IsEmpty(sourceCell.value) Then
                        Set firstValueCell = sourceCell
                    End If
    
                    cellColor = sourceCell.Interior.Color
                    Select Case cellColor
                        Case 16711680
                            overUnderCell.value = "Over"
                        Case 255
                            overUnderCell.value = "Under"
                    End Select
                Next colNum
            End If

            ' --- Indicators block for Reason (AD) ---
            reason = ""
            Application.Calculation = xlCalculationAutomatic
            Application.Calculation = xlCalculationManual
            
            Call EvaluateTradeSignal(wsSource, wsDest, srcRow, destRow)
            Call CopyValidTradesToNewSheet(wsSource, srcRow)
            
            destRow = destRow + 1
        End If
    Next srcRow
    
    For r = 2 To destRow + 1
        Select Case LCase(LTrim(ThisWorkbook.Sheets("New_Trades").Range("B" & r).value))
            Case "buy"
                ThisWorkbook.Sheets("New_Trades").Range("A" & r).Font.Color = RGB(0, 150, 0)
            Case "sell"
                ThisWorkbook.Sheets("New_Trades").Range("A" & r).Font.Color = RGB(200, 0, 0)
        End Select
        ThisWorkbook.Sheets("New_Trades").Range("A" & r).Font.Bold = True
    Next r

        HighlightKeywordsFast Sheets("New_Trades").Range("M2:O" & destRow + 1), True

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
Private Sub EvaluateTradeSignal(wsSource As Worksheet, wsDest As Worksheet, srcRow As Long, destRow As Long)
    Dim action As String, reason As String, flag As String
    Dim bullishScore As Integer, bearishScore As Integer
    action = wsDest.cells(destRow, 2).value

    Dim macdLabels As Variant: macdLabels = Array("MACD 5m", "MACD 15m", "MACD 1H")
    Dim emaLabels As Variant: emaLabels = Array("EMA 5m", "EMA 15m", "EMA 1H")

    Dim i As Long
    
    Dim BuySell As String: BuySell = wsSource.cells(srcRow, 10).value
    If BuySell <> "" Then
        If InStr(BuySell, "BUY") > 0 Then
            reason = reason & BuySell & "; " ': bullishScore = bullishScore + 1
        ElseIf InStr(BuySell, "SELL") > 0 Then
            reason = reason & BuySell & "; " ': bearishScore = bearishScore + 1
        End If
    End If
    
    
    ' MACD signals (Cols 32–34)
    For i = 0 To 2
        Select Case LCase(Trim(wsDest.cells(destRow, 32 + i).value))
            Case "over": reason = reason & macdLabels(i) & " Over; ": bullishScore = bullishScore + 1
            Case "under": reason = reason & macdLabels(i) & " Under; ": bearishScore = bearishScore + 1
        End Select
    Next i
    
    ' EMA signals 10 Vs 30 (Cols 40–42)
    For i = 0 To 2
        Select Case LCase(Trim(wsDest.cells(destRow, 36 + i).value))
            Case "over": reason = reason & emaLabels(i) & " 10_30 Over; ": bullishScore = bullishScore + 1
            Case "under": reason = reason & emaLabels(i) & " 10_30 Under; ": bearishScore = bearishScore + 1
        End Select
    Next i
    
    ' EMA signals 10 Vs 50 (Cols 40–42)
    For i = 0 To 2
        Select Case LCase(Trim(wsDest.cells(destRow, 40 + i).value))
            Case "over": reason = reason & emaLabels(i) & " 10_50 Over; ": bullishScore = bullishScore + 1
            Case "under": reason = reason & emaLabels(i) & " 10_50 Under; ": bearishScore = bearishScore + 1
        End Select
    Next i

    ' RSI: Cols 18–21
    Dim rsi(1 To 4) As Double
    For i = 1 To 4
        'rsi(i) = CDbl(IIf(IsEmpty(wsDest.Cells(destRow, 17 + i).Value), 0, wsDest.Cells(destRow, 17 + i).Value))
        rsi(i) = wsDest.cells(destRow, 17 + i).value
    Next i

    

    
    If rsi(1) < 25 Then reason = reason & "5m RSI oversold; ": bullishScore = bullishScore + 1
    If rsi(2) < 25 Then reason = reason & "15m RSI oversold; ": bullishScore = bullishScore + 1
    If rsi(3) < 25 Then reason = reason & "1H RSI oversold; ": bullishScore = bullishScore + 1
    If rsi(4) < 25 Then reason = reason & "1D RSI oversold; ": bullishScore = bullishScore + 1

    If rsi(1) > 75 Then reason = reason & "5m RSI overbought; ": bearishScore = bearishScore + 1
    If rsi(2) > 75 Then reason = reason & "15m RSI overbought; ": bearishScore = bearishScore + 1
    If rsi(3) > 75 Then reason = reason & "1H RSI overbought; ": bearishScore = bearishScore + 1
    If rsi(4) > 75 Then reason = reason & "1D RSI overbought; ": bearishScore = bearishScore + 1

    ' Supertrend (Col 22)
    Dim trend As String: trend = Trim(wsDest.cells(destRow, 22).value)
    If trend <> "" Then
        reason = reason & "Supertrend: " & trend & "; "
        If LCase(trend) = "u" Then bullishScore = bullishScore + 1
        If LCase(trend) = "d" Then bearishScore = bearishScore + 1
    End If

    
    ' Volume Spike (Col 71)
    Dim spike As Variant: spike = wsSource.cells(srcRow, 71).value
    Dim vwap As Double: vwap = wsSource.cells(srcRow, "BR").value
    Dim Price As Double: Price = wsSource.cells(srcRow, 20).value
    
    If Not IsEmpty(spike) And spike <> 0 Then
        reason = reason & "Vol Spike; "
        If Price > vwap Then
            bullishScore = bullishScore + 1
            reason = reason & "Price > VWAP; "
        ElseIf CurrentPrice < vwap Then
            bearishScore = bearishScore + 1
            reason = reason & "Price < VWAP; "
        End If
    End If

    ' Determine Action
    If action = "" Then
        If bullishScore > bearishScore Then
            action = "Buy"
        ElseIf bearishScore > bullishScore Then
            action = "Sell"
        Else
            flag = "MomentumOnly"
        End If
        If action <> "" Then wsDest.cells(destRow, 2).value = action
    End If

    ' Write reason and flag
    wsDest.cells(destRow, 30).value = Trim(reason)
    wsDest.cells(destRow, 31).value = flag
    If action <> "Buy" And action <> "Sell" Then Exit Sub

    ' Pricing & Levels
    
    Dim supertrendVal As Double: supertrendVal = wsSource.cells(srcRow, 72).value
    Dim ATR_5m As Double: ATR_5m = wsSource.cells(srcRow, 81).value
    Dim bbU_5m As Double: bbU_5m = wsSource.cells(srcRow, 87).value
    Dim bbL_5m As Double: bbL_5m = wsSource.cells(srcRow, 89).value

    Dim highVal As Double: highVal = wsSource.cells(srcRow, 74).value
    Dim lowVal As Double: lowVal = wsSource.cells(srcRow, 75).value
    Dim closeVal As Double: closeVal = wsSource.cells(srcRow, 76).value
    Dim pp As Double: pp = (highVal + lowVal + closeVal) / 3

    Dim s3 As Double, s2 As Double, s1 As Double, r1 As Double, r2 As Double, r3 As Double
        s3 = wsDest.cells(destRow, 23).value
        s2 = wsDest.cells(destRow, 24).value
        s1 = wsDest.cells(destRow, 25).value
        pp = wsDest.cells(destRow, 26).value
        r1 = wsDest.cells(destRow, 27).value
        r2 = wsDest.cells(destRow, 28).value
        r3 = wsDest.cells(destRow, 29).value
    
    
    Dim entryPrice As Double, stopLossPrice As Double, targetPrice As Double
     
    Select Case action
    Case "Buy"
        
        If Price < wsSource.cells(srcRow, 45).value Then
            entryPrice = Price
        Else
            entryPrice = wsSource.cells(srcRow, 45).value ' picks up the 10EMA on a 5 min from Sheet1
        End If
    
    
        If Price >= s3 And Price < s2 Then
            targetPrice = s2
            stopLossPrice = s3
        ElseIf Price >= s2 And Price < s1 Then
            targetPrice = s1
            stopLossPrice = s2
        ElseIf Price >= s1 And Price < pp Then
            targetPrice = pp
            stopLossPrice = s1
        ElseIf Price >= pp And Price < r1 Then
            targetPrice = r1
            stopLossPrice = pp
        ElseIf Price >= r1 And Price < r2 Then
            targetPrice = r2
            stopLossPrice = r1
        ElseIf Price >= r2 And Price <= r3 Then
            targetPrice = r3
            stopLossPrice = r2
        End If
        
    Case "Sell"
        If Price > wsSource.cells(srcRow, 45).value Then
            entryPrice = Price
        Else
            entryPrice = wsSource.cells(srcRow, 45).value ' picks up the 10EMA on a 5 min from Sheet1
        End If
        If Price <= r3 And Price > r2 Then
            targetPrice = r1
            stopLossPrice = r3
        ElseIf Price <= r2 And Price > r1 Then
            targetPrice = pp
            stopLossPrice = r2
        ElseIf Price <= r1 And Price > pp Then
            targetPrice = s1
            stopLossPrice = r1
        ElseIf Price <= pp And Price > s1 Then
            targetPrice = s2
            stopLossPrice = pp
        ElseIf Price <= s1 And Price > s2 Then
            targetPrice = s3
            stopLossPrice = s1
        ElseIf Price <= s2 And Price >= s3 Then
            targetPrice = s3
            stopLossPrice = s2
        End If
    End Select
    
    
  

    wsDest.cells(destRow, 3).value = entryPrice
    wsDest.cells(destRow, 4).value = targetPrice
    wsDest.cells(destRow, 5).value = stopLossPrice
    
    
    ' Adding more details to the ticker to get those minor variations!!
    '==================================================================
    
    
    
    ' Stochastics comment
    
    Dim stochK5, stochD5, stochK15, stochD15 As Double
    Dim stochComment As String
    
    
    
        stoch5K = wsSource.cells(srcRow, "CE").value
        stoch5D = wsSource.cells(srcRow, "CF").value
        stoch15K = wsSource.cells(srcRow, "CG").value
        stoch15D = wsSource.cells(srcRow, "CH").value
    
    
    If stoch5K > stoch5D Then
        stochComment = "  STOC: 5m Bullish"
    ElseIf stoch5K < stoch5D Then
        stochComment = "  STOC: 5m Bearish"
    Else
        stochComment = "  STOC: 5m Blanks"
    End If
    
    If stoch15K > stoch15D Then
        stochComment = stochComment & "  STOC: 15m Bullish"
    ElseIf stoch15K < stoch15D Then
        stochComment = stochComment & "  STOC: 15m Bearish"
    Else
        stochComment = stochComment & "  STOC: 15m Blanks"
    End If
    
    
    'MACD Posturing
    
    Dim FL_5m As Double, SL_5m As Double, SIG_5m As Double
    Dim FL_15m As Double, SL_15m As Double, SIG_15m As Double
    Dim MACDLine_5m As Double, MACDLine_15m As Double
    Dim macdNarrative As String
    
 
        FL_5m = wsSource.cells(srcRow, "CY").value
        SL_5m = wsSource.cells(srcRow, "CZ").value
        SIG_5m = wsSource.cells(srcRow, "DA").value
    
        FL_15m = wsSource.cells(srcRow, "DB").value
        SL_15m = wsSource.cells(srcRow, "DC").value
        SIG_15m = wsSource.cells(srcRow, "DD").value
    
    
    MACDLine_5m = FL_5m - SL_5m
    MACDLine_15m = FL_15m - SL_15m
    
    macdNarrative = ""
    
    ' --- 5m posture ---
    If FL_5m > SL_5m Then
        'macdNarrative = "      MACD(5): Bullish: FL > SL"
        macdNarrative = "      5m: FL>SL"
    ElseIf FL_5m < SL_5m Then
        'macdNarrative = "      MACD(5): Bearish: FL < SL"
        macdNarrative = "      5m: FL<SL"
    Else
        macdNarrative = "      5m: FL=SL"
    End If
    
    ' --- 5m crossover ---
    If MACDLine_5m > SIG_5m Then
        macdNarrative = macdNarrative & " || M>S"
    ElseIf MACDLine_5m < SIG_5m Then
        macdNarrative = macdNarrative & " || M<S"
    Else
        macdNarrative = macdNarrative & " || M=S"
    End If
    
    ' --- 15m posture ---
    If FL_15m > SL_15m Then
        'If macdNarrative = "" Then
            'macdNarrative = "      MACD: Bullish(15m) : FL > SL"
        'Else
            macdNarrative = macdNarrative & " || 15m: FL>SL"
        'End If
    ElseIf FL_15m < SL_15m Then
        'If macdNarrative = "" Then
        '    macdNarrative = "      MACD: Bearish(15m) : FL < SL"
        'Else
            macdNarrative = macdNarrative & " || 15m: FL<SL"
        'End If
    Else
        macdNarrative = macdNarrative & " || 15m: FL=SL"
    End If
    
    ' --- 15m crossover ---
    If MACDLine_15m > SIG_15m Then
        macdNarrative = macdNarrative & " || 15m: M>S"
    ElseIf MACDLine_15m < SIG_15m Then
        macdNarrative = macdNarrative & " || 15m: M<S"
    Else
        macdNarrative = macdNarrative & " || 15m: M=S"
    End If
    
    wsDest.cells(destRow, 44).value = Trim(macdNarrative & stochComment)
    '========================================================================

End Sub

Sub CopyValidTradesToNewSheet(wsMainSource As Worksheet, srcMRow As Long)

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim srcData As Variant, destData() As Variant
    Dim lastRow As Long, destRow As Long
    Dim i As Long
    Dim sReason As String
    Dim pp As Double
    Dim pos As Long, txt As String
    Dim r As Long

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("Trades")
    Set wsDest = ThisWorkbook.Sheets("New_Trades")

    ' Find last row in source
    lastRow = wsSource.cells(wsSource.Rows.count, "A").End(xlUp).row

    ' Load source data into array (Rows 3 to lastRow, Columns A to AD)
    'srcData = wsSource.Range("A3:AD" & lastRow).Value
    srcData = wsSource.Range("A3:AX" & lastRow).value

    ' Clear old content in destination sheet from A2 to N
    Dim lastDestRow As Long
    lastDestRow = wsDest.cells(wsDest.Rows.count, "A").End(xlUp).row
    
    'wsDest.Range("A2:N" & lastDestRow).Font.ColorIndex = 0
    
    wsDest.Range("A2:O" & lastDestRow).ClearContents
    

    ' Optional: Write headers
    wsDest.Range("A1:N1").value = Array("Ticker", "Action", "Entry", "Exit", "SL", "S3", "S2", "S1", "PP", "R1", "R2", "R3", "Reason", "Time")

    ' Resize destination array
    ReDim destData(1 To UBound(srcData), 1 To 15)
    destRow = 1

    ' Loop through each row in source data
    For i = 1 To UBound(srcData)
        If Trim(srcData(i, 2)) <> "" Then ' Column B (Action)
            destData(destRow, 1) = srcData(i, 1)   ' A - Ticker
            destData(destRow, 2) = srcData(i, 2)   ' B - Action
            destData(destRow, 3) = srcData(i, 3)   ' C - Entry
            destData(destRow, 4) = srcData(i, 4)   ' D - Exit
            destData(destRow, 5) = srcData(i, 5)   ' E - SL

          
            
            destData(destRow, 6) = srcData(i, 23)  ' S3
            destData(destRow, 7) = srcData(i, 24)  ' S2
            destData(destRow, 8) = srcData(i, 25)  ' S1
            destData(destRow, 9) = srcData(i, 26)  ' PP
            destData(destRow, 10) = srcData(i, 27) ' R1
            destData(destRow, 11) = srcData(i, 28) ' R2
            destData(destRow, 12) = srcData(i, 29) ' R3

            destData(destRow, 13) = srcData(i, 30)  ' Reason
            destData(destRow, 14) = Format(Now, "hh:mm") ' Timestamp
            destData(destRow, 15) = srcData(i, 44)
            destRow = destRow + 1
            
            'HighlightKeywords destData(destRow, 13), True
            'HighlightKeywords destData(destRow, 15), True ' skip EMA/MACD/Supertrend

        End If
    Next i

    ' Write filtered rows to destination sheet
    If destRow > 1 Then
        wsDest.Range("A2").Resize(destRow - 1, 15).value = destData
    End If

    ' Format Reason (Column M = 13)
    'For r = 2 To destRow - 1
    ' Format Reason column
    
'    For r = 2 To destRow + 1
'        Select Case LCase(LTrim(wsDest.Range("B" & r).value))
'            Case "buy"
'                wsDest.Range("A" & r).Font.color = RGB(0, 150, 0)
'            Case "sell"
'                wsDest.Range("A" & r).Font.color = RGB(200, 0, 0)
'        End Select
'    Next r
'
'    ' Format columns M and O using reusable logic
'    For r = 2 To destRow + 1
'        HighlightKeywords wsDest.Range("M" & r)
'        HighlightKeywords wsDest.Range("O" & r), True ' skip EMA/MACD/Supertrend
'    Next r
    
End Sub
'Private Sub HighlightKeywords(cell As Range, Optional skipWords As Boolean = False)
'    Dim keywords As Variant
'    keywords = Array( _
'        Array("EMA", RGB(255, 0, 0)), Array("MACD", RGB(0, 0, 255)), Array("Supertrend", RGB(200, 125, 0)), _
'        Array("5m", RGB(0, 0, 255)), Array("15m", RGB(0, 0, 255)), Array("1H", RGB(0, 0, 255)), _
'        Array("STOC", RGB(0, 0, 255)), Array("FL<SL", RGB(200, 0, 0)), Array("M<S", RGB(200, 0, 0)), Array("FL>SL", RGB(0, 150, 0)), _
'        Array("M>S", RGB(0, 150, 0)), Array("Bullish", RGB(0, 150, 0)), Array("Bearish", RGB(200, 0, 0)), Array("Under", RGB(200, 0, 0)), _
'        Array("BUY", RGB(0, 150, 0)), Array("S-BUY", RGB(0, 150, 0)), Array("SELL", RGB(180, 0, 0)), Array("S-SELL", RGB(180, 0, 0)), _
'        Array("RSI", RGB(0, 0, 255)), Array("oversold", RGB(0, 150, 0)), Array("overbought", RGB(180, 0, 0)), Array("Vol Spike", RGB(0, 150, 0)), _
'        Array("Price < VWAP", RGB(180, 0, 0)), Array("Price > VWAP", RGB(180, 0, 0)) _
'    )
'
'    Dim re As Object, matches As Object, match As Object
'    Dim i As Long, txt As String, pattern As String, color As Long
'    Dim startPos As Long
'
'    txt = cell.value
'    cell.Font.ColorIndex = 0
'    cell.Font.Bold = False
'
'    Set re = CreateObject("VBScript.RegExp")
'    re.Global = True
'    re.IgnoreCase = True
'
'    For i = LBound(keywords) To UBound(keywords)
'        Dim word As String: word = keywords(i)(0)
'
'        ' skip EMA/MACD/Supertrend in column O
'        If skipWords And (word = "EMA" Or word = "MACD" Or word = "Supertrend") Then GoTo SkipThis
'
'        ' Match full word only using \b boundary
'        pattern = "\b" & word & "\b"
'        re.pattern = pattern
'
'        If re.Test(txt) Then
'            Set matches = re.Execute(txt)
'            For Each match In matches
'                startPos = match.FirstIndex + 1 ' VBA is 1-based
'                With cell.Characters(startPos, Len(word)).Font
'                    .color = keywords(i)(1)
'                    .Bold = True
'                End With
'            Next match
'        End If
'SkipThis:
'    Next i
'End Sub

Private Sub HighlightKeywordsFast(rng As Range, Optional skipWords As Boolean = False)
    Dim keywords As Variant
    keywords = Array( _
        Array("EMA", RGB(255, 0, 0)), Array("MACD", RGB(0, 0, 255)), Array("Supertrend", RGB(200, 125, 0)), _
        Array("5m", RGB(0, 0, 255)), Array("15m", RGB(0, 0, 255)), Array("1H", RGB(0, 0, 255)), _
        Array("STOC", RGB(0, 0, 255)), Array("FL<SL", RGB(200, 0, 0)), Array("M<S", RGB(200, 0, 0)), Array("FL>SL", RGB(0, 150, 0)), _
        Array("M>S", RGB(0, 150, 0)), Array("Bullish", RGB(0, 150, 0)), Array("Bearish", RGB(200, 0, 0)), Array("Under", RGB(200, 0, 0)), _
        Array("BUY", RGB(0, 150, 0)), Array("S-BUY", RGB(0, 150, 0)), Array("SELL", RGB(180, 0, 0)), Array("S-SELL", RGB(180, 0, 0)), _
        Array("RSI", RGB(0, 0, 255)), Array("oversold", RGB(0, 150, 0)), Array("overbought", RGB(180, 0, 0)), Array("Vol Spike", RGB(0, 150, 0)), _
        Array("Price < VWAP", RGB(180, 0, 0)), Array("Price > VWAP", RGB(180, 0, 0)) _
    )

    Dim cell As Range, i As Long, pos As Long, word As String, txt As String

    For Each cell In rng.cells
        txt = CStr(cell.value)
        If Len(txt) = 0 Then GoTo NextCell

        ' reset formatting
        cell.Font.ColorIndex = 0
        cell.Font.Bold = False

        For i = LBound(keywords) To UBound(keywords)
            word = keywords(i)(0)

            ' skip EMA/MACD/Supertrend in column O if asked
            'If skipWords And (word = "EMA" Or word = "MACD" Or word = "Supertrend") Then GoTo SkipThis

            pos = InStr(1, txt, word, vbTextCompare)
            Do While pos > 0
                cell.Characters(pos, Len(word)).Font.Color = keywords(i)(1)
                cell.Characters(pos, Len(word)).Font.Bold = True
                pos = InStr(pos + Len(word), txt, word, vbTextCompare)
            Loop
SkipThis:
        Next i
NextCell:
    Next cell
End Sub




' ============================================
' Smart Position Update with Course Correction Logic (Enhanced Version)
' ============================================
Function GetTrendText(cell As Range, value As Variant) As String
    Select Case cell.Interior.Color
        Case RGB(255, 0, 0): GetTrendText = "Bearish (" & value & ")"
        Case RGB(0, 0, 255): GetTrendText = "Bullish (" & value & ")"
        Case Else: GetTrendText = "Flat (" & value & ")"
    End Select
End Function
' --------------------------------------------------------------------------
' UpdateSmartNextSteps - improved decision engine & concise next-steps
' - Preserves original indicator reads and cell layout.
' - Adds a weighted scoring engine, concise Next Steps, and verbose detail.
' - Color-codes outputs and retains your keyword-highlighting behavior.
' --------------------------------------------------------------------------
Sub UpdateSmartNextSteps()
    Dim wsLive As Worksheet, wsMain As Worksheet
    Dim i As Long, lastRow As Long
    Dim ticker As String, action As String, status As String
    Dim matchRow As Variant, nowStamp As String
    Dim nextStepMsg As String, fullLine As String
    Dim entryPrice As Double, slPrice As Double, CurrentPrice As Double
    Dim RSI_5m As Double, rsi_15m As Double, vwap As Double, VolumeSpike As Variant
    Dim MACD_5m As String, macd_15m As String, EMA_5m As String, EMA_15m As String
    Dim superTrend As Double
    Dim BB_U As Double, BB_L As Double, ATR_5m As Double
    Dim trailSL As Double, tightenSL As Double
    Dim s1 As Double, s2 As Double, s3 As Double, r1 As Double, r2 As Double, r3 As Double, pp As Double
    Dim highVal As Variant, lowVal As Variant, closeVal As Variant
    Dim gainPct As Double
    Dim cellHistory As Range, celllatest As Range
    Dim trendOutlook As String, slAssessment As String, reversalWatch As String, rsiDivergence As String, SandR As String
    Dim candleBull As String, candleBear As String
    Dim reversalConfidence As Integer
    Dim PriceDiff As Double, tolerance As Double
    Dim quantity As Double
    Dim stochK As Double, stochD As Double, stochComment As String, bollComment As String
    Dim FL_5m As Double, SL_5m As Double, SIG_5m As Double
    Dim FL_15m As Double, SL_15m As Double, SIG_15m As Double
    Dim MACDLine_5m As Double, MACDLine_15m As Double
    Dim CurrentVolume As Double
    Dim confidenceScore As Integer
    Dim lastDataRow As Long

    Set wsLive = ThisWorkbook.Sheets("Live")
    Set wsMain = ThisWorkbook.Sheets("Sheet1")
    lastRow = wsLive.cells(wsLive.Rows.count, "A").End(xlUp).row

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        momentumList = ""
        nowStamp = "  DATE: " & Format(Now, "dd-mmm hh:mm")
        status = Trim(wsLive.cells(i, "I").value)
        'If status <> "O" Then GoTo SkipRow

        ticker = Trim(wsLive.cells(i, "A").value)
        action = LCase(Trim(wsLive.cells(i, "B").value))
        quantity = val(wsLive.cells(i, "C").value)
        entryPrice = val(wsLive.cells(i, "D").value)
        slPrice = val(wsLive.cells(i, "F").value)

        Set celllatest = wsLive.cells(i, "H")

        matchRow = Application.Match(ticker, wsMain.Range("A:A"), 0)
        If IsError(matchRow) Then GoTo SkipRow

        With wsMain
            RSI_5m = val(.cells(matchRow, "U").value)
            rsi_15m = val(.cells(matchRow, "V").value)
            VolumeSpike = .cells(matchRow, "BS").value
            vwap = val(.cells(matchRow, "BR").value)
            superTrend = val(.cells(matchRow, "BT").value)
            BB_U = val(.cells(matchRow, "CI").value)
            BB_L = val(.cells(matchRow, "CK").value)
            ATR_5m = val(.cells(matchRow, "CC").value)
            CurrentPrice = val(.cells(matchRow, "T").value)
            CurrentVolume = val(.cells(matchRow, "H").value) ' assume column H holds current volume for that ticker
            highVal = .cells(matchRow, "BV").value
            lowVal = .cells(matchRow, "BW").value
            closeVal = .cells(matchRow, "BX").value
            candleBull = .cells(matchRow, "AK").value & .cells(matchRow, "AL").value
            candleBear = .cells(matchRow, "AO").value & .cells(matchRow, "AP").value

            stochK = val(.cells(matchRow, "CE").value)
            stochD = val(.cells(matchRow, "CF").value)

            FL_5m = val(.cells(matchRow, "CY").value)
            SL_5m = val(.cells(matchRow, "CZ").value)
            SIG_5m = val(.cells(matchRow, "DA").value)
            FL_15m = val(.cells(matchRow, "DB").value)
            SL_15m = val(.cells(matchRow, "DC").value)
            SIG_15m = val(.cells(matchRow, "DD").value)
        End With

        If stochK > stochD Then
            stochComment = "  STOCHASTIC: Bullish crossover on 5m – Confirmation of Upward Momentum"
        ElseIf stochK < stochD Then
            stochComment = "  STOCHASTIC: Bearish crossover on 5m – Possible Downside Pressure"
        Else
            stochComment = "  STOCHASTIC: No crossover – Stochastics are in sync or flat"
        End If

        If CurrentPrice > BB_U Then
            bollComment = "  BOLLINGER: Price above upper Bollinger Band – Overbought/Breakout scenario"
        ElseIf CurrentPrice < BB_L Then
            bollComment = "  BOLLINGER: Price dropped below lower Bollinger Band – Potential Mean Reversion/Oversold"
        Else
            bollComment = "  BOLLINGER: Price within bands – no edge from Bollinger signal"
        End If

        s3 = Round(wsMain.cells(matchRow, "K").value, 2)
        s2 = Round(wsMain.cells(matchRow, "L").value, 2)
        s1 = Round(wsMain.cells(matchRow, "M").value, 2)
        pp = Round(wsMain.cells(matchRow, "N").value, 2)
        r1 = Round(wsMain.cells(matchRow, "O").value, 2)
        r2 = Round(wsMain.cells(matchRow, "P").value, 2)
        r3 = Round(wsMain.cells(matchRow, "Q").value, 2)

        SandR = "   (S3:  " & s3 & ") (S2:  " & s2 & ") (S1:  " & s1 & ") (PP:  " & pp & ") (R1:  " & r1 & ") (R2:  " & r2 & ") (R3:  " & r3 & ")"

        PriceDiff = Round(CurrentPrice - entryPrice, 2)
        nowStamp = nowStamp & "  ||  CURRENT PRICE: " & CStr(CurrentPrice) & " "
        wsLive.cells(i, "G").value = Abs(PriceDiff * quantity)
        wsLive.cells(i, "G").Font.Bold = True

        Select Case Trim(LCase(action))
        Case "buy"
            If entryPrice >= s3 And entryPrice < s2 Then
                wsLive.cells(i, "E").value = s2
                wsLive.cells(i, "F").value = s3
                nowStamp = nowStamp & "(Between S3 and S2)"
            ElseIf entryPrice >= s2 And entryPrice < s1 Then
                wsLive.cells(i, "E").value = s1
                wsLive.cells(i, "F").value = s2
                nowStamp = nowStamp & "(Between S2 and S1)"
            ElseIf entryPrice >= s1 And entryPrice < pp Then
                wsLive.cells(i, "E").value = pp
                wsLive.cells(i, "F").value = s1
                nowStamp = nowStamp & "(Between S1 and PP)"
            ElseIf entryPrice >= pp And entryPrice < r1 Then
                wsLive.cells(i, "E").value = r1
                wsLive.cells(i, "F").value = pp
                nowStamp = nowStamp & "(Between PP and R1)"
            ElseIf entryPrice >= r1 And entryPrice < r2 Then
                wsLive.cells(i, "E").value = r2
                wsLive.cells(i, "F").value = r1
                nowStamp = nowStamp & "(Between R1 and R2)"
            ElseIf entryPrice >= r2 And entryPrice <= r3 Then
                wsLive.cells(i, "E").value = r3
                wsLive.cells(i, "F").value = r2
                nowStamp = nowStamp & "(Between R2 and R3)"
            End If
        Case "sell"
            If entryPrice <= r3 And entryPrice > r2 Then
                wsLive.cells(i, "E").value = r1
                wsLive.cells(i, "F").value = r3
                nowStamp = nowStamp & "(Between R2 and R3)"
            ElseIf entryPrice <= r2 And entryPrice > r1 Then
                wsLive.cells(i, "E").value = pp
                wsLive.cells(i, "F").value = r2
                nowStamp = nowStamp & "(Between R1 and R2)"
            ElseIf entryPrice <= r1 And entryPrice > pp Then
                wsLive.cells(i, "E").value = s1
                wsLive.cells(i, "F").value = r1
                nowStamp = nowStamp & "(Between PP and R1)"
            ElseIf entryPrice <= pp And entryPrice > s1 Then
                wsLive.cells(i, "E").value = s2
                wsLive.cells(i, "F").value = pp
                nowStamp = nowStamp & "(Between S1 and PP)"
            ElseIf entryPrice <= s1 And entryPrice > s2 Then
                wsLive.cells(i, "E").value = s3
                wsLive.cells(i, "F").value = s1
                nowStamp = nowStamp & "(Between S2 and S1)"
            ElseIf entryPrice <= s2 And entryPrice >= s3 Then
                wsLive.cells(i, "E").value = s3
                wsLive.cells(i, "F").value = s2
                nowStamp = nowStamp & "(Between S3 and S2)"
            End If
        End Select

        Select Case Trim(LCase(action))
            Case "sell"
                wsLive.cells(i, "G").Font.Color = IIf(PriceDiff >= 0, RGB(255, 0, 0), RGB(0, 180, 0))
            Case "buy"
                wsLive.cells(i, "G").Font.Color = IIf(PriceDiff >= 0, RGB(0, 180, 0), RGB(255, 0, 0))
        End Select

        Dim CandleStickPattern As String
        If candleBear = "" Or candleBull = "" Then
            CandleStickPattern = " No discernable candle patterns deteted"
        Else
            If (action = "buy" And candleBear <> "") Then CandleStickPattern = "  Bearish Reversal candle : " & candleBear
            If (action = "sell" And candleBull <> "") Then CandleStickPattern = "  Bullish Reversal candle : " & candleBull
        End If

        Dim strSupertrend As String
        tolerance = 0.001 * CurrentPrice
        If (action = "buy" And CurrentPrice < superTrend - tolerance) Or _
           (action = "sell" And CurrentPrice > superTrend + tolerance) Then
            strSupertrend = "    SUPERTREND (" & CStr(superTrend) & ") - Opposite to position; "
            reversalConfidence = reversalConfidence + 2
        Else
            strSupertrend = "  SUPERTREND: " & CStr(superTrend)
        End If

        Dim strVol As String
        If VolumeSpike = 1 Then
            Select Case LCase(action)
                Case "buy":
                    If (CurrentPrice < vwap) Then
                        strVol = "  Buy setup losing support – Volume Spike but price below VWAP"
                        reversalConfidence = reversalConfidence + 2
                    Else
                        strVol = "  HOLD BUY - Volume Spike detected and price above VWAP"
                    End If
                Case "sell":
                    If (CurrentPrice > vwap) Then
                        strVol = "  Sell setup under pressure – Volume Spike but price above VWAP"
                        reversalConfidence = reversalConfidence + 2
                    Else
                        strVol = "  HOLD SELL - Volume Spike detected and price below VWAP"
                    End If
            End Select
        Else
            strVol = "  No Volume Spike observed"
        End If

        If RSI_5m > 70 And rsi_15m < 65 Then
            rsiDivergence = "RSI divergence – overbought 5m, weakening 15m"
        ElseIf RSI_5m < 30 And rsi_15m > 35 Then
            rsiDivergence = "  RSI Divergence – oversold 5m, recovering 15m"
        Else
            rsiDivergence = "  RSI Divergence - None Observed"
        End If

        If InStr(reversalWatch, "Reversal") > 0 Then reversalConfidence = reversalConfidence + 3
        If rsiDivergence <> "" Then reversalConfidence = reversalConfidence + 3
        If reversalWatch <> "" Then
            reversalWatch = "Reversal Risk: " & reversalWatch & rsiDivergence
            If reversalConfidence >= 8 Then reversalWatch = reversalWatch & " >> Candle Watch across 5m–15m suggests potential reversal; wait for confirmation."
        End If

        If wsMain.cells(matchRow, "AS").value >= wsMain.cells(matchRow, "AT").value Then
            trendOutlook = "  EMA(10_30) : 5m - Bullish "
        Else
            trendOutlook = "  EMA(10_30) : 5m - Bearish "
        End If

        If wsMain.cells(matchRow, "AX").value >= wsMain.cells(matchRow, "AY").value Then
            trendOutlook = trendOutlook & " || 15m - Bullish  "
        Else
            trendOutlook = trendOutlook & " || 15m - Bearish  "
        End If

        If wsMain.cells(matchRow, "AS").value >= wsMain.cells(matchRow, "AU").value Then
            trendOutlook = trendOutlook & "  EMA(10_50) : 5m - Bullish  "
        Else
            trendOutlook = trendOutlook & "  EMA(10_50) : 5m - Bearish  "
        End If

        If wsMain.cells(matchRow, "AX").value >= wsMain.cells(matchRow, "AZ").value Then
            trendOutlook = trendOutlook & " || 15m - Bullish "
        Else
            trendOutlook = trendOutlook & " || 15m - Bearish"
        End If

        MACDLine_5m = FL_5m - SL_5m
        MACDLine_15m = FL_15m - SL_15m
        macdNarrative = ""

        If FL_5m >= SL_5m Then
            macdNarrative = "  MACD(F_S): 5m - Bullish  "
        Else
            macdNarrative = "  MACD(F_S): 5m - Bearish  "
        End If

        If FL_15m >= SL_15m Then
            macdNarrative = macdNarrative & " ||  15m - Bullish  "
        Else
            macdNarrative = macdNarrative & " ||  15m - Bearish  "
        End If

        If MACDLine_5m >= SIG_5m Then
            macdNarrative = macdNarrative & "  MACD(M_S): 5m - Bullish  "
        Else
            macdNarrative = macdNarrative & "  MACD(M_S): 5m - Bearish  "
        End If

        If MACDLine_15m >= SIG_15m Then
            macdNarrative = macdNarrative & " || 15m - Bullish"
        Else
            macdNarrative = macdNarrative & " || 15m - Bearish"
        End If

        Dim SLgap As Double
        SLgap = Abs(CurrentPrice - slPrice)
        If SLgap <= ATR_5m Then
            slAssessment = "[SL Warning] SL may be too tight"
        ElseIf (action = "buy" And slPrice < s3) Or (action = "sell" And slPrice > r3) Then
            slAssessment = "[Wide SL] SL beyond extreme levels"
        Else
            slAssessment = "SL in safe zone"
        End If

        Dim pctGap As Double
        pctGap = Abs(CurrentPrice - slPrice) / CurrentPrice

        If action = "sell" Then
            If (CurrentPrice - slPrice) <= 0 Then
                If pctGap <= 0.1 Then
                    slComment = "  SL is within 10% of current price"
                Else
                    slComment = "  SL is " & CStr(pctGap * 100) & "% away from current price"
                End If
            Else
                slComment = "  SL Breached"
            End If
        Else
            If (CurrentPrice - slPrice) >= 0 Then
                If pctGap <= 0.1 Then
                    slComment = "  SL is within 10% of current price"
                Else
                    slComment = "  SL is " & CStr(pctGap * 100) & "% away from current price"
                End If
            Else
                slComment = "  SL Breached"
            End If
        End If

        Dim strRSI As String
        strRSI = "   RSI: (5m: " & Format(RSI_5m, "0") & ") (15m: " & Format(rsi_15m, "0") & _
               ") (1H: " & Format(val(wsMain.cells(matchRow, "W").value)) & _
               ") (1D: " & Format(val(wsMain.cells(matchRow, "X").value) & ")")

        fullLine = nowStamp & vbCrLf & _
                    strRSI & vbCrLf & _
                    SandR & vbCrLf & _
                    "------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
                    trendOutlook & vbCrLf & _
                    macdNarrative & vbCrLf & _
                    strSupertrend & vbCrLf & _
                    "------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
                    stochComment & vbCrLf & _
                    bollComment & vbCrLf & _
                    "------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
                    CandleStickPattern & vbCrLf & _
                    strVol & vbCrLf & _
                    slComment & vbCrLf & _
                    rsiDivergence

        
        With celllatest
            .ClearContents
            .value = RTrim(fullLine)
        End With

        Dim COLOR_RED As Long: COLOR_RED = RGB(200, 0, 0)
        Dim COLOR_GREEN As Long: COLOR_GREEN = RGB(0, 160, 0)
        Dim COLOR_BLUE As Long: COLOR_BLUE = RGB(0, 0, 255)
        Dim COLOR_YELLOW As Long: COLOR_YELLOW = RGB(200, 140, 0)

        Dim keywordMap As Variant

        keywordMap = Array( _
            Array("MOMENTUM", COLOR_YELLOW), Array("PIVOTS", COLOR_YELLOW), Array("INDICATORS", COLOR_YELLOW), Array("RISK CHECKS", COLOR_YELLOW), _
            Array("EMA(10_30)", COLOR_BLUE), Array("EMA(10_50)", COLOR_BLUE), Array("MACD(F_S)", COLOR_BLUE), Array("MACD(M_S)", COLOR_BLUE), Array("SUPERTREND", COLOR_BLUE), _
            Array("STOCHASTIC", COLOR_BLUE), Array("BOLLINGER", COLOR_BLUE), Array("RSI", COLOR_BLUE), Array("5m", COLOR_BLUE), Array("15m", COLOR_BLUE), Array("1H", COLOR_BLUE), Array("1D", COLOR_BLUE), _
            Array("S3", COLOR_RED), Array("S2", COLOR_RED), Array("S1", COLOR_RED), Array("PP", COLOR_BLUE), Array("R1", COLOR_GREEN), Array("R2", COLOR_GREEN), Array("R3", COLOR_GREEN), _
            Array("Bearish", COLOR_RED), Array("Bullish", COLOR_GREEN), Array("Opposite to position", COLOR_RED), _
            Array("Volume Spike", COLOR_BLUE), Array("VWAP", COLOR_BLUE), Array("SL", COLOR_RED), Array("HOLD", COLOR_BLUE), _
            Array("Reversal candle", COLOR_BLUE), Array("Overbought/Breakout", COLOR_BLUE), Array("Potential Mean Reversion/Oversold", COLOR_BLUE), _
            Array("Confirmation of Upward Momentum", COLOR_BLUE), Array("Possible Downside Pressure", COLOR_BLUE), _
            Array("DATE", COLOR_BLUE), Array("CURRENT PRICE", COLOR_BLUE))

        Dim p As Long, pos As Long
        Dim word As String, colorVal As Long

        For p = 0 To UBound(keywordMap)
            word = keywordMap(p)(0)
            colorVal = keywordMap(p)(1)
            pos = 1
            Do
                pos = InStr(pos, fullLine, word, vbBinaryCompare)
                If pos = 0 Then Exit Do
                On Error Resume Next
                celllatest.Characters(pos, Len(word)).Font.Color = colorVal
                celllatest.Characters(pos, Len(word)).Font.Bold = True
                On Error GoTo 0
                pos = pos + Len(word)
            Loop
        Next p
        
        'celllatest.EntireRow.RowHeight = 300
        celllatest.WrapText = True
        celllatest.VerticalAlignment = xlTop
        'celllatest.Characters.Font.Size = 10
        'Debug.Print celllatest.EntireRow.RowHeight
        'Debug.Print ActiveSheet.ProtectContents
        

       ' === COLUMN J NEXT STEPS ===
        If wsLive.cells(i, "G") <> "" Then
            Dim nextStepsText As String
            Dim tradeDirection As String: tradeDirection = LCase(Trim(action))
            Dim isBuy As Boolean: isBuy = (tradeDirection = "buy")
            Set cellNextSteps = wsLive.cells(i, "J")

            '--- Trend Support and Oppose Counts ---
            Dim trendSupport As Integer: trendSupport = 0
            Dim trendOppose As Integer: trendOppose = 0

            If (wsMain.cells(matchRow, "AS") < wsMain.cells(matchRow, "AU") And isBuy) Or _
               (wsMain.cells(matchRow, "AS") > wsMain.cells(matchRow, "AU") And Not isBuy) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            If (wsMain.cells(matchRow, "AX") < wsMain.cells(matchRow, "AZ") And isBuy) Or _
               (wsMain.cells(matchRow, "AX") > wsMain.cells(matchRow, "AZ") And Not isBuy) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            If (FL_5m < SL_5m And isBuy) Or (FL_5m > SL_5m And Not isBuy) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            If (FL_15m < SL_15m And isBuy) Or (FL_15m > SL_15m And Not isBuy) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            If (isBuy And CurrentPrice < superTrend) Or (Not isBuy And CurrentPrice > superTrend) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            ' Determine Trade Verdict
            Dim tradeVerdict As String
            If trendSupport >= 4 Then
                tradeVerdict = " STRONG TRADE    — Most trend indicators SUPPORT your " & tradeDirection & " (EMA 5m/15m, MACD, Supertrend)."
            ElseIf trendOppose >= 3 Then
                Dim opposeDesc As String: opposeDesc = ""
                If (wsMain.cells(matchRow, "AS") < wsMain.cells(matchRow, "AU") And isBuy) Or (wsMain.cells(matchRow, "AS") > wsMain.cells(matchRow, "AU") And Not isBuy) Then opposeDesc = opposeDesc & "EMA 10/50 5m, "
                If (wsMain.cells(matchRow, "AX") < wsMain.cells(matchRow, "AZ") And isBuy) Or (wsMain.cells(matchRow, "AX") > wsMain.cells(matchRow, "AZ") And Not isBuy) Then opposeDesc = opposeDesc & "EMA 10/50 15m, "
                If (FL_5m < SL_5m And isBuy) Or (FL_5m > SL_5m And Not isBuy) Then opposeDesc = opposeDesc & "MACD 5m, "
                If (FL_15m < SL_15m And isBuy) Or (FL_15m > SL_15m And Not isBuy) Then opposeDesc = opposeDesc & "MACD 15m, "
                If (isBuy And CurrentPrice < superTrend) Or (Not isBuy And CurrentPrice > superTrend) Then opposeDesc = opposeDesc & "Supertrend"
                ' inline trim trailing comma
                If Right(Trim(opposeDesc), 2) = ", " Then opposeDesc = Left(opposeDesc, Len(opposeDesc) - 2)
                tradeVerdict = " WEAK TRADE    — " & trendOppose & " indicators AGAINST: " & opposeDesc & "."
            Else
                tradeVerdict = " INDECISIVE    — " & trendSupport & " for vs " & trendOppose & " against your position (watch EMA/MACD)."
            End If

            nextStepsText = tradeVerdict & vbCrLf

            '--- Momentum Support (Stochastic, VWAP, Volume Spike) ---
            Dim momentumSupport As Integer: momentumSupport = 0
            Dim mixedSupportList As String: mixedSupportList = ""
            Dim mixedOpposeList As String: mixedOpposeList = ""

            ' Stochastic 5m
            If (isBuy And stochK > stochD) Or (Not isBuy And stochK < stochD) Then
                momentumSupport = momentumSupport + 1
                mixedSupportList = mixedSupportList & "Stochastic (5m), "
            Else
                mixedOpposeList = mixedOpposeList & "Stochastic (5m), "
            End If

            ' VWAP
            If (isBuy And CurrentPrice >= vwap) Or (Not isBuy And CurrentPrice <= vwap) Then
                momentumSupport = momentumSupport + 1
                mixedSupportList = mixedSupportList & "VWAP, "
            Else
                mixedOpposeList = mixedOpposeList & "VWAP, "
            End If

            ' Safe average volume for last up to 10 periods (inline numeric-safe loop)
            Dim sumVolume As Double: sumVolume = 0
            Dim countVolume As Long: countVolume = 0
            Dim volCell As Range
            Dim startRow As Long
            avgVolume10 = 0
            With wsMain
                lastDataRow = .cells(.Rows.count, "H").End(xlUp).row
                If lastDataRow >= 10 Then
                    startRow = lastDataRow - 9
                ElseIf lastDataRow > 0 Then
                    startRow = 1
                Else
                    startRow = 0
                End If

                If startRow > 0 Then
                    For Each volCell In .Range("H" & startRow & ":H" & lastDataRow)
                        If IsNumeric(volCell.value) And Not IsError(volCell.value) Then
                            sumVolume = sumVolume + volCell.value
                            countVolume = countVolume + 1
                        End If
                    Next volCell
                    If countVolume > 0 Then avgVolume10 = sumVolume / countVolume
                End If
            End With

            ' Volume spike: use currentVolume and avgVolume10; treat missing avgVolume10 conservatively
            Dim isVolumeSpike As Boolean: isVolumeSpike = False
            If avgVolume10 > 0 Then
                isVolumeSpike = (CurrentVolume > 1.2 * avgVolume10)
            Else
                isVolumeSpike = False
            End If

            If isVolumeSpike Then
                momentumSupport = momentumSupport + 1
                mixedSupportList = mixedSupportList & "Volume Spike, "
            Else
                mixedOpposeList = mixedOpposeList & "Volume, "
            End If

            ' Trim trailing commas inline
            If Right(mixedSupportList, 2) = ", " Then mixedSupportList = Left(mixedSupportList, Len(mixedSupportList) - 2)
            If Right(mixedOpposeList, 2) = ", " Then mixedOpposeList = Left(mixedOpposeList, Len(mixedOpposeList) - 2)

            '--- Momentum detailed explanation and action (MACD checks robust) ---
            'nextStepsText = nextStepsText & " MOMENTUM INDICATORS:" & vbCrLf

            ' prepare numeric-availability flags
            Dim haveFL5 As Boolean, haveSL5 As Boolean, haveM5 As Boolean, haveS5 As Boolean
            Dim haveFL15 As Boolean, haveSL15 As Boolean, haveM15 As Boolean, haveS15 As Boolean
            haveFL5 = IsNumeric(FL_5m)
            haveSL5 = IsNumeric(SL_5m)
            haveM5 = IsNumeric(MACDLine_5m)
            haveS5 = IsNumeric(SIG_5m)
            haveFL15 = IsNumeric(FL_15m)
            haveSL15 = IsNumeric(SL_15m)
            haveM15 = IsNumeric(MACDLine_15m)
            haveS15 = IsNumeric(SIG_15m)

            Dim macd5mAligned As Boolean: macd5mAligned = False
            Dim macd15mAligned As Boolean: macd15mAligned = False
            Dim eps As Double: eps = 0 ' set small epsilon if you want tolerance, e.g. 0.0001

            ' MACD 5m: require FL/SL and MACDLine/SIG when available; otherwise fall back to FL/SL only
            If haveFL5 And haveSL5 Then
                If haveM5 And haveS5 Then
                    If isBuy Then
                        macd5mAligned = ((FL_5m - SL_5m) >= -eps) And ((MACDLine_5m - SIG_5m) >= -eps)
                    Else
                        macd5mAligned = ((FL_5m - SL_5m) <= eps) And ((MACDLine_5m - SIG_5m) <= eps)
                    End If
                Else
                    If isBuy Then
                        macd5mAligned = (FL_5m - SL_5m) >= -eps
                    Else
                        macd5mAligned = (FL_5m - SL_5m) <= eps
                    End If
                End If
            Else
                macd5mAligned = False
            End If

            If macd5mAligned Then
                nextStepsText = nextStepsText & "   • MACD 5m: (ALIGNED) — Momentum supports your position. Maintain or consider adding if other signals confirm." & vbCrLf
            Else
                nextStepsText = nextStepsText & "   • MACD 5m: (NOT ALIGNED) — MOMENTUM WEAKENING. Consider tightening stop-loss or reducing exposure." & vbCrLf
            End If

            ' MACD 15m
            If haveFL15 And haveSL15 Then
                If haveM15 And haveS15 Then
                    If isBuy Then
                        macd15mAligned = ((FL_15m - SL_15m) >= -eps) And ((MACDLine_15m - SIG_15m) >= -eps)
                    Else
                        macd15mAligned = ((FL_15m - SL_15m) <= eps) And ((MACDLine_15m - SIG_15m) <= eps)
                    End If
                Else
                    If isBuy Then
                        macd15mAligned = (FL_15m - SL_15m) >= -eps
                    Else
                        macd15mAligned = (FL_15m - SL_15m) <= eps
                    End If
                End If
            Else
                macd15mAligned = False
            End If

            If macd15mAligned Then
                nextStepsText = nextStepsText & "   • MACD 15m: (ALIGNED) — Higher-timeframe momentum supports your position. Positive confirmation." & vbCrLf
            Else
                nextStepsText = nextStepsText & "   • MACD 15m: (NOT ALIGNED) — Higher-timeframe momentum does not confirm. Monitor for reversal." & vbCrLf
            End If

            ' Stochastic 5m: concise output
            If (isBuy And stochK > stochD) Or (Not isBuy And stochK < stochD) Then
                nextStepsText = nextStepsText & "   • Stochastic 5m: (ALIGNED) (%K " & IIf(isBuy, ">", "<") & " %D) — MOMENTUM STRENGTHENING. Consider holding or adding." & vbCrLf
            Else
                nextStepsText = nextStepsText & "   • Stochastic 5m: (NOT ALIGNED) (%K " & IIf(isBuy, "<", ">") & " %D) — MOMENTUM WEAKENING. Exercise caution; consider tightening stops." & vbCrLf
            End If

            '--- Plan ---
            'nextStepsText = nextStepsText & vbCrLf & " PLAN" & vbCrLf
            nextStepsText = nextStepsText & vbCrLf & IIf(isBuy, _
                "   BREAK ABOVE " & r1 & " -- Confirmation(ADD / HOLD); " & _
                "   STALL BELOW " & pp & " -- Weakness (TRIM); " & _
                "   FALL BELOW " & slPrice & " -- SL breached (EXIT); ", _
                "   BREAK BELOW " & s1 & " -- Stay Short; " & _
                "   STALL ABOVE " & pp & " -- Reversal Risk; " & _
                "   RISE ABOVE " & slPrice & " -- SL breached (EXIT)")

            ' --- Sell SL Fix (conservative flip if SL illogical) ---
            If tradeDirection = "sell" And slPrice <= entryPrice Then
                ' If slPrice is below entry for a sell, flip it symmetrically above entry
                slPrice = entryPrice + (entryPrice - slPrice)
            End If

            '--- Stop Loss and Risk ---
            Dim riskPercent As Double
            riskPercent = 0
            On Error Resume Next
            If IsNumeric(entryPrice) And entryPrice <> 0 Then
                riskPercent = Abs((slPrice - entryPrice) / entryPrice) * 100
            End If
            On Error GoTo 0

            nextStepsText = nextStepsText & vbCrLf & _
                "- Maintain stop-loss at " & Format(slPrice, "0.00") & " (~" & Format(riskPercent, "0.00") & "% risk relative to entry)." & _
                " Adjust Qty based on your risk tolerance." & vbCrLf

            '--- Volume Info ---
            nextStepsText = nextStepsText & _
                "- Volume spike threshold: volume > 120% of recent 10-period average." & " " & _
                " Current volume is " & IIf(CurrentVolume > 1.2 * avgVolume10, "above", "below") & " this threshold." & vbCrLf

            '--- Confidence Level ---
            Dim confidenceLabelText As String
            Dim confidenceExplanation As String
            ' confidenceScore left intact (if set elsewhere); default 0 => LOW
            If confidenceScore >= 6 Then
                confidenceLabelText = "HIGH"
                confidenceExplanation = "(Strong support from indicators; favorable conditions.)"
            ElseIf confidenceScore >= 4 Then
                confidenceLabelText = "MEDIUM"
                confidenceExplanation = "(Moderate support; monitor momentum closely.)"
            Else
                confidenceLabelText = "LOW"
                confidenceExplanation = "(Limited support; higher risk of false signals.)"
            End If

            nextStepsText = nextStepsText & vbCrLf & _
                            "  CONFIDENCE LEVEL: " & confidenceLabelText & " " & confidenceExplanation & vbCrLf & _
                            "  Target: " & Format(IIf(isBuy, r2, s2), "0.00") & " ; " & _
                            "  SL: " & Format(slPrice, "0.00")

            ' Write to cell
            With cellNextSteps
                .ClearContents
                .value = RTrim(nextStepsText)
                .WrapText = True
                .VerticalAlignment = xlTop
                .Characters.Font.size = 10
                '.Rows.AutoFit
            End With

            ' --- Keyword Highlight (optimized Len per word) ---
            Dim keywordMap1 As Variant, wLen As Long
            keywordMap1 = Array( _
                          Array("WEAK TRADE", COLOR_RED), Array("STRONG TRADE", COLOR_GREEN), Array("MOMENTUM", COLOR_BLUE), Array("PLAN", COLOR_BLUE), Array("CONFIDENCE LEVEL", COLOR_BLUE), Array("MEDIUM", COLOR_YELLOW), _
                          Array("WEAK", COLOR_RED), Array("STRONG", COLOR_GREEN), Array("INDECISIVE", COLOR_YELLOW), Array("AGAINST", COLOR_RED), Array("FOR", COLOR_GREEN), Array("Reversal Risk", COLOR_YELLOW), _
                          Array("ALIGNED", COLOR_GREEN), Array("TRIM", COLOR_YELLOW), Array("Opposite Position", COLOR_RED), Array("Weakness", COLOR_YELLOW), Array("Confirmation", COLOR_GREEN), _
                          Array("BREAK BELOW", COLOR_BLUE), Array("BREAK ABOVE", COLOR_BLUE), Array("STALL BELOW", COLOR_BLUE), Array("FALL BELOW", COLOR_BLUE), Array("Target", COLOR_GREEN), Array("SL breached", COLOR_RED), _
                          Array("SL", COLOR_RED), Array("LOW", COLOR_RED), Array("HIGH", COLOR_GREEN), Array("EXIT", COLOR_RED), Array("VWAP", COLOR_BLUE), Array("ISSUES", COLOR_RED), _
                          Array("ADD", COLOR_GREEN), Array("HOLD", COLOR_GREEN), Array("Volume", COLOR_BLUE), Array("Spike", COLOR_GREEN), Array("No Spike", COLOR_RED), Array("Stay Short", COLOR_GREEN), _
                          Array("BREAK BELOW", COLOR_BLUE), Array("STALL ABOVE", COLOR_BLUE), Array("RISE ABOVE", COLOR_BLUE), Array("SUPPORT", COLOR_GREEN), Array("Stochastic", COLOR_BLUE), _
                          Array("Supertrend", COLOR_BLUE), Array("EMA 10/50 5m", COLOR_BLUE), Array("EMA 10/50 15m", COLOR_BLUE), Array("EMA 10/30 5m", COLOR_BLUE), Array("EMA 10/30 15m", COLOR_BLUE), _
                          Array("Bullish", COLOR_GREEN), Array("BEARISH", COLOR_RED), Array("MIXED", COLOR_YELLOW), Array("NOT ALIGNED", COLOR_RED), Array("MOMENTUM WEAKENING", COLOR_RED), Array("MOMENTUM STRENGTHENING", COLOR_GREEN), _
                          Array("EMA 5m/15m", COLOR_BLUE), Array("MACD", COLOR_BLUE), Array("MACD 15m", COLOR_BLUE), Array("MACD 5m", COLOR_BLUE), Array("EMA 10/30 15m", COLOR_BLUE))

            For p = 0 To UBound(keywordMap1)
                word = keywordMap1(p)(0)
                colorVal = keywordMap1(p)(1)
                wLen = Len(word)
                pos = 1
                Do
                    pos = InStr(pos, nextStepsText, word, vbBinaryCompare)
                    If pos = 0 Then Exit Do
                    On Error Resume Next
                    cellNextSteps.Characters(pos, wLen).Font.Color = colorVal
                    cellNextSteps.Characters(pos, wLen).Font.Bold = True
                    On Error GoTo 0
                    pos = pos + wLen
                Loop
            Next p
            cellNextSteps.Rows.AutoFit
        End If

SkipRow:
    Next i

    Application.ScreenUpdating = True
End Sub

' --------------------------------------------------------------------------
' UpdateSmartNextSteps - improved decision engine & concise next-steps
' - Preserves original indicator reads and cell layout.
' - Adds a weighted scoring engine, concise Next Steps, and verbose detail.
' - Color-codes outputs and retains your keyword-highlighting behavior.
' --------------------------------------------------------------------------
Sub UpdateSmartNextSteps_V2()
    Dim wsLive As Worksheet, wsMain As Worksheet
    Dim i As Long, lastRow As Long
    Dim ticker As String, action As String, status As String
    Dim matchRow As Variant, nowStamp As String
    Dim nextStepMsg As String, fullLine As String
    Dim entryPrice As Double, slPrice As Double, CurrentPrice As Double
    Dim RSI_5m As Double, rsi_15m As Double, vwap As Double, VolumeSpike As Variant
    Dim MACD_5m As String, macd_15m As String, EMA_5m As String, EMA_15m As String
    Dim superTrend As Double
    Dim BB_U As Double, BB_L As Double, ATR_5m As Double
    Dim trailSL As Double, tightenSL As Double
    Dim s1 As Double, s2 As Double, s3 As Double, r1 As Double, r2 As Double, r3 As Double, pp As Double
    Dim highVal As Variant, lowVal As Variant, closeVal As Variant
    Dim gainPct As Double
    Dim cellHistory As Range, celllatest As Range
    Dim trendOutlook As String, slAssessment As String, reversalWatch As String, rsiDivergence As String, SandR As String
    Dim candleBull As String, candleBear As String
    Dim reversalConfidence As Integer
    Dim PriceDiff As Double, tolerance As Double
    Dim quantity As Double
    Dim stochK As Double, stochD As Double, stochComment As String, bollComment As String
    Dim FL_5m As Double, SL_5m As Double, SIG_5m As Double
    Dim FL_15m As Double, SL_15m As Double, SIG_15m As Double
    Dim MACDLine_5m As Double, MACDLine_15m As Double
    Dim CurrentVolume As Double
    Dim confidenceScore As Integer
    Dim lastDataRow As Long

    Set wsLive = ThisWorkbook.Sheets("Live_New")
    Set wsMain = ThisWorkbook.Sheets("Sheet1")
    lastRow = wsLive.cells(wsLive.Rows.count, "A").End(xlUp).row

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        momentumList = ""
        nowStamp = "  DATE: " & Format(Now, "dd-mmm hh:mm")
        status = Trim(wsLive.cells(i, "I").value)
        'If status <> "O" Then GoTo SkipRow

        ticker = Trim(wsLive.cells(i, "A").value)
        action = LCase(Trim(wsLive.cells(i, "B").value))
        quantity = val(wsLive.cells(i, "C").value)
        entryPrice = val(wsLive.cells(i, "D").value)
        slPrice = val(wsLive.cells(i, "F").value)

        Set celllatest = wsLive.cells(i, "H")

        matchRow = Application.Match(ticker, wsMain.Range("A:A"), 0)
        If IsError(matchRow) Then GoTo SkipRow
        ' [V2] Nifty regime + ML lookup + lifecycle
        Dim niftyRegime As String: niftyRegime = "RANGE"
        Dim nseRow As Variant
        nseRow = Application.Match("NSEI", wsMain.Range("A:A"), 0)
        If Not IsError(nseRow) Then
            Dim ddV2 As String: ddV2 = UCase(Trim(wsMain.cells(nseRow, "DD").value))
            If ddV2 <> "" Then niftyRegime = ddV2
        End If
        Dim mlConf As Double: mlConf = 0
        Dim mlSignal As String: mlSignal = ""
        Dim mlGrade As String: mlGrade = ""
        Dim mlSheet As Worksheet, mlRowV As Variant
        On Error Resume Next
        Set mlSheet = ThisWorkbook.Sheets("AI_Master")
        On Error GoTo 0
        If Not mlSheet Is Nothing Then
            mlRowV = Application.Match(ticker, mlSheet.Range("A:A"), 0)
            If Not IsError(mlRowV) Then
                mlConf = val(mlSheet.cells(mlRowV, "H").value)
                mlSignal = UCase(Trim(mlSheet.cells(mlRowV, "D").value))
                mlGrade = Trim(mlSheet.cells(mlRowV, "I").value)
            End If
        End If
        Dim barsHeld As Long: barsHeld = 0
        If IsNumeric(wsLive.cells(i, "K").value) Then barsHeld = CLng(wsLive.cells(i, "K").value)
        barsHeld = barsHeld + 1
        wsLive.cells(i, "K").value = barsHeld
        Dim isBuy As Boolean: isBuy = (LCase(Trim(action)) = "buy")

        With wsMain
            RSI_5m = val(.cells(matchRow, "U").value)
            rsi_15m = val(.cells(matchRow, "V").value)
            VolumeSpike = .cells(matchRow, "BS").value
            vwap = val(.cells(matchRow, "BR").value)
            superTrend = val(.cells(matchRow, "BT").value)
            BB_U = val(.cells(matchRow, "CI").value)
            BB_L = val(.cells(matchRow, "CK").value)
            ATR_5m = val(.cells(matchRow, "CC").value)
            CurrentPrice = val(.cells(matchRow, "T").value)
            CurrentVolume = val(.cells(matchRow, "H").value) ' assume column H holds current volume for that ticker
            highVal = .cells(matchRow, "BV").value
            lowVal = .cells(matchRow, "BW").value
            closeVal = .cells(matchRow, "BX").value
            candleBull = .cells(matchRow, "AK").value & .cells(matchRow, "AL").value
            candleBear = .cells(matchRow, "AO").value & .cells(matchRow, "AP").value

            stochK = val(.cells(matchRow, "CE").value)
            stochD = val(.cells(matchRow, "CF").value)

            FL_5m = val(.cells(matchRow, "CY").value)
            SL_5m = val(.cells(matchRow, "CZ").value)
            SIG_5m = val(.cells(matchRow, "DA").value)
            FL_15m = val(.cells(matchRow, "DB").value)
            SL_15m = val(.cells(matchRow, "DC").value)
            SIG_15m = val(.cells(matchRow, "DD").value)
        End With

        If stochK > stochD Then
            stochComment = "  STOCHASTIC: Bullish crossover on 5m - Confirmation of Upward Momentum"
        ElseIf stochK < stochD Then
            stochComment = "  STOCHASTIC: Bearish crossover on 5m - Possible Downside Pressure"
        Else
            stochComment = "  STOCHASTIC: No crossover - Stochastics are in sync or flat"
        End If

        If CurrentPrice > BB_U Then
            bollComment = "  BOLLINGER: Price above upper Bollinger Band - Overbought/Breakout scenario"
        ElseIf CurrentPrice < BB_L Then
            bollComment = "  BOLLINGER: Price dropped below lower Bollinger Band - Potential Mean Reversion/Oversold"
        Else
            bollComment = "  BOLLINGER: Price within bands - no edge from Bollinger signal"
        End If

        s3 = Round(wsMain.cells(matchRow, "K").value, 2)
        s2 = Round(wsMain.cells(matchRow, "L").value, 2)
        s1 = Round(wsMain.cells(matchRow, "M").value, 2)
        pp = Round(wsMain.cells(matchRow, "N").value, 2)
        r1 = Round(wsMain.cells(matchRow, "O").value, 2)
        r2 = Round(wsMain.cells(matchRow, "P").value, 2)
        r3 = Round(wsMain.cells(matchRow, "Q").value, 2)

        SandR = "   (S3:  " & s3 & ") (S2:  " & s2 & ") (S1:  " & s1 & ") (PP:  " & pp & ") (R1:  " & r1 & ") (R2:  " & r2 & ") (R3:  " & r3 & ")"

        PriceDiff = Round(CurrentPrice - entryPrice, 2)
        nowStamp = nowStamp & "  ||  CURRENT PRICE: " & CStr(CurrentPrice) & " "
        wsLive.cells(i, "G").value = Abs(PriceDiff * quantity)
        wsLive.cells(i, "G").Font.Bold = True

        Select Case Trim(LCase(action))
        Case "buy"
            If entryPrice >= s3 And entryPrice < s2 Then
                wsLive.cells(i, "E").value = s2
                wsLive.cells(i, "F").value = s3
                nowStamp = nowStamp & "(Between S3 and S2)"
            ElseIf entryPrice >= s2 And entryPrice < s1 Then
                wsLive.cells(i, "E").value = s1
                wsLive.cells(i, "F").value = s2
                nowStamp = nowStamp & "(Between S2 and S1)"
            ElseIf entryPrice >= s1 And entryPrice < pp Then
                wsLive.cells(i, "E").value = pp
                wsLive.cells(i, "F").value = s1
                nowStamp = nowStamp & "(Between S1 and PP)"
            ElseIf entryPrice >= pp And entryPrice < r1 Then
                wsLive.cells(i, "E").value = r1
                wsLive.cells(i, "F").value = pp
                nowStamp = nowStamp & "(Between PP and R1)"
            ElseIf entryPrice >= r1 And entryPrice < r2 Then
                wsLive.cells(i, "E").value = r2
                wsLive.cells(i, "F").value = r1
                nowStamp = nowStamp & "(Between R1 and R2)"
            ElseIf entryPrice >= r2 And entryPrice <= r3 Then
                wsLive.cells(i, "E").value = r3
                wsLive.cells(i, "F").value = r2
                nowStamp = nowStamp & "(Between R2 and R3)"
            End If
        Case "sell"
            If entryPrice <= r3 And entryPrice > r2 Then
                wsLive.cells(i, "E").value = r1
                wsLive.cells(i, "F").value = r3
                nowStamp = nowStamp & "(Between R2 and R3)"
            ElseIf entryPrice <= r2 And entryPrice > r1 Then
                wsLive.cells(i, "E").value = pp
                wsLive.cells(i, "F").value = r2
                nowStamp = nowStamp & "(Between R1 and R2)"
            ElseIf entryPrice <= r1 And entryPrice > pp Then
                wsLive.cells(i, "E").value = s1
                wsLive.cells(i, "F").value = r1
                nowStamp = nowStamp & "(Between PP and R1)"
            ElseIf entryPrice <= pp And entryPrice > s1 Then
                wsLive.cells(i, "E").value = s2
                wsLive.cells(i, "F").value = pp
                nowStamp = nowStamp & "(Between S1 and PP)"
            ElseIf entryPrice <= s1 And entryPrice > s2 Then
                wsLive.cells(i, "E").value = s3
                wsLive.cells(i, "F").value = s1
                nowStamp = nowStamp & "(Between S2 and S1)"
            ElseIf entryPrice <= s2 And entryPrice >= s3 Then
                wsLive.cells(i, "E").value = s3
                wsLive.cells(i, "F").value = s2
                nowStamp = nowStamp & "(Between S3 and S2)"
            End If
        End Select

        Select Case Trim(LCase(action))
            Case "sell"
                wsLive.cells(i, "G").Font.Color = IIf(PriceDiff >= 0, RGB(255, 0, 0), RGB(0, 180, 0))
            Case "buy"
                wsLive.cells(i, "G").Font.Color = IIf(PriceDiff >= 0, RGB(0, 180, 0), RGB(255, 0, 0))
        End Select

        Dim CandleStickPattern As String
        If candleBear = "" Or candleBull = "" Then
            CandleStickPattern = " No discernable candle patterns deteted"
        Else
            If (action = "buy" And candleBear <> "") Then CandleStickPattern = "  Bearish Reversal candle : " & candleBear
            If (action = "sell" And candleBull <> "") Then CandleStickPattern = "  Bullish Reversal candle : " & candleBull
        End If

        Dim strSupertrend As String
        tolerance = 0.001 * CurrentPrice
        If (action = "buy" And CurrentPrice < superTrend - tolerance) Or _
           (action = "sell" And CurrentPrice > superTrend + tolerance) Then
            strSupertrend = "    SUPERTREND (" & CStr(superTrend) & ") - Opposite to position; "
            reversalConfidence = reversalConfidence + 2
        Else
            strSupertrend = "  SUPERTREND: " & CStr(superTrend)
        End If

        Dim strVol As String
        If VolumeSpike = 1 Then
            Select Case LCase(action)
                Case "buy":
                    If (CurrentPrice < vwap) Then
                        strVol = "  Buy setup losing support - Volume Spike but price below VWAP"
                        reversalConfidence = reversalConfidence + 2
                    Else
                        strVol = "  HOLD BUY - Volume Spike detected and price above VWAP"
                    End If
                Case "sell":
                    If (CurrentPrice > vwap) Then
                        strVol = "  Sell setup under pressure - Volume Spike but price above VWAP"
                        reversalConfidence = reversalConfidence + 2
                    Else
                        strVol = "  HOLD SELL - Volume Spike detected and price below VWAP"
                    End If
            End Select
        Else
            strVol = "  No Volume Spike observed"
        End If

        If RSI_5m > 70 And rsi_15m < 65 Then
            rsiDivergence = "RSI divergence - overbought 5m, weakening 15m"
        ElseIf RSI_5m < 30 And rsi_15m > 35 Then
            rsiDivergence = "  RSI Divergence - oversold 5m, recovering 15m"
        Else
            rsiDivergence = "  RSI Divergence - None Observed"
        End If

        If InStr(reversalWatch, "Reversal") > 0 Then reversalConfidence = reversalConfidence + 3
        If rsiDivergence <> "" Then reversalConfidence = reversalConfidence + 3
        If reversalWatch <> "" Then
            reversalWatch = "Reversal Risk: " & reversalWatch & rsiDivergence
            If reversalConfidence >= 8 Then reversalWatch = reversalWatch & " >> Candle Watch across 5m-15m suggests potential reversal; wait for confirmation."
        End If

        If wsMain.cells(matchRow, "AS").value >= wsMain.cells(matchRow, "AT").value Then
            trendOutlook = "  EMA(10_30) : 5m - Bullish "
        Else
            trendOutlook = "  EMA(10_30) : 5m - Bearish "
        End If

        If wsMain.cells(matchRow, "AX").value >= wsMain.cells(matchRow, "AY").value Then
            trendOutlook = trendOutlook & " || 15m - Bullish  "
        Else
            trendOutlook = trendOutlook & " || 15m - Bearish  "
        End If

        If wsMain.cells(matchRow, "AS").value >= wsMain.cells(matchRow, "AU").value Then
            trendOutlook = trendOutlook & "  EMA(10_50) : 5m - Bullish  "
        Else
            trendOutlook = trendOutlook & "  EMA(10_50) : 5m - Bearish  "
        End If

        If wsMain.cells(matchRow, "AX").value >= wsMain.cells(matchRow, "AZ").value Then
            trendOutlook = trendOutlook & " || 15m - Bullish "
        Else
            trendOutlook = trendOutlook & " || 15m - Bearish"
        End If

        MACDLine_5m = FL_5m - SL_5m
        MACDLine_15m = FL_15m - SL_15m
        macdNarrative = ""

        If FL_5m >= SL_5m Then
            macdNarrative = "  MACD(F_S): 5m - Bullish  "
        Else
            macdNarrative = "  MACD(F_S): 5m - Bearish  "
        End If

        If FL_15m >= SL_15m Then
            macdNarrative = macdNarrative & " ||  15m - Bullish  "
        Else
            macdNarrative = macdNarrative & " ||  15m - Bearish  "
        End If

        If MACDLine_5m >= SIG_5m Then
            macdNarrative = macdNarrative & "  MACD(M_S): 5m - Bullish  "
        Else
            macdNarrative = macdNarrative & "  MACD(M_S): 5m - Bearish  "
        End If

        If MACDLine_15m >= SIG_15m Then
            macdNarrative = macdNarrative & " || 15m - Bullish"
        Else
            macdNarrative = macdNarrative & " || 15m - Bearish"
        End If

        Dim SLgap As Double
        SLgap = Abs(CurrentPrice - slPrice)
        If SLgap <= ATR_5m Then
            slAssessment = "[SL Warning] SL may be too tight"
        ElseIf (action = "buy" And slPrice < s3) Or (action = "sell" And slPrice > r3) Then
            slAssessment = "[Wide SL] SL beyond extreme levels"
        Else
            slAssessment = "SL in safe zone"
        End If

        Dim pctGap As Double
        pctGap = Abs(CurrentPrice - slPrice) / CurrentPrice

        If action = "sell" Then
            If (CurrentPrice - slPrice) <= 0 Then
                If pctGap <= 0.1 Then
                    slComment = "  SL is within 10% of current price"
                Else
                    slComment = "  SL is " & CStr(pctGap * 100) & "% away from current price"
                End If
            Else
                slComment = "  SL Breached"
            End If
        Else
            If (CurrentPrice - slPrice) >= 0 Then
                If pctGap <= 0.1 Then
                    slComment = "  SL is within 10% of current price"
                Else
                    slComment = "  SL is " & CStr(pctGap * 100) & "% away from current price"
                End If
            Else
                slComment = "  SL Breached"
            End If
        End If

        Dim strRSI As String
        strRSI = "   RSI: (5m: " & Format(RSI_5m, "0") & ") (15m: " & Format(rsi_15m, "0") & _
               ") (1H: " & Format(val(wsMain.cells(matchRow, "W").value)) & _
               ") (1D: " & Format(val(wsMain.cells(matchRow, "X").value) & ")")

        fullLine = nowStamp & vbCrLf & _
                    strRSI & vbCrLf & _
                    SandR & vbCrLf & _
                    "------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
                    trendOutlook & vbCrLf & _
                    macdNarrative & vbCrLf & _
                    strSupertrend & vbCrLf & _
                    "------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
                    stochComment & vbCrLf & _
                    bollComment & vbCrLf & _
                    "------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
                    CandleStickPattern & vbCrLf & _
                    strVol & vbCrLf & _
                    slComment & vbCrLf & vbCrLf & _
                    "------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & _
                    "  NIFTY REGIME: " & niftyRegime & vbCrLf & _
                    "  ML SIGNAL: " & IIf(mlSignal <> "", mlSignal, "None") & _
                    IIf(mlSignal <> "", " | Conf: " & Format(mlConf, "0.00") & " | Grade: " & mlGrade, "") & _
                    IIf(mlSignal <> "" And ((mlSignal = "BUY_NOW" And isBuy) Or (mlSignal = "SELL_NOW" And Not isBuy)), " [ALIGNED]", IIf(mlSignal <> "", " [DIVERGES]", ""))


        With celllatest
            .ClearContents
            .Font.Color = RGB(0, 0, 0)
            .value = RTrim(fullLine)
        End With

        Dim COLOR_RED As Long: COLOR_RED = RGB(200, 0, 0)
        Dim COLOR_GREEN As Long: COLOR_GREEN = RGB(0, 160, 0)
        Dim COLOR_BLUE As Long: COLOR_BLUE = RGB(0, 0, 255)
        Dim COLOR_YELLOW As Long: COLOR_YELLOW = RGB(200, 140, 0)

        Dim keywordMap As Variant

        keywordMap = Array( _
            Array("MOMENTUM", COLOR_YELLOW), Array("PIVOTS", COLOR_YELLOW), Array("INDICATORS", COLOR_YELLOW), Array("RISK CHECKS", COLOR_YELLOW), _
            Array("EMA(10_30)", COLOR_BLUE), Array("EMA(10_50)", COLOR_BLUE), Array("MACD(F_S)", COLOR_BLUE), Array("MACD(M_S)", COLOR_BLUE), Array("SUPERTREND", COLOR_BLUE), _
            Array("STOCHASTIC", COLOR_BLUE), Array("BOLLINGER", COLOR_BLUE), Array("RSI", COLOR_BLUE), Array("5m", COLOR_BLUE), Array("15m", COLOR_BLUE), Array("1H", COLOR_BLUE), Array("1D", COLOR_BLUE), _
            Array("S3", COLOR_RED), Array("S2", COLOR_RED), Array("S1", COLOR_RED), Array("PP", COLOR_BLUE), Array("R1", COLOR_GREEN), Array("R2", COLOR_GREEN), Array("R3", COLOR_GREEN), _
            Array("Bearish", COLOR_RED), Array("Bullish", COLOR_GREEN), Array("Opposite to position", COLOR_RED), _
            Array("Volume Spike", COLOR_BLUE), Array("VWAP", COLOR_BLUE), Array("SL", COLOR_RED), Array("HOLD", COLOR_BLUE), _
            Array("Reversal candle", COLOR_BLUE), Array("Overbought/Breakout", COLOR_BLUE), Array("Potential Mean Reversion/Oversold", COLOR_BLUE), _
            Array("Confirmation of Upward Momentum", COLOR_BLUE), Array("Possible Downside Pressure", COLOR_BLUE), _
            Array("DATE", COLOR_BLUE), Array("CURRENT PRICE", COLOR_BLUE))

        Dim p As Long, pos As Long
        Dim word As String, colorVal As Long

        For p = 0 To UBound(keywordMap)
            word = keywordMap(p)(0)
            colorVal = keywordMap(p)(1)
            pos = 1
            Do
                pos = InStr(pos, fullLine, word, vbBinaryCompare)
                If pos = 0 Then Exit Do
                On Error Resume Next
                celllatest.Characters(pos, Len(word)).Font.Color = colorVal
                celllatest.Characters(pos, Len(word)).Font.Bold = True
                On Error GoTo 0
                pos = pos + Len(word)
            Loop
        Next p
        
        'celllatest.EntireRow.RowHeight = 300
        celllatest.WrapText = True
        celllatest.VerticalAlignment = xlTop
        'celllatest.Characters.Font.Size = 10
        'Debug.Print celllatest.EntireRow.RowHeight
        'Debug.Print ActiveSheet.ProtectContents
        
       ' === COLUMN J NEXT STEPS ===
        If wsLive.cells(i, "G") <> "" Then
            Dim nextStepsText As String
            Dim tradeDirection As String: tradeDirection = LCase(Trim(action))
            'Dim isBuy As Boolean: isBuy = (tradeDirection = "buy")
            isBuy = (tradeDirection = "buy")
            Set cellNextSteps = wsLive.cells(i, "J")

            '--- Trend Support and Oppose Counts ---
            Dim trendSupport As Integer: trendSupport = 0
            Dim trendOppose As Integer: trendOppose = 0

            If (wsMain.cells(matchRow, "AS") < wsMain.cells(matchRow, "AU") And isBuy) Or _
               (wsMain.cells(matchRow, "AS") > wsMain.cells(matchRow, "AU") And Not isBuy) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            If (wsMain.cells(matchRow, "AX") < wsMain.cells(matchRow, "AZ") And isBuy) Or _
               (wsMain.cells(matchRow, "AX") > wsMain.cells(matchRow, "AZ") And Not isBuy) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            If (FL_5m < SL_5m And isBuy) Or (FL_5m > SL_5m And Not isBuy) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            If (FL_15m < SL_15m And isBuy) Or (FL_15m > SL_15m And Not isBuy) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            If (isBuy And CurrentPrice < superTrend) Or (Not isBuy And CurrentPrice > superTrend) Then
                trendOppose = trendOppose + 1
            Else
                trendSupport = trendSupport + 1
            End If

            ' Determine Trade Verdict
            Dim tradeVerdict As String
            If trendSupport >= 4 Then
                tradeVerdict = " STRONG TRADE    -- Most trend indicators SUPPORT your " & tradeDirection & " (EMA 5m/15m, MACD, Supertrend)."
            ElseIf trendOppose >= 3 Then
                Dim opposeDesc As String: opposeDesc = ""
                If (wsMain.cells(matchRow, "AS") < wsMain.cells(matchRow, "AU") And isBuy) Or (wsMain.cells(matchRow, "AS") > wsMain.cells(matchRow, "AU") And Not isBuy) Then opposeDesc = opposeDesc & "EMA 10/50 5m, "
                If (wsMain.cells(matchRow, "AX") < wsMain.cells(matchRow, "AZ") And isBuy) Or (wsMain.cells(matchRow, "AX") > wsMain.cells(matchRow, "AZ") And Not isBuy) Then opposeDesc = opposeDesc & "EMA 10/50 15m, "
                If (FL_5m < SL_5m And isBuy) Or (FL_5m > SL_5m And Not isBuy) Then opposeDesc = opposeDesc & "MACD 5m, "
                If (FL_15m < SL_15m And isBuy) Or (FL_15m > SL_15m And Not isBuy) Then opposeDesc = opposeDesc & "MACD 15m, "
                If (isBuy And CurrentPrice < superTrend) Or (Not isBuy And CurrentPrice > superTrend) Then opposeDesc = opposeDesc & "Supertrend"
                ' inline trim trailing comma
                If Right(Trim(opposeDesc), 2) = ", " Then opposeDesc = Left(opposeDesc, Len(opposeDesc) - 2)
                tradeVerdict = " WEAK TRADE    -- " & trendOppose & " indicators AGAINST: " & opposeDesc & "."
            Else
                tradeVerdict = " INDECISIVE    -- " & trendSupport & " for vs " & trendOppose & " against your position (watch EMA/MACD)."
            End If

            nextStepsText = tradeVerdict & vbCrLf

            '--- Momentum Support (Stochastic, VWAP, Volume Spike) ---
            Dim momentumSupport As Integer: momentumSupport = 0
            Dim mixedSupportList As String: mixedSupportList = ""
            Dim mixedOpposeList As String: mixedOpposeList = ""

            ' Stochastic 5m
            If (isBuy And stochK > stochD) Or (Not isBuy And stochK < stochD) Then
                momentumSupport = momentumSupport + 1
                mixedSupportList = mixedSupportList & "Stochastic (5m), "
            Else
                mixedOpposeList = mixedOpposeList & "Stochastic (5m), "
            End If

            ' VWAP
            If (isBuy And CurrentPrice >= vwap) Or (Not isBuy And CurrentPrice <= vwap) Then
                momentumSupport = momentumSupport + 1
                mixedSupportList = mixedSupportList & "VWAP, "
            Else
                mixedOpposeList = mixedOpposeList & "VWAP, "
            End If

            ' Safe average volume for last up to 10 periods (inline numeric-safe loop)
            Dim sumVolume As Double: sumVolume = 0
            Dim countVolume As Long: countVolume = 0
            Dim volCell As Range
            Dim startRow As Long
            avgVolume10 = 0
            With wsMain
                lastDataRow = .cells(.Rows.count, "H").End(xlUp).row
                If lastDataRow >= 10 Then
                    startRow = lastDataRow - 9
                ElseIf lastDataRow > 0 Then
                    startRow = 1
                Else
                    startRow = 0
                End If

                If startRow > 0 Then
                    For Each volCell In .Range("H" & startRow & ":H" & lastDataRow)
                        If IsNumeric(volCell.value) And Not IsError(volCell.value) Then
                            sumVolume = sumVolume + volCell.value
                            countVolume = countVolume + 1
                        End If
                    Next volCell
                    If countVolume > 0 Then avgVolume10 = sumVolume / countVolume
                End If
            End With

            ' Volume spike: use currentVolume and avgVolume10; treat missing avgVolume10 conservatively
            Dim isVolumeSpike As Boolean: isVolumeSpike = False
            If avgVolume10 > 0 Then
                isVolumeSpike = (CurrentVolume > 1.2 * avgVolume10)
            Else
                isVolumeSpike = False
            End If

            If isVolumeSpike Then
                momentumSupport = momentumSupport + 1
                mixedSupportList = mixedSupportList & "Volume Spike, "
            Else
                mixedOpposeList = mixedOpposeList & "Volume, "
            End If

            ' Trim trailing commas inline
            If Right(mixedSupportList, 2) = ", " Then mixedSupportList = Left(mixedSupportList, Len(mixedSupportList) - 2)
            If Right(mixedOpposeList, 2) = ", " Then mixedOpposeList = Left(mixedOpposeList, Len(mixedOpposeList) - 2)

            '--- Momentum detailed explanation and action (MACD checks robust) ---
            'nextStepsText = nextStepsText & " MOMENTUM INDICATORS:" & vbCrLf

            ' prepare numeric-availability flags
            Dim haveFL5 As Boolean, haveSL5 As Boolean, haveM5 As Boolean, haveS5 As Boolean
            Dim haveFL15 As Boolean, haveSL15 As Boolean, haveM15 As Boolean, haveS15 As Boolean
            haveFL5 = IsNumeric(FL_5m)
            haveSL5 = IsNumeric(SL_5m)
            haveM5 = IsNumeric(MACDLine_5m)
            haveS5 = IsNumeric(SIG_5m)
            haveFL15 = IsNumeric(FL_15m)
            haveSL15 = IsNumeric(SL_15m)
            haveM15 = IsNumeric(MACDLine_15m)
            haveS15 = IsNumeric(SIG_15m)

            Dim macd5mAligned As Boolean: macd5mAligned = False
            Dim macd15mAligned As Boolean: macd15mAligned = False
            Dim eps As Double: eps = 0 ' set small epsilon if you want tolerance, e.g. 0.0001

            ' MACD 5m: require FL/SL and MACDLine/SIG when available; otherwise fall back to FL/SL only
            If haveFL5 And haveSL5 Then
                If haveM5 And haveS5 Then
                    If isBuy Then
                        macd5mAligned = ((FL_5m - SL_5m) >= -eps) And ((MACDLine_5m - SIG_5m) >= -eps)
                    Else
                        macd5mAligned = ((FL_5m - SL_5m) <= eps) And ((MACDLine_5m - SIG_5m) <= eps)
                    End If
                Else
                    If isBuy Then
                        macd5mAligned = (FL_5m - SL_5m) >= -eps
                    Else
                        macd5mAligned = (FL_5m - SL_5m) <= eps
                    End If
                End If
            Else
                macd5mAligned = False
            End If

            If macd5mAligned Then
                nextStepsText = nextStepsText & "   • MACD 5m: (ALIGNED) -- Momentum supports your position. Maintain or consider adding if other signals confirm." & vbCrLf
            Else
                nextStepsText = nextStepsText & "   • MACD 5m: (NOT ALIGNED) -- MOMENTUM WEAKENING. Consider tightening stop-loss or reducing exposure." & vbCrLf
            End If

            ' MACD 15m
            If haveFL15 And haveSL15 Then
                If haveM15 And haveS15 Then
                    If isBuy Then
                        macd15mAligned = ((FL_15m - SL_15m) >= -eps) And ((MACDLine_15m - SIG_15m) >= -eps)
                    Else
                        macd15mAligned = ((FL_15m - SL_15m) <= eps) And ((MACDLine_15m - SIG_15m) <= eps)
                    End If
                Else
                    If isBuy Then
                        macd15mAligned = (FL_15m - SL_15m) >= -eps
                    Else
                        macd15mAligned = (FL_15m - SL_15m) <= eps
                    End If
                End If
            Else
                macd15mAligned = False
            End If

            If macd15mAligned Then
                nextStepsText = nextStepsText & "   • MACD 15m: (ALIGNED) -- Higher-timeframe momentum supports your position. Positive confirmation." & vbCrLf
            Else
                nextStepsText = nextStepsText & "   • MACD 15m: (NOT ALIGNED) -- Higher-timeframe momentum does not confirm. Monitor for reversal." & vbCrLf
            End If

            ' Stochastic 5m: concise output
            If (isBuy And stochK > stochD) Or (Not isBuy And stochK < stochD) Then
                nextStepsText = nextStepsText & "   • Stochastic 5m: (ALIGNED) (%K " & IIf(isBuy, ">", "<") & " %D) -- MOMENTUM STRENGTHENING. Consider holding or adding." & vbCrLf
            Else
                nextStepsText = nextStepsText & "   • Stochastic 5m: (NOT ALIGNED) (%K " & IIf(isBuy, "<", ">") & " %D) -- MOMENTUM WEAKENING. Exercise caution; consider tightening stops." & vbCrLf
            End If

            '--- Plan ---
            'nextStepsText = nextStepsText & vbCrLf & " PLAN" & vbCrLf
            nextStepsText = nextStepsText & vbCrLf & IIf(isBuy, _
                "   BREAK ABOVE " & r1 & " -- Confirmation(ADD / HOLD); " & _
                "   STALL BELOW " & pp & " -- Weakness (TRIM); " & _
                "   FALL BELOW " & slPrice & " -- SL breached (EXIT); ", _
                "   BREAK BELOW " & s1 & " -- Stay Short; " & _
                "   STALL ABOVE " & pp & " -- Reversal Risk; " & _
                "   RISE ABOVE " & slPrice & " -- SL breached (EXIT)")

            ' --- Sell SL Fix (conservative flip if SL illogical) ---
            If tradeDirection = "sell" And slPrice <= entryPrice Then
                ' If slPrice is below entry for a sell, flip it symmetrically above entry
                slPrice = entryPrice + (entryPrice - slPrice)
            End If

            '--- Stop Loss and Risk ---
            Dim riskPercent As Double
            riskPercent = 0
            On Error Resume Next
            If IsNumeric(entryPrice) And entryPrice <> 0 Then
                riskPercent = Abs((slPrice - entryPrice) / entryPrice) * 100
            End If
            On Error GoTo 0

            nextStepsText = nextStepsText & vbCrLf & _
                "- Maintain stop-loss at " & Format(slPrice, "0.00") & " (~" & Format(riskPercent, "0.00") & "% risk relative to entry)." & _
                " Adjust Qty based on your risk tolerance." & vbCrLf

            '--- Volume Info ---
            nextStepsText = nextStepsText & _
                "- Volume spike threshold: volume > 120% of recent 10-period average." & " " & _
                " Current volume is " & IIf(CurrentVolume > 1.2 * avgVolume10, "above", "below") & " this threshold." & vbCrLf

            '--- Confidence Level ---
            Dim confidenceLabelText As String
            Dim confidenceExplanation As String
                        ' [V2] confidenceScore computed from tech + ML alignment
            confidenceScore = trendSupport
            If (isBuy And CurrentPrice >= vwap) Or (Not isBuy And CurrentPrice <= vwap) Then confidenceScore = confidenceScore + 1
            If (isBuy And stochK > stochD) Or (Not isBuy And stochK < stochD) Then confidenceScore = confidenceScore + 1
            Dim mlAligned As Boolean: mlAligned = (mlSignal = "BUY_NOW" And isBuy) Or (mlSignal = "SELL_NOW" And Not isBuy)
            If mlAligned Then confidenceScore = confidenceScore + 2
            If confidenceScore >= 6 Then
                confidenceLabelText = "HIGH"
                confidenceExplanation = "(Strong support from indicators; favorable conditions.)"
            ElseIf confidenceScore >= 4 Then
                confidenceLabelText = "MEDIUM"
                confidenceExplanation = "(Moderate support; monitor momentum closely.)"
            Else
                confidenceLabelText = "LOW"
                confidenceExplanation = "(Limited support; higher risk of false signals.)"
            End If

            nextStepsText = nextStepsText & vbCrLf & _
                            "  CONFIDENCE LEVEL: " & confidenceLabelText & " " & confidenceExplanation & vbCrLf & _
                            "  Target: " & Format(IIf(isBuy, r2, s2), "0.00") & " ; " & _
                            "  SL: " & Format(slPrice, "0.00") & vbCrLf & _
                            "  [" & niftyRegime & "] " & IIf(mlAligned, "ML: ALIGNED", IIf(mlSignal <> "", "ML: DIVERGES (" & mlSignal & ")", "ML: no signal")) & " | Held " & barsHeld & " bars | PnL " & Format(IIf(entryPrice > 0, IIf(isBuy, 1, -1) * (CurrentPrice - entryPrice) / entryPrice * 100, 0), "0.00") & "% | Conf " & Format(mlConf, "0.00")

            ' Write to cell
            With cellNextSteps
                .ClearContents
                .Font.Color = RGB(0, 0, 0)
                .value = RTrim(nextStepsText)
                .WrapText = True
                .VerticalAlignment = xlTop
                .Characters.Font.size = 10
                '.Rows.AutoFit
            End With

            ' --- Keyword Highlight (optimized Len per word) ---
            Dim keywordMap1 As Variant, wLen As Long
            keywordMap1 = Array( _
                          Array("WEAK TRADE", COLOR_RED), Array("STRONG TRADE", COLOR_GREEN), Array("MOMENTUM", COLOR_BLUE), Array("PLAN", COLOR_BLUE), Array("CONFIDENCE LEVEL", COLOR_BLUE), Array("MEDIUM", COLOR_YELLOW), _
                          Array("WEAK", COLOR_RED), Array("STRONG", COLOR_GREEN), Array("INDECISIVE", COLOR_YELLOW), Array("AGAINST", COLOR_RED), Array("FOR", COLOR_GREEN), Array("Reversal Risk", COLOR_YELLOW), _
                          Array("ALIGNED", COLOR_GREEN), Array("TRIM", COLOR_YELLOW), Array("Opposite Position", COLOR_RED), Array("Weakness", COLOR_YELLOW), Array("Confirmation", COLOR_GREEN), _
                          Array("BREAK BELOW", COLOR_BLUE), Array("BREAK ABOVE", COLOR_BLUE), Array("STALL BELOW", COLOR_BLUE), Array("FALL BELOW", COLOR_BLUE), Array("Target", COLOR_GREEN), Array("SL breached", COLOR_RED), _
                          Array("SL", COLOR_RED), Array("LOW", COLOR_RED), Array("HIGH", COLOR_GREEN), Array("EXIT", COLOR_RED), Array("VWAP", COLOR_BLUE), Array("ISSUES", COLOR_RED), _
                          Array("ADD", COLOR_GREEN), Array("HOLD", COLOR_GREEN), Array("Volume", COLOR_BLUE), Array("Spike", COLOR_GREEN), Array("No Spike", COLOR_RED), Array("Stay Short", COLOR_GREEN), _
                          Array("BREAK BELOW", COLOR_BLUE), Array("STALL ABOVE", COLOR_BLUE), Array("RISE ABOVE", COLOR_BLUE), Array("SUPPORT", COLOR_GREEN), Array("Stochastic", COLOR_BLUE), _
                          Array("Supertrend", COLOR_BLUE), Array("EMA 10/50 5m", COLOR_BLUE), Array("EMA 10/50 15m", COLOR_BLUE), Array("EMA 10/30 5m", COLOR_BLUE), Array("EMA 10/30 15m", COLOR_BLUE), _
                          Array("Bullish", COLOR_GREEN), Array("BEARISH", COLOR_RED), Array("MIXED", COLOR_YELLOW), Array("NOT ALIGNED", COLOR_RED), Array("MOMENTUM WEAKENING", COLOR_RED), Array("MOMENTUM STRENGTHENING", COLOR_GREEN), _
                          Array("EMA 5m/15m", COLOR_BLUE), Array("MACD", COLOR_BLUE), Array("MACD 15m", COLOR_BLUE), Array("MACD 5m", COLOR_BLUE), Array("EMA 10/30 15m", COLOR_BLUE))

            For p = 0 To UBound(keywordMap1)
                word = keywordMap1(p)(0)
                colorVal = keywordMap1(p)(1)
                wLen = Len(word)
                pos = 1
                Do
                    pos = InStr(pos, nextStepsText, word, vbBinaryCompare)
                    If pos = 0 Then Exit Do
                    On Error Resume Next
                    cellNextSteps.Characters(pos, wLen).Font.Color = colorVal
                    cellNextSteps.Characters(pos, wLen).Font.Bold = True
                    On Error GoTo 0
                    pos = pos + wLen
                Loop
            Next p
            'cellNextSteps.Rows.AutoFit
        End If

SkipRow:
    Next i

    Application.ScreenUpdating = True
End Sub

Sub GeneratePreMarketIdeas()
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim startRow As Long, endRow As Long, rowCount As Long
    Dim i As Long
    Dim data As Variant, output() As Variant
    Dim ticker As String, action As String, reason As String
    Dim confidence As Integer, coreScore As Integer
    Dim openP As Double, highP As Double, lowP As Double, closeP As Double, volumeP As Double
    Dim ema10 As Double, ema30 As Double, rsi1d As Double, superTrend As Double
    Dim ema10_1h As Double, ema30_1h As Double, rsi1h As Double
    Dim fib38 As Double, fib50 As Double, fib61 As Double, bbLower As Double, bbUpper As Double, stochK As Double, stochD As Double
    Dim entryPrice As Double, exitPrice As Double, slPrice As Double
    Dim macd_15m As Double, macdSignal_15m As Double
    Dim bullishCandle As Boolean, bearishCandle As Boolean
    Dim VolumeSpike As Integer
    Dim s3 As Double, s2 As Double, s1 As Double, pp As Double, r1 As Double, r2 As Double, r3 As Double
    Dim flags As String
    
    ' Color constants for formatting
    Dim COLOR_RED As Long: COLOR_RED = RGB(200, 0, 0)
    Dim COLOR_GREEN As Long: COLOR_GREEN = RGB(0, 160, 0)
    Dim COLOR_BLUE As Long: COLOR_BLUE = RGB(0, 0, 255)
    Dim COLOR_YELLOW As Long: COLOR_YELLOW = RGB(200, 140, 0)
    
    ' Keyword-color mapping
    Dim keywordMap As Variant

    
    keywordMap = Array( _
        Array("BULLISH", COLOR_GREEN), Array("BEARISH", COLOR_RED), Array("TREND", COLOR_YELLOW), Array("REVERSAL", COLOR_BLUE), Array("EMA", COLOR_BLUE), Array("MACD", COLOR_BLUE), _
        Array("crossover", COLOR_GREEN), Array("crossunder", COLOR_RED), Array("Close near high", COLOR_GREEN), Array("Weak close", COLOR_RED), Array("Price > Supertrend", COLOR_GREEN), _
        Array("Bearish Candle", COLOR_RED), Array("Bullish Candle", COLOR_GREEN), Array("RSI", COLOR_BLUE), Array("1D", COLOR_BLUE), Array("1H", COLOR_BLUE), _
        Array("10 > 30", COLOR_GREEN), Array("10 < 30", COLOR_RED), Array("Price < Supertrend", COLOR_RED) _
    )
    
    ' Setup
    Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Set wsDest = ThisWorkbook.Sheets("PreMarketIdeas")
    
    wsDest.Range("A2:G" & wsDest.Rows.count).ClearContents
    wsDest.Range("A2:G" & wsDest.Rows.count).Interior.ColorIndex = 0
    wsDest.cells(1, 1).Resize(1, 7).value = Array("Ticker", "Action", "Entry", "Exit", "SL", "Confidence", "Reason")
    
    startRow = wsSrc.Range("P1").value
    endRow = wsSrc.Range("R1").value
    rowCount = endRow - startRow + 1
    
    data = wsSrc.Range("A" & startRow & ":DD" & endRow).value
    ReDim output(1 To rowCount, 1 To 7)
    
    ' Main loop
    For i = 1 To rowCount
        ticker = Trim(data(i, 1))
        If ticker = "" Then GoTo NextRow
        
        openP = val(data(i, 3))
        highP = val(data(i, 4))
        lowP = val(data(i, 5))
        closeP = val(data(i, 6))
        volumeP = val(data(i, 7))
        
        If closeP = 0 Or highP = 0 Or lowP = 0 Then
            action = "Ignore": reason = "Missing OHLC data"
            GoTo writeRow
        End If
        
        rsi1d = val(data(i, 24))
        ema10 = val(data(i, 60))
        ema30 = val(data(i, 61))
        superTrend = val(data(i, 72))
        stochK = val(data(i, 83))
        stochD = val(data(i, 84))
        bbLower = val(data(i, 87))
        bbUpper = val(data(i, 89))
        VolumeSpike = val(data(i, 71))
        fib38 = val(data(i, 94))
        fib50 = val(data(i, 95))
        fib61 = val(data(i, 96))
        
        ema10_1h = val(data(i, 55))
        ema30_1h = val(data(i, 56))
        rsi1h = val(data(i, 23))
        
        macd_15m = Round(val(data(i, 106)) - val(data(i, 107)), 2)
        macdSignal_15m = val(data(i, 108))
        
        bullishCandle = (data(i, 8) <> "")
        bearishCandle = (data(i, 9) <> "")
        s3 = val(data(i, 11))
        s2 = val(data(i, 12))
        s1 = val(data(i, 13))
        pp = val(data(i, 14))
        r1 = val(data(i, 15))
        r2 = val(data(i, 16))
        r3 = val(data(i, 17))
        
        action = "Ignore": confidence = 0: coreScore = 0: reason = "": flags = ""
        
        ' === Bullish Setup ===
        If closeP >= lowP + 0.75 * (highP - lowP) Then flags = flags & "Close near high || ": coreScore = coreScore + 1 Else flags = flags & "Weak close || "
        If ema10 > ema30 Then flags = flags & "EMA 10 > 30 (1D) || ": coreScore = coreScore + 1 Else flags = flags & "EMA(1D) 10 < 30 || "
        If macd_15m > macdSignal_15m Then flags = flags & "MACD crossover (15m) || ": coreScore = coreScore + 1 Else flags = flags & "MACD flat/down || "
        If rsi1d >= 50 And rsi1d <= 65 Then flags = flags & "RSI  within 50–65 (1D) || ": coreScore = coreScore + 1 Else flags = flags & "RSI not between 50 and 65 range  || "
        If closeP > superTrend Then flags = flags & "Price > Supertrend || ": coreScore = coreScore + 1 Else flags = flags & "Price < Supertrend || "
        If bullishCandle Then flags = flags & "Bullish Candle (" & data(i, 8) & ") || ": coreScore = coreScore + 1 Else flags = flags & "No Bullish candle || "
        If ema10_1h > ema30_1h Then flags = flags & "EMA 10 > 30 (1H) || ": confidence = confidence + 1 Else flags = flags & "EMA(1H) 10 < 30 || "
        If rsi1h > 50 Then flags = flags & "1H RSI > 50 || ": confidence = confidence + 1 Else flags = flags & "1H RSI <= 50 || "
        
        If coreScore >= 4 Then
            action = "Buy"
            entryPrice = IIf(closeP <= ema10, closeP, ema10)
            
            If closeP >= s3 And closeP < s2 Then exitPrice = s2: slPrice = s3
            If closeP >= s2 And closeP < s1 Then exitPrice = s1: slPrice = s2
            If closeP >= s1 And closeP < pp Then exitPrice = pp: slPrice = s1
            If closeP >= pp And closeP < r1 Then exitPrice = r1: slPrice = pp
            If closeP >= r1 And closeP < r2 Then exitPrice = r2: slPrice = r1
            If closeP >= r2 And closeP <= r3 Then exitPrice = r3: slPrice = r2
            
            confidence = confidence + coreScore
            reason = "BULLISH " & flags
'        Else
'            reason = flags
        End If
        
        ' === Bearish Setup ===
        If action = "Ignore" Then
            flags = "": coreScore = 0
            If closeP <= lowP + 0.25 * (highP - lowP) Then flags = flags & "Close near low || ": coreScore = coreScore + 1 Else flags = flags & "Weak close || "
            If ema10 < ema30 Then flags = flags & "EMA10<30 (1D) || ": coreScore = coreScore + 1 Else flags = flags & "EMA(1D) 10 > 30 || "
            If macd_15m < macdSignal_15m Then flags = flags & "MACD crossunder (15m) || ": coreScore = coreScore + 1 Else flags = flags & "MACD flat/up || "
            If rsi1d >= 35 And rsi1d <= 50 Then flags = flags & "RSI within 35–50 (1D) || ": coreScore = coreScore + 1 Else flags = flags & "RSI not between 35 and 50 || "
            If closeP < superTrend Then flags = flags & "Price < Supertrend || ": coreScore = coreScore + 1 Else flags = flags & "Price > Supertrend || "
            If bearishCandle Then flags = flags & "Bearish Candle (" & data(i, 9) & ") || ": coreScore = coreScore + 1 Else flags = flags & "No Bearish candle || "
            If ema10_1h < ema30_1h Then flags = flags & "EMA10<30 (1H) || ": confidence = confidence + 1 Else flags = flags & "EMA(1H) 10 > 30 || "
            If rsi1h < 50 Then flags = flags & "1H RSI < 50 || ": confidence = confidence + 1 Else flags = flags & "1H RSI >= 50 || "
            
            If coreScore >= 4 Then
                action = "Sell"
                entryPrice = IIf(closeP > ema10, closeP, ema10)
                
                If closeP <= r3 And closeP > r2 Then exitPrice = r2: slPrice = r3
                If closeP <= r2 And closeP > r1 Then exitPrice = r1: slPrice = r2
                If closeP <= r1 And closeP > pp Then exitPrice = pp: slPrice = r1
                If closeP <= pp And closeP > s1 Then exitPrice = s1: slPrice = pp
                If closeP <= s1 And closeP > s2 Then exitPrice = s2: slPrice = s1
                If closeP <= s2 And closeP >= s3 Then exitPrice = s3: slPrice = s2
                
                confidence = confidence + coreScore
                reason = "BEARISH " & flags
            End If
'        Else
'            reason = flags
        End If
        
        ' Tag setup type
        If action = "Buy" And (stochK > stochD And stochK < 40 Or closeP <= bbLower Or VolumeSpike = 1) Then reason = "REVERSAL : " & reason
        If action = "Sell" And (stochK < stochD And stochK > 60 Or closeP >= bbUpper Or VolumeSpike = 1) Then reason = "REVERSAL : " & reason
        If (action = "Buy" Or action = "Sell") And confidence >= 7 Then reason = "TREND : " & reason
        
        ' Remove trailing "||"
        If Right(Trim(reason), 2) = "||" Then reason = Left(Trim(reason), Len(Trim(reason)) - 2)
        
writeRow:
        output(i, 1) = ticker
        output(i, 2) = action
        output(i, 3) = IIf(action <> "Ignore", Round(entryPrice, 2), "")
        output(i, 4) = IIf(action <> "Ignore", Round(exitPrice, 2), "")
        output(i, 5) = IIf(action <> "Ignore", Round(slPrice, 2), "")
        output(i, 6) = IIf(action <> "Ignore", confidence, "")
'        output(i, 3) = Round(entryPrice, 2)
'        output(i, 4) = Round(exitPrice, 2)
'        output(i, 5) = Round(slPrice, 2)
'        output(i, 6) = confidence
        output(i, 7) = reason
NextRow:
    Next i
    
    ' Write all results at once
    wsDest.Range("A2").Resize(rowCount, 7).value = output
    
    ' Apply keyword formatting
    Dim celllatest As Range, p As Long, word As String, colorVal As Long, pos As Long
    For i = 2 To rowCount + 1
        Set celllatest = wsDest.cells(i, "G")
        If celllatest.value <> "" Then
            celllatest.ClearFormats
            For p = 0 To UBound(keywordMap)
                word = keywordMap(p)(0)
                colorVal = keywordMap(p)(1)
                pos = 1
                Do
                    pos = InStr(pos, celllatest.value, word, vbBinaryCompare)
                    If pos = 0 Then Exit Do
                    On Error Resume Next
                    celllatest.Characters(pos, Len(word)).Font.Color = colorVal
                    celllatest.Characters(pos, Len(word)).Font.Bold = True
                    On Error GoTo 0
                    pos = pos + Len(word)
                Loop
            Next p
            ' Action column background
            With wsDest.cells(i, 2)
                Select Case .value
                    Case "Buy": .Interior.Color = RGB(200, 255, 200)
                    Case "Sell": .Interior.Color = RGB(255, 200, 200)
            End Select
            End With
        End If
    Next i
    
    MsgBox "Pre-market ideas written to 'PreMarketIdeas'.", vbInformation
End Sub



Sub FetchNewsFromGoogleRSS()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual


    Dim wsSrc As Worksheet, wsOut As Worksheet
    Set wsSrc = ThisWorkbook.Sheets("Sheet1")
    Set wsOut = ThisWorkbook.Sheets("StockNews")

    Dim startRow As Long, endRow As Long
    startRow = wsSrc.Range("P1").value
    endRow = wsSrc.Range("R1").value

    Dim ticker As String, i As Long
    Dim xml As Object, node As Object, item As Object
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    Dim doc As Object: Set doc = CreateObject("MSXML2.DOMDocument")

    ' Clear old results (except headers in row 1)
    wsOut.Range("A2:F" & wsOut.Rows.count).ClearContents
    wsOut.Range("A2:F" & wsOut.Rows.count).Interior.ColorIndex = 0

    Dim outArr() As Variant
    Dim rowIndex As Long: rowIndex = 1
    ReDim outArr(1 To (endRow - startRow + 1) * 3, 1 To 5) ' Up to 3 headlines per ticker

    For i = startRow To endRow
        ticker = Trim(wsSrc.cells(i, 1).value)
        If ticker = "" Then GoTo NextTicker
        
        Dim searchTerm As String

        Select Case LCase(ticker)
            Case "nifty_50": searchTerm = "Nifty"
            Case "nifty_bank": searchTerm = "Nifty Bank"
            Case "nifty_it": searchTerm = "Nifty IT"
            Case "nifty_pharma": searchTerm = "Nifty Pharma"
            Case "nifty_auto": searchTerm = "Nifty Auto"
            Case "nifty_fmcg": searchTerm = "Nifty FMCG"
            Case Else: searchTerm = ticker
        End Select


        Dim rssUrl As String
        rssUrl = "https://news.google.com/rss/search?q=" & searchTerm & "+India&hl=en-IN&gl=IN&ceid=IN:en"

        http.Open "GET", rssUrl, False
        'Application.StatusBar = "Fetching news for: " & ticker & " (" & i & " of " & endRow & ")"

        http.Send

        If http.status = 200 Then
            doc.LoadXML http.ResponseText

            Dim count As Long: count = 0
            For Each item In doc.getElementsByTagName("item")
                If count >= 5 Then Exit For ' Limit to 3 headlines
                Dim title As String, pubDate As String, link As String, source As String
                title = item.SelectSingleNode("title").Text
                
                Dim cleanDate As String, dt As Date
                pubDate = item.SelectSingleNode("pubDate").Text
                
                ' Clean the date string: remove day name and "GMT"
                If InStr(pubDate, ",") > 0 Then
                    cleanDate = Trim(Mid(pubDate, InStr(pubDate, ",") + 1))
                Else
                    cleanDate = pubDate
                End If
                
                cleanDate = Replace(cleanDate, "GMT", "")
                cleanDate = Trim(cleanDate)
                
                On Error Resume Next
                dt = CDate(cleanDate)
                On Error GoTo 0

                If dt < Date - 3 Then GoTo SkipItem
                link = item.SelectSingleNode("link").Text
                source = item.SelectSingleNode("source").Text

                ' Write to array
                outArr(rowIndex, 1) = ticker
                outArr(rowIndex, 2) = pubDate
                outArr(rowIndex, 3) = title
                outArr(rowIndex, 4) = source
                outArr(rowIndex, 5) = link
                rowIndex = rowIndex + 1
                count = count + 1
SkipItem:
            Next item
        Else
            Debug.Print "Failed to fetch for " & ticker
        End If

NextTicker:
    If (i - startRow + 1) Mod 10 = 0 Then DoEvents
    Next i

    ' Trim unused rows
    If rowIndex > 1 Then
        wsOut.Range("A2").Resize(rowIndex - 1, 5).value = outArr
        ScoreHeadlines
    Else
        MsgBox "No news found.", vbInformation
    End If

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic


    MsgBox "News fetched for " & (endRow - startRow + 1) & " tickers.", vbInformation
End Sub

Sub ScoreHeadlines()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("StockNews")
    Dim lastRow As Long: lastRow = ws.cells(ws.Rows.count, "A").End(xlUp).row

    Dim cell As Range, score As String
    Dim headline As String
    Dim i As Long

    For i = 2 To lastRow
        headline = LCase(ws.cells(i, 3).value)
        
        score = "Neutral"
        If headline Like "*loss*" Or headline Like "*fall*" Or headline Like "*slump*" Or _
           headline Like "*downgrade*" Or headline Like "*miss*" Or headline Like "*exit*" Then
            score = "Bearish"
            ws.Range("A" & i & ":E" & i).Interior.Color = RGB(255, 204, 204) ' Light Red
        ElseIf headline Like "*profit*" Or headline Like "*beats*" Or headline Like "*surge*" Or _
               headline Like "*buy*" Or headline Like "*record*" Or headline Like "*upgraded*" Then
            score = "Bullish"
            ws.Range("A" & i & ":E" & i).Interior.Color = RGB(198, 239, 206) ' Light Green
        Else
            'ws.Range("A" & i & ":E" & i).Interior.Color = RGB(255, 255, 153) ' Light Yellow
        End If

        ' Add a 6th column if you want to show the sentiment tag
        ws.cells(i, 6).value = score
    Next i

    'MsgBox "Sentiment scoring applied!", vbInformation
End Sub

' ? MARKET CONTEXT MACRO WITHOUT SHEET CREATION OR FILTERING
Sub FetchMarketContextNewsWithSentiment()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("MarketContext")

    ' Clear old data (leave headers in Row 1)
    
    ws.Range("A2:F500").ClearContents
    ws.Range("A2:F500").Interior.ColorIndex = 0
    
'    ws.Range("A2:F" & ws.Cells(ws.Rows.count, "B").End(xlUp).row).Interior.ColorIndex = 0
'    ws.Range("A2:F" & ws.Cells(ws.Rows.count, "A").End(xlUp).row).ClearContents

    Dim topics As Variant
    topics = Array( _
        "Nifty 50 India", "Bank Nifty", "RBI India", "Inflation India", "Op Sindoor", "India GDP", "India Elections", _
        "Tariff", "Dow Jones", "Nasdaq", "Fed rate", "US inflation", "US treasury yield", _
        "ECB rate", "Eurozone inflation", "FTSE 100", "DAX", _
        "Nikkei 225", "Hang Seng", "China stimulus", "BOJ Japan", _
        "Russia Ukraine", "Oil Prices", "Gold", "USD INR" _
    )

    Dim http As Object, xml As Object
    Dim rssUrl As String, topic As Variant
    Dim row As Long: row = 2
    Dim maxItemsPerTopic As Integer: maxItemsPerTopic = 4

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Fetching market context..."

    Dim sentimentBullish As Variant, sentimentBearish As Variant, riskWords As Variant
    sentimentBullish = Array("rises", "beats", "surges", "strong", "positive", "expands", "growth")
    sentimentBearish = Array("falls", "cuts", "misses", "downturn", "loss", "warns", "decline")
    riskWords = Array("war", "conflict", "crisis", "inflation", "recession", "geopolitical", "hike", "slowdown")

    Dim riskHeadlines As Collection: Set riskHeadlines = New Collection

    For Each topic In topics
        rssUrl = "https://news.google.com/rss/search?q=" & Replace(topic, " ", "+") & "&hl=en-IN&gl=IN&ceid=IN:en"
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", rssUrl, False
        http.Send

        If http.status = 200 Then
            Set xml = CreateObject("MSXML2.DOMDocument")
            xml.LoadXML http.ResponseText
            Dim items As Object
            Set items = xml.getElementsByTagName("item")

            Dim i As Integer
            For i = 0 To WorksheetFunction.Min(items.Length - 1, maxItemsPerTopic - 1)
                With items.item(i)
                    Dim title As String, pubDate As String, link As String, dt As Date
                    title = .SelectSingleNode("title").Text
                    pubDate = .SelectSingleNode("pubDate").Text
                    link = .SelectSingleNode("link").Text

                    Dim cleanDate As String
                    cleanDate = Trim(Mid(pubDate, InStr(pubDate, ",") + 1))
                    cleanDate = Replace(cleanDate, "GMT", "")
                    On Error Resume Next: dt = CDate(cleanDate): On Error GoTo 0


                    On Error Resume Next
                    dt = CDate(cleanDate)
                    On Error GoTo 0
    
                    If dt < Date - 7 Then GoTo SkipItem

                    ' Sentiment detection
                    Dim sentiment As String: sentiment = "Neutral"
                    Dim word As Variant
                    For Each word In sentimentBullish
                        If InStr(1, LCase(title), word) > 0 Then sentiment = "Bullish": Exit For
                    Next word
                    If sentiment = "Neutral" Then
                        For Each word In sentimentBearish
                            If InStr(1, LCase(title), word) > 0 Then sentiment = "Bearish": Exit For
                        Next word
                    End If

                    ' Source from link
                    Dim source As String: source = "Unknown"
                    Dim matches As Object, re As Object
                    Set re = CreateObject("VBScript.RegExp")
                    re.Pattern = "https?:\/\/(?:www\.)?([^\/]+)"
                    re.IgnoreCase = True
                    If re.Test(link) Then
                        Set matches = re.Execute(link)
                        source = matches(0).SubMatches(0)
                    End If

                    ' Risk highlight
                    For Each word In riskWords
                        If InStr(1, LCase(title), word) > 0 Then
                            riskHeadlines.Add topic & ": " & title & " | " & link
                            Exit For
                        End If
                    Next word

                    ' Write
                    ws.cells(row, 1).value = topic
                    ws.cells(row, 2).value = title
                    ws.cells(row, 3).value = source
                    ws.cells(row, 4).value = dt
                    ws.cells(row, 5).value = sentiment
                    ws.cells(row, 6).value = link
                    ws.Hyperlinks.Add Anchor:=ws.cells(row, 6), Address:=link, TextToDisplay:="Read"

                    Select Case sentiment
                        Case "Bullish": ws.Range("A" & row & ":F" & row).Interior.Color = RGB(204, 255, 204)
                        Case "Bearish": ws.Range("A" & row & ":F" & row).Interior.Color = RGB(255, 204, 204)
                    End Select
                    row = row + 1
                End With
SkipItem:
            Next i
        End If
        DoEvents
    Next topic

    ' Risk summary
    If riskHeadlines.count > 0 Then
        row = row + 1
        ws.cells(row, 1).value = "SUMMARY: RISK HEADLINES"
        ws.cells(row, 1).Font.Bold = True
        ws.cells(row, 1).Interior.Color = RGB(255, 230, 204)
        row = row + 1
        Dim h As Variant
        For Each h In riskHeadlines
            ws.cells(row, 2).value = Split(h, " | ")(0)
            ws.cells(row, 6).value = Split(h, " | ")(1)
            ws.Hyperlinks.Add Anchor:=ws.cells(row, 6), Address:=ws.cells(row, 6).value, TextToDisplay:="Read"
            row = row + 1
        Next h
    End If

    With ws
        .Columns("A:F").AutoFit
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(200, 200, 200)
        .Rows("1:1").WrapText = True
    End With

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "News with sentiment fetched.", vbInformation
End Sub





