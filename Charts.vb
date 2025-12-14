Option Explicit

Sub BuildAllCharts()
    Dim wsData As Worksheet, wsCharts As Worksheet
    Dim lastRow As Long, r As Long, chartCount As Long
    Dim sheetName As String
    sheetName = InputBox("Enter SheetName for Charts", "ChartMaker (1404) Majid SharifiTehrani", "My Charts")
    Set wsData = ThisWorkbook.Worksheets("Sheet1")
    Set wsCharts = GetOrCreateSheet(sheetName)
    wsCharts.Cells.Clear
    
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
    r = 1
    chartCount = 0
    
    Do While r <= lastRow
        If HasText(wsData.Cells(r, "A").Value) And InStr(1, wsData.Cells(r, "A").Value, "Fig", vbTextCompare) > 0 Then
            Dim chartTitle As String, yAxisTitle As String
            chartTitle = Trim$(wsData.Cells(r, "A").Value)
            yAxisTitle = Trim$(wsData.Cells(r, "B").Value)
            
            ' Read series count from next row (r+1), column A
            Dim seriesCount As Long
            If r + 1 <= lastRow And IsNumeric(wsData.Cells(r + 1, "A").Value) Then
                seriesCount = CLng(wsData.Cells(r + 1, "A").Value)
            Else
                seriesCount = 1
            End If
            
            Dim series1Name As String, series2Name As String
            series1Name = Trim$(wsData.Cells(r, "C").Value & "")
            series2Name = Trim$(wsData.Cells(r, "D").Value & "")
            
            If seriesCount = 2 Then
                ' === DUAL SERIES CHART ===
                Dim cats() As Variant, vals1() As Double, vals2() As Double
                Dim letters1() As Variant, letters2() As Variant
                Dim sd1() As Double, sd2() As Double
                Dim n As Long: n = 0
                Dim rr As Long: rr = r + 1  ' Start reading data from row after "2"
                
                Dim catB As String
                Dim v1 As Double, v2 As Double, s1 As Double, s2 As Double
                Dim ok1 As Boolean, ok2 As Boolean, okSD1 As Boolean, okSD2 As Boolean
                Dim let1 As String, let2 As String
                
                Do While rr <= lastRow + 20
                    If HasText(wsData.Cells(rr, "A").Value) And InStr(1, wsData.Cells(rr, "A").Value, "Fig", vbTextCompare) > 0 Then Exit Do
                    
                    catB = Trim$(wsData.Cells(rr, "B").Value)
                    
                    If HasText(catB) Then
                        ok1 = TryToDouble(wsData.Cells(rr, "C").Value, v1)
                        ok2 = TryToDouble(wsData.Cells(rr, "F").Value, v2)
                        okSD1 = TryToDouble(wsData.Cells(rr, "E").Value, s1)
                        okSD2 = TryToDouble(wsData.Cells(rr, "H").Value, s2)
                        
                        let1 = Trim$(wsData.Cells(rr, "D").Value & "")
                        let2 = Trim$(wsData.Cells(rr, "G").Value & "")
                        
                        If ok1 Or ok2 Then
                            n = n + 1
                            ReDim Preserve cats(1 To n), vals1(1 To n), vals2(1 To n)
                            ReDim Preserve sd1(1 To n), sd2(1 To n)
                            ReDim Preserve letters1(1 To n), letters2(1 To n)
                            
                            cats(n) = catB
                            vals1(n) = IIf(ok1, v1, 0#)
                            vals2(n) = IIf(ok2, v2, 0#)
                            sd1(n) = IIf(okSD1, s1, 0#)
                            sd2(n) = IIf(okSD2, s2, 0#)
                            letters1(n) = IIf(HasText(let1), let1, "")
                            letters2(n) = IIf(HasText(let2), let2, "")
                        End If
                    End If
                    rr = rr + 1
                Loop
                
                If n > 0 Then
                    chartCount = chartCount + 1
                    CreateBarChartDual wsCharts, chartTitle, yAxisTitle, "Treatments", _
                                       cats, vals1, vals2, sd1, sd2, _
                                       letters1, letters2, _
                                       series1Name, series2Name, chartCount
                End If
                
                r = rr - 1
                
            Else
                ' === SINGLE SERIES CHART ===
                Dim catsS() As Variant, valsS() As Double, sdS() As Double, lettersS() As Variant
                Dim m As Long: m = 0
                Dim rs As Long: rs = r + 1  ' Start after "1"
                
                Dim catB2 As String
                Dim v As Double, sdVal As Double
                Dim okV As Boolean, okSD As Boolean
                Dim letS As String
                
                Do While rs <= lastRow + 20
                    If HasText(wsData.Cells(rs, "A").Value) And InStr(1, wsData.Cells(rs, "A").Value, "Fig", vbTextCompare) > 0 Then Exit Do
                    
                    catB2 = Trim$(wsData.Cells(rs, "B").Value)
                    
                    If HasText(catB2) Then
                        okV = TryToDouble(wsData.Cells(rs, "C").Value, v)
                        okSD = TryToDouble(wsData.Cells(rs, "E").Value, sdVal)
                        letS = Trim$(wsData.Cells(rs, "D").Value & "")
                        
                        If okV Then
                            m = m + 1
                            ReDim Preserve catsS(1 To m), valsS(1 To m), sdS(1 To m), lettersS(1 To m)
                            
                            catsS(m) = catB2
                            valsS(m) = v
                            sdS(m) = IIf(okSD, sdVal, 0#)
                            lettersS(m) = IIf(HasText(letS), letS, "")
                        End If
                    End If
                    rs = rs + 1
                Loop
                
                If m > 0 Then
                    chartCount = chartCount + 1
                    CreateBarChartSingle wsCharts, chartTitle, yAxisTitle, "Treatments", _
                                         catsS, valsS, sdS, lettersS, chartCount
                End If
                
                r = rs - 1
            End If
        End If
        r = r + 1
    Loop
    
    wsCharts.Activate
    MsgBox chartCount & " Column Charts Added To:  " & sheetName, vbInformation, "ChartMaker (1404) Majid SharifiTehrani"
End Sub

' ====================== HELPERS ======================
Private Function GetOrCreateSheet(name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = name
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function HasText(v As Variant) As Boolean
    HasText = Not IsError(v) And Trim$(CStr(v & "")) <> ""
End Function

' ====================== CREATE SINGLE SERIES CHART ======================
Private Sub CreateBarChartSingle(ws As Worksheet, chartTitle As String, yAxisTitle As String, xAxisTitle As String, ByRef categories() As Variant, ByRef values() As Double, ByRef sd() As Double, ByRef letters() As Variant, chartIndex As Long)
    
    Dim co As ChartObject, ch As Chart
    Dim leftPos As Double, topPos As Double
    CalcChartPosition chartIndex, 3, 380, 260, 20, 20, leftPos, topPos
    
    Set co = ws.ChartObjects.Add(Left:=leftPos, Top:=topPos, Width:=380, Height:=260)
    co.name = "Chart_" & chartIndex
    Set ch = co.Chart
    
    ch.ChartType = xlColumnClustered
    
    With ch.SeriesCollection.NewSeries
        .name = ""
        .XValues = categories
        .values = values
        .Format.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    With ch.SeriesCollection(1)
        .ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlCustom, Amount:=sd, MinusValues:=sd
        .HasDataLabels = True
        .DataLabels.Font.Bold = True
        .DataLabels.Font.Size = 10
    End With
    
    With ch
        .HasTitle = False
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = xAxisTitle
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = yAxisTitle
        .HasLegend = False
        .ChartStyle = 215
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlValue).HasMajorGridlines = False
        .Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Axes(xlValue).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    Dim i As Long
    For i = 1 To UBound(letters)
        With ch.SeriesCollection(1).Points(i).DataLabel
            If Trim$(letters(i) & "") <> "" Then
                .Text = letters(i)
                On Error Resume Next                 ' Temporarily ignore any rare glitch
                .Position = xlLabelPositionAbove     ' Excel places it above the bar top
                .Top = .Top - 6                      ' Small consistent nudge upward
                On Error GoTo 0
            Else
                .Delete                              ' Clean: no label if no letter
            End If
        End With
    Next i
       
End Sub

' ====================== CREATE DUAL SERIES CHART ======================
Private Sub CreateBarChartDual( _
    ws As Worksheet, chartTitle As String, yAxisTitle As String, _
    xAxisTitle As String, ByRef categories() As Variant, _
    ByRef series1() As Double, ByRef series2() As Double, _
    ByRef sd1() As Double, ByRef sd2() As Double, _
    ByRef letters1() As Variant, ByRef letters2() As Variant, _
    series1Name As String, series2Name As String, chartIndex As Long)
    
    Dim co As ChartObject, ch As Chart
    Dim leftPos As Double, topPos As Double
    CalcChartPosition chartIndex, 3, 380, 260, 20, 20, leftPos, topPos
    
    Set co = ws.ChartObjects.Add(Left:=leftPos, Top:=topPos, Width:=380, Height:=260)
    co.name = "Chart_" & chartIndex
    Set ch = co.Chart
    
    ch.ChartType = xlColumnClustered
    
    With ch.SeriesCollection.NewSeries
        .name = "=""" & IIf(HasText(series1Name), series1Name, "Series 1") & """"
        .XValues = categories
        .values = series1
    End With
    
    With ch.SeriesCollection.NewSeries
        .name = "=""" & IIf(HasText(series2Name), series2Name, "Series 2") & """"
        .XValues = categories
        .values = series2
    End With
    
    With ch.SeriesCollection(1)
        .ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlCustom, Amount:=sd1, MinusValues:=sd1
        .HasDataLabels = True
        .DataLabels.Font.Bold = True
        .DataLabels.Font.Size = 10
    End With
    
    With ch.SeriesCollection(2)
        .ErrorBar Direction:=xlY, Include:=xlBoth, Type:=xlCustom, Amount:=sd2, MinusValues:=sd2
        .HasDataLabels = True
        .DataLabels.Font.Bold = True
        .DataLabels.Font.Size = 10
    End With
    
    Dim i As Long
    
    With ch
        .HasTitle = False
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = xAxisTitle
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = yAxisTitle
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
        .ChartStyle = 215
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlValue).HasMajorGridlines = False
        .Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Axes(xlValue).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
        .Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    With ch.ChartGroups(1)
        .Overlap = 0
        .GapWidth = 150
    End With
    
    
    ' Series 1 letters
    For i = 1 To UBound(letters1)
        With ch.SeriesCollection(1).Points(i).DataLabel
            If Trim$(letters1(i) & "") <> "" Then
                .Text = letters1(i)
                On Error Resume Next                 ' Temporarily ignore any rare glitch
                .Position = xlLabelPositionAbove     ' Excel places it above the bar top
                .Top = .Top - 6                      ' Small consistent nudge upward
                On Error GoTo 0
            Else
                .Delete
            End If
        End With
    Next i
    
    ' Series 2 letters
    For i = 1 To UBound(letters2)
        With ch.SeriesCollection(2).Points(i).DataLabel
            If Trim$(letters2(i) & "") <> "" Then
                .Text = letters2(i)
                On Error Resume Next                 ' Temporarily ignore any rare glitch
                .Position = xlLabelPositionAbove     ' Excel places it above the bar top
                .Top = .Top - 6                      ' Small consistent nudge upward
                On Error GoTo 0
            Else
                .Delete
            End If
        End With
    Next i
    
End Sub
Private Function TryToDouble(v As Variant, ByRef d As Double) As Boolean
    If IsError(v) Or IsEmpty(v) Then
        TryToDouble = False
        Exit Function
    End If
    Dim s As String: s = Trim$(CStr(v))
    If s = "" Or Not IsNumeric(s) Then
        TryToDouble = False
    Else
        d = CDbl(s)
        TryToDouble = True
    End If
End Function

Private Sub CalcChartPosition( _
    ByVal chartIndex As Long, ByVal chartsPerRow As Long, _
    ByVal chartWidth As Double, ByVal chartHeight As Double, _
    ByVal padX As Double, ByVal padY As Double, _
    ByRef leftPos As Double, ByRef topPos As Double)
    
    Dim col As Long, row As Long
    col = ((chartIndex - 1) Mod chartsPerRow)
    row = ((chartIndex - 1) \ chartsPerRow)
    leftPos = padX + col * (chartWidth + padX)
    topPos = padY + row * (chartHeight + padY)
End Sub



