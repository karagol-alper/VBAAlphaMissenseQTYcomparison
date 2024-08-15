Attribute VB_Name = "Module1"
Sub AnalyzeQTYVariants()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim qtyVariants As Collection
    Dim revQtyVariants As Collection
    Dim otherPolarToNonpolar As Collection
    Dim otherNonpolarToPolar As Collection
    
    ' Set the data worksheet
    Set dataWs = ActiveSheet
    
    ' Create a new worksheet for results
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "QTY Analysis"
    
    ' Find the last row of data
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize collections
    Set qtyVariants = New Collection
    Set revQtyVariants = New Collection
    Set otherPolarToNonpolar = New Collection
    Set otherNonpolarToPolar = New Collection
    
    ' Analyze variants
    For i = 2 To lastRow ' Assuming row 1 is header
        Dim aa1 As String, aa2 As String, score As Double
        aa1 = dataWs.Cells(i, 7).Value ' Column G: a.a.1
        aa2 = dataWs.Cells(i, 9).Value ' Column I: a.a.2
        score = dataWs.Cells(i, 10).Value ' Column J: pathogenicity score
        
        ' Identify and categorize variants
        Select Case aa1
            Case "L"
                If aa2 = "Q" Then
                    qtyVariants.Add score
                ElseIf InStr("DERKHNSTY", aa2) > 0 Then
                    otherNonpolarToPolar.Add score
                End If
            Case "I"
                If aa2 = "T" Then
                    qtyVariants.Add score
                ElseIf InStr("DERKHNSQY", aa2) > 0 Then
                    otherNonpolarToPolar.Add score
                End If
            Case "F"
                If aa2 = "Y" Then
                    qtyVariants.Add score
                ElseIf InStr("DERKHNSQT", aa2) > 0 Then
                    otherNonpolarToPolar.Add score
                End If
            Case "Q"
                If aa2 = "L" Then
                    revQtyVariants.Add score
                ElseIf InStr("ACGIMFPWV", aa2) > 0 Then
                    otherPolarToNonpolar.Add score
                End If
            Case "T"
                If aa2 = "I" Then
                    revQtyVariants.Add score
                ElseIf InStr("ACGLMFPWV", aa2) > 0 Then
                    otherPolarToNonpolar.Add score
                End If
            Case "Y"
                If aa2 = "F" Then
                    revQtyVariants.Add score
                ElseIf InStr("ACGILMPWV", aa2) > 0 Then
                    otherPolarToNonpolar.Add score
                End If
        End Select
    Next i
    
    ' Create scatter plot
    CreateScatterPlot ws, qtyVariants, revQtyVariants, otherPolarToNonpolar, otherNonpolarToPolar
    
    ' Calculate and display average scores
    DisplayAverageScores ws, qtyVariants, revQtyVariants, otherPolarToNonpolar, otherNonpolarToPolar
    
    ' Highlight significant differences
    HighlightSignificantDifferences ws, qtyVariants, revQtyVariants, otherPolarToNonpolar, otherNonpolarToPolar

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Sub CreateScatterPlot(ws As Worksheet, qtyVariants As Collection, revQtyVariants As Collection, otherPolarToNonpolar As Collection, otherNonpolarToPolar As Collection)
    On Error GoTo ErrorHandler
    
    Dim cht As ChartObject
    Dim rng As Range
    Dim i As Long
    
    ' Prepare data for plotting
    ws.Cells(1, 1).Value = "Category"
    ws.Cells(1, 2).Value = "Score"
    
    i = 2
    For Each score In qtyVariants
        ws.Cells(i, 1).Value = "QTY Variants"
        ws.Cells(i, 2).Value = score
        i = i + 1
    Next score
    
    For Each score In revQtyVariants
        ws.Cells(i, 1).Value = "Reverse QTY Variants"
        ws.Cells(i, 2).Value = score
        i = i + 1
    Next score
    
    For Each score In otherPolarToNonpolar
        ws.Cells(i, 1).Value = "Other Polar to Nonpolar"
        ws.Cells(i, 2).Value = score
        i = i + 1
    Next score
    
    For Each score In otherNonpolarToPolar
        ws.Cells(i, 1).Value = "Other Nonpolar to Polar"
        ws.Cells(i, 2).Value = score
        i = i + 1
    Next score
    
    ' Create scatter plot
    Set rng = ws.Range("A1:B" & i - 1)
    Set cht = ws.ChartObjects.Add(Left:=300, Width:=450, Top:=10, Height:=250)
    
    With cht.Chart
        .ChartType = xlXYScatter
        .SetSourceData Source:=rng
        .HasTitle = True
        .ChartTitle.Text = "Pathogenicity Scores by Variant Type"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Variant Type"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Pathogenicity Score"
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in CreateScatterPlot: " & Err.Description
End Sub

Sub DisplayAverageScores(ws As Worksheet, qtyVariants As Collection, revQtyVariants As Collection, otherPolarToNonpolar As Collection, otherNonpolarToPolar As Collection)
    On Error GoTo ErrorHandler
    
    ws.Cells(1, 4).Value = "Category"
    ws.Cells(1, 5).Value = "Average Score"
    
    ws.Cells(2, 4).Value = "QTY Variants"
    ws.Cells(2, 5).Value = IIf(qtyVariants.Count > 0, WorksheetFunction.Average(CollectionToArray(qtyVariants)), "N/A")
    
    ws.Cells(3, 4).Value = "Reverse QTY Variants"
    ws.Cells(3, 5).Value = IIf(revQtyVariants.Count > 0, WorksheetFunction.Average(CollectionToArray(revQtyVariants)), "N/A")
    
    ws.Cells(4, 4).Value = "Other Polar to Nonpolar"
    ws.Cells(4, 5).Value = IIf(otherPolarToNonpolar.Count > 0, WorksheetFunction.Average(CollectionToArray(otherPolarToNonpolar)), "N/A")
    
    ws.Cells(5, 4).Value = "Other Nonpolar to Polar"
    ws.Cells(5, 5).Value = IIf(otherNonpolarToPolar.Count > 0, WorksheetFunction.Average(CollectionToArray(otherNonpolarToPolar)), "N/A")
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in DisplayAverageScores: " & Err.Description
End Sub

Sub HighlightSignificantDifferences(ws As Worksheet, qtyVariants As Collection, revQtyVariants As Collection, otherPolarToNonpolar As Collection, otherNonpolarToPolar As Collection)
    On Error GoTo ErrorHandler
    
    Dim qtyAvg As Double, revQtyAvg As Double, otherPolarToNonpolarAvg As Double, otherNonpolarToPolarAvg As Double
    Dim qtyStdDev As Double, revQtyStdDev As Double, otherPolarToNonpolarStdDev As Double, otherNonpolarToPolarStdDev As Double
    
    ws.Cells(7, 4).Value = "Significant Differences:"
    
    If qtyVariants.Count > 1 And otherPolarToNonpolar.Count > 1 Then
        qtyAvg = WorksheetFunction.Average(CollectionToArray(qtyVariants))
        otherPolarToNonpolarAvg = WorksheetFunction.Average(CollectionToArray(otherPolarToNonpolar))
        qtyStdDev = WorksheetFunction.StDev(CollectionToArray(qtyVariants))
        otherPolarToNonpolarStdDev = WorksheetFunction.StDev(CollectionToArray(otherPolarToNonpolar))
        
        If Abs(qtyAvg - otherPolarToNonpolarAvg) > 2 * (qtyStdDev + otherPolarToNonpolarStdDev) / 2 Then
            ws.Cells(8, 4).Value = "QTY Variants vs Other Polar to Nonpolar"
        End If
    End If
    
    If revQtyVariants.Count > 1 And otherNonpolarToPolar.Count > 1 Then
        revQtyAvg = WorksheetFunction.Average(CollectionToArray(revQtyVariants))
        otherNonpolarToPolarAvg = WorksheetFunction.Average(CollectionToArray(otherNonpolarToPolar))
        revQtyStdDev = WorksheetFunction.StDev(CollectionToArray(revQtyVariants))
        otherNonpolarToPolarStdDev = WorksheetFunction.StDev(CollectionToArray(otherNonpolarToPolar))
        
        If Abs(revQtyAvg - otherNonpolarToPolarAvg) > 2 * (revQtyStdDev + otherNonpolarToPolarStdDev) / 2 Then
            ws.Cells(9, 4).Value = "Reverse QTY Variants vs Other Nonpolar to Polar"
        End If
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in HighlightSignificantDifferences: " & Err.Description
End Sub

Function CollectionToArray(col As Collection) As Variant
    Dim arr() As Variant
    Dim i As Long
    
    If col.Count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = col(i)
    Next i
    
    CollectionToArray = arr
End Function