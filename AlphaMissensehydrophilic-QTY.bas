Attribute VB_Name = "Module1"
Sub AnalyzeQTYVariants()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim qtyVariants As Collection
    Dim revQtyVariants As Collection
    Dim otherHydroToPhil As Collection
    Dim otherPhilToHydro As Collection
    
    ' Hydrophobic amino acids
    Const HYDROPHOBIC = "AILFWV"
    ' Hydrophilic amino acids (excluding Q, T, Y)
    Const HYDROPHILIC = "RKDENPHSGMC"
    
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
    Set otherHydroToPhil = New Collection
    Set otherPhilToHydro = New Collection
    
    ' Analyze variants
    For i = 2 To lastRow ' Assuming row 1 is header
        Dim aa1 As String, aa2 As String, score As Double
        aa1 = dataWs.Cells(i, 7).Value ' Column G: a.a.1
        aa2 = dataWs.Cells(i, 9).Value ' Column I: a.a.2
        score = dataWs.Cells(i, 10).Value ' Column J: pathogenicity score
        
        ' Identify and categorize variants
        If InStr("QTY", aa1) > 0 And InStr(HYDROPHOBIC, aa2) > 0 Then
            qtyVariants.Add score
        ElseIf InStr(HYDROPHOBIC, aa1) > 0 And InStr("QTY", aa2) > 0 Then
            revQtyVariants.Add score
        ElseIf InStr(HYDROPHILIC, aa1) > 0 And InStr(HYDROPHOBIC, aa2) > 0 Then
            otherPhilToHydro.Add score
        ElseIf InStr(HYDROPHOBIC, aa1) > 0 And InStr(HYDROPHILIC, aa2) > 0 Then
            otherHydroToPhil.Add score
        End If
    Next i
    
    ' Create scatter plot
    CreateScatterPlot ws, qtyVariants, revQtyVariants, otherPhilToHydro, otherHydroToPhil
    
    ' Calculate and display average scores
    DisplayAverageScores ws, qtyVariants, revQtyVariants, otherPhilToHydro, otherHydroToPhil
    
    ' Highlight significant differences
    HighlightSignificantDifferences ws, qtyVariants, revQtyVariants, otherPhilToHydro, otherHydroToPhil

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Sub CreateScatterPlot(ws As Worksheet, qtyVariants As Collection, revQtyVariants As Collection, otherPhilToHydro As Collection, otherHydroToPhil As Collection)
    On Error GoTo ErrorHandler
    
    Dim cht As ChartObject
    Dim rng As Range
    Dim i As Long
    
    ' Prepare data for plotting
    ws.Cells(1, 1).Value = "Category"
    ws.Cells(1, 2).Value = "Score"
    
    i = 2
    For Each score In qtyVariants
        ws.Cells(i, 1).Value = "QTY to Hydrophobic"
        ws.Cells(i, 2).Value = score
        i = i + 1
    Next score
    
    For Each score In revQtyVariants
        ws.Cells(i, 1).Value = "Hydrophobic to QTY"
        ws.Cells(i, 2).Value = score
        i = i + 1
    Next score
    
    For Each score In otherPhilToHydro
        ws.Cells(i, 1).Value = "Other Hydrophilic to Hydrophobic"
        ws.Cells(i, 2).Value = score
        i = i + 1
    Next score
    
    For Each score In otherHydroToPhil
        ws.Cells(i, 1).Value = "Other Hydrophobic to Hydrophilic"
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

Sub DisplayAverageScores(ws As Worksheet, qtyVariants As Collection, revQtyVariants As Collection, otherPhilToHydro As Collection, otherHydroToPhil As Collection)
    On Error GoTo ErrorHandler
    
    ws.Cells(1, 4).Value = "Category"
    ws.Cells(1, 5).Value = "Average Score"
    
    ws.Cells(2, 4).Value = "QTY to Hydrophobic"
    ws.Cells(2, 5).Value = IIf(qtyVariants.Count > 0, WorksheetFunction.Average(CollectionToArray(qtyVariants)), "N/A")
    
    ws.Cells(3, 4).Value = "Hydrophobic to QTY"
    ws.Cells(3, 5).Value = IIf(revQtyVariants.Count > 0, WorksheetFunction.Average(CollectionToArray(revQtyVariants)), "N/A")
    
    ws.Cells(4, 4).Value = "Other Hydrophilic to Hydrophobic"
    ws.Cells(4, 5).Value = IIf(otherPhilToHydro.Count > 0, WorksheetFunction.Average(CollectionToArray(otherPhilToHydro)), "N/A")
    
    ws.Cells(5, 4).Value = "Other Hydrophobic to Hydrophilic"
    ws.Cells(5, 5).Value = IIf(otherHydroToPhil.Count > 0, WorksheetFunction.Average(CollectionToArray(otherHydroToPhil)), "N/A")
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in DisplayAverageScores: " & Err.Description
End Sub

Sub HighlightSignificantDifferences(ws As Worksheet, qtyVariants As Collection, revQtyVariants As Collection, otherPhilToHydro As Collection, otherHydroToPhil As Collection)
    On Error GoTo ErrorHandler
    
    Dim qtyAvg As Double, revQtyAvg As Double, otherPhilToHydroAvg As Double, otherHydroToPhilAvg As Double
    Dim qtyStdDev As Double, revQtyStdDev As Double, otherPhilToHydroStdDev As Double, otherHydroToPhilStdDev As Double
    
    ws.Cells(7, 4).Value = "Significant Differences:"
    
    If qtyVariants.Count > 1 And otherPhilToHydro.Count > 1 Then
        qtyAvg = WorksheetFunction.Average(CollectionToArray(qtyVariants))
        otherPhilToHydroAvg = WorksheetFunction.Average(CollectionToArray(otherPhilToHydro))
        qtyStdDev = WorksheetFunction.StDev(CollectionToArray(qtyVariants))
        otherPhilToHydroStdDev = WorksheetFunction.StDev(CollectionToArray(otherPhilToHydro))
        
        If Abs(qtyAvg - otherPhilToHydroAvg) > 2 * (qtyStdDev + otherPhilToHydroStdDev) / 2 Then
            ws.Cells(8, 4).Value = "QTY to Hydrophobic vs Other Hydrophilic to Hydrophobic"
        End If
    End If
    
    If revQtyVariants.Count > 1 And otherHydroToPhil.Count > 1 Then
        revQtyAvg = WorksheetFunction.Average(CollectionToArray(revQtyVariants))
        otherHydroToPhilAvg = WorksheetFunction.Average(CollectionToArray(otherHydroToPhil))
        revQtyStdDev = WorksheetFunction.StDev(CollectionToArray(revQtyVariants))
        otherHydroToPhilStdDev = WorksheetFunction.StDev(CollectionToArray(otherHydroToPhil))
        
        If Abs(revQtyAvg - otherHydroToPhilAvg) > 2 * (revQtyStdDev + otherHydroToPhilStdDev) / 2 Then
            ws.Cells(9, 4).Value = "Hydrophobic to QTY vs Other Hydrophobic to Hydrophilic"
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
