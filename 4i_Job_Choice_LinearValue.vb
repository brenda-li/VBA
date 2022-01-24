Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    'Hide Worksheets
    Worksheets("Model").Visible = False
    Worksheets("DataEntry").Visible = False
    Worksheets("ViewData").Visible = False
    Worksheets("RankObjectives").Visible = False
    Worksheets("ObjectiveWeights").Visible = False
    Worksheets("ViewRanks").Visible = False
    Worksheets("EvenSwap").Visible = False
    Worksheets("LinearValue").Visible = False
    Worksheets("ExponValue").Visible = False
    Worksheets("SensAnalysis").Visible = False

    'Clear Location Combo Box
    Worksheets("DataEntry").Activate
    Range("R8:R10").Select
    Selection.ClearContents
    Range("A1").Select

    'Clear Previous Data Entries
    Worksheets("ViewData").Activate
    Range("H7:L9").Select
    Selection.ClearContents
    Range("A1").Select

    Worksheets("ViewData").Activate
    Range("K6:L6").Select
    Selection.ClearContents
    Range("A1").Select

    'Clear Previous Rankings
    Worksheets("ViewRanks").Activate
    Range("H6:I8").Select
    Selection.ClearContents

    Worksheets("ViewRanks").Activate
    Range("L6:L10").Select
    Selection.ClearContents

    Worksheets("ViewRanks").Activate
    Range("N5:O8").Select
    Selection.ClearContents
    Range("A1").Select

    'Hide Show Recommendation & Sensitivity Analysis Button on Linear Value Worksheet
    Worksheets("LinearValue").CommandButton3.Visible = False
    Worksheets("LinearValue").CommandButton8.Visible = False

    'Hide Show Recommendation & Sensitivity Analysis Button on Exponential Value Worksheet
    Worksheets("ExponValue").CommandButton3.Visible = False
    Worksheets("ExponValue").CommandButton8.Visible = False

    'Hide Return to ___ Model Buttons
    Worksheets("SensAnalysis").CommandButton10.Visible = False
    Worksheets("SensAnalysis").CommandButton11.Visible = False

    'Clear Sensitivity Analysis Model Title
    Worksheets("SensAnalysis").Activate
    Worksheets("SensAnalysis").Range("B3").Select
    Selection.ClearContents
    Range("A1").Select

    'Clear Sensitivity Analysis
    Worksheets("SensAnalysis").Activate
    Range("A6:D101").Select
    Selection.ClearContents
    Range("A1").Select

    'Format View Data
    Worksheets("ViewData").Activate
    Range("E3:N3").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Range("E3:L3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("G5:L5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Selection.UnMerge
    Range("G5:J5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("K6:L9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("J6:J9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Range("M3:N3").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("D10").Select

    'Format ViewRanks
    Worksheets("ViewRanks").Activate
    Range("N5:O8").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("D1").Select

    'Clear Even Swap
    Worksheets("EvenSwap").Activate
    Range("H7:J9").Select
    Selection.ClearContents
    Range("A1").Select

    'Clear Linear Value
    Worksheets("LinearValue").Activate
    Range("H7:J11").Select
    Selection.ClearContents
    Range("A1").Select

    'Clear Exponential Value
    Worksheets("ExponValue").Activate
    Range("H7:J11").Select
    Selection.ClearContents
    Range("A1").Select

    'Make sure Even Swap Data Visible
    Worksheets("EvenSwap").Activate
    Range("G6:J9").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Worksheets("EvenSwap").Range("A1").Select

    'Show Welcome Screen
    Worksheets("Welcome").Activate

    Application.ScreenUpdating = True
End Sub
