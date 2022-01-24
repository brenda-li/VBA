Option Explicit

Private Sub CommandButton1_Click()
'Swap Data Button
    Dim MaxSalary As Double
    Dim BestOption As String
    Dim BestSalary As Double
    Dim BestLocation As String
    Dim BestEnjoyment As String

    Dim SalaryA, SalaryB, SalaryC As Double
    Dim LocationA, LocationB, LocationC As Double
    Dim EnjoymentA, EnjoymentB, EnjoymentC As Double

    SalaryA = Worksheets("EvenSwap").Range("H7").value
    SalaryB = Worksheets("EvenSwap").Range("I7").value
    SalaryC = Worksheets("EvenSwap").Range("J7").value

    EnjoymentA = Worksheets("EvenSwap").Range("H9").value
    EnjoymentB = Worksheets("EvenSwap").Range("I9").value
    EnjoymentC = Worksheets("EvenSwap").Range("J9").value

'Show Recommendation if no more swap options.
If EvenSwap.OptionButton1.Visible = False And EvenSwap.OptionButton2.Visible = False Then
    SwapRecommendation

ElseIf (SalaryA > SalaryB) And (SalaryA > SalaryC) And (LocationA > LocationB) And (LocationA > LocationC) And (EnjoymentA > EnjoymentB) And (EnjoymentA > EnjoymentC) Then
    Worksheets("EvenSwap").Activate
    MaxSalary = WorksheetFunction.Max(Worksheets("EvenSwap").Range("H7:J7"))

    BestOption = WorksheetFunction.Index(Worksheets("EvenSwap").Range("H6:J7"), 1, WorksheetFunction.Match(MaxSalary, Worksheets("EvenSwap").Range("H7:J7"), 0))

    BestSalary = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 2, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))
    BestLocation = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 3, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))
    BestEnjoyment = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 4, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))

    MsgBox ("Based on the even swap, you should select " & BestOption & _
            " which has a monthly salary of " & format(BestSalary, "Currency") & ". This job is located in " & BestLocation & _
            " and has a " & BestEnjoyment & " Enjoyment of work rating."), , "Recommendation"
ElseIf (SalaryB > SalaryA) And (SalaryB > SalaryC) And (LocationB > LocationA) And (LocationB > LocationC) And (EnjoymentB > EnjoymentA) And (EnjoymentB > EnjoymentC) Then
    Worksheets("EvenSwap").Activate
    MaxSalary = WorksheetFunction.Max(Worksheets("EvenSwap").Range("H7:J7"))

    BestOption = WorksheetFunction.Index(Worksheets("EvenSwap").Range("H6:J7"), 1, WorksheetFunction.Match(MaxSalary, Worksheets("EvenSwap").Range("H7:J7"), 0))

    BestSalary = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 2, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))
    BestLocation = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 3, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))
    BestEnjoyment = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 4, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))

    MsgBox ("Based on the even swap, you should select " & BestOption & _
            " which has a monthly salary of " & format(BestSalary, "Currency") & ". This job is located in " & BestLocation & _
            " and has a " & BestEnjoyment & " Enjoyment of work rating."), , "Recommendation"
ElseIf (SalaryC > SalaryA) And (SalaryC > SalaryB) And (LocationC > LocationA) And (LocationC > LocationB) And (EnjoymentC > EnjoymentA) And (EnjoymentC > EnjoymentB) Then
    Worksheets("EvenSwap").Activate
    MaxSalary = WorksheetFunction.Max(Worksheets("EvenSwap").Range("H7:J7"))

    BestOption = WorksheetFunction.Index(Worksheets("EvenSwap").Range("H6:J7"), 1, WorksheetFunction.Match(MaxSalary, Worksheets("EvenSwap").Range("H7:J7"), 0))

    BestSalary = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 2, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))
    BestLocation = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 3, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))
    BestEnjoyment = WorksheetFunction.Index(Worksheets("ViewData").Range("H6:J9"), 4, WorksheetFunction.Match(BestOption, Worksheets("ViewData").Range("H6:J6"), 0))

    MsgBox ("Based on the even swap, you should select " & BestOption & _
            " which has a monthly salary of " & format(BestSalary, "Currency") & ". This job is located in " & BestLocation & _
            " and has a " & BestEnjoyment & " Enjoyment of work rating."), , "Recommendation"
ElseIf EvenSwap.OptionButton1.Visible = False Or EvenSwap.OptionButton2.Visible = False Then
    SwapRecommendation
Else
    EvenSwap.Show
End If

End Sub

Private Sub CommandButton3_Click()
    'Help
    DecisionModelHelp.MultiPage1.Pages(0).Visible = False
    DecisionModelHelp.MultiPage1.Pages(2).Visible = False
    DecisionModelHelp.MultiPage1.Pages(3).Visible = False
    DecisionModelHelp.MultiPage1.Pages(4).Visible = False

    DecisionModelHelp.MultiPage1.Pages(1).Visible = True
    DecisionModelHelp.MultiPage1.value = 1

    DecisionModelHelp.Show
End Sub

Private Sub CommandButton5_Click()
'Reset Swap
    'Monthly Salary
    Worksheets("EvenSwap").Range("H7").value = Worksheets("ViewData").Range("H7").value
    Worksheets("EvenSwap").Range("I7").value = Worksheets("ViewData").Range("I7").value
    Worksheets("EvenSwap").Range("J7").value = Worksheets("ViewData").Range("J7").value

    'Location
    Worksheets("EvenSwap").Range("H8").value = Worksheets("ViewRanks").Range("I6")
    Worksheets("EvenSwap").Range("I8").value = Worksheets("ViewRanks").Range("I7")
    Worksheets("EvenSwap").Range("J8").value = Worksheets("ViewRanks").Range("I8")

    'Enjoyment of Work
    Dim Enjoyment As String
    Dim EnjoymentValue As Double
    Dim i As Integer

    For i = 1 To 3
        'Read in enjoyment level attribute as boring, okay, good, great, excellent.
        Enjoyment = Worksheets("ViewData").Cells(9, (7 + i)).value

        'Based on enjoyment level, read in the numerical representation to EnjoymentValue.
        Select Case Enjoyment
            Case "Boring":          EnjoymentValue = Worksheets("ViewRanks").Range("L6")
            Case "Okay":            EnjoymentValue = Worksheets("ViewRanks").Range("L7")
            Case "Good":            EnjoymentValue = Worksheets("ViewRanks").Range("L8")
            Case "Great":           EnjoymentValue = Worksheets("ViewRanks").Range("L9")
            Case "Excellent":       EnjoymentValue = Worksheets("ViewRanks").Range("L10")
        End Select

        'Read in the numerical representation for enjoyment of work to the model for decision making
        Worksheets("EvenSwap").Cells(9, (7 + i)).value = EnjoymentValue
    Next ' i

    'Show All Option Buttons
    EvenSwap.OptionButton1.Visible = True
    EvenSwap.OptionButton2.Visible = True

    'Show All Text
    Worksheets("EvenSwap").Range("G6:J9").Select
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With

    Worksheets("EvenSwap").Range("A1").Select
End Sub

Private Sub CommandButton8_Click()
    'Go Button
    If ComboNavigation = "Welcome" Then
        Worksheets("Welcome").Activate
    ElseIf ComboNavigation = "Model Selection" Then
        Worksheets("Model").Activate
    ElseIf ComboNavigation = "Data Entry" Then
        Worksheets("DataEntry").Activate
    ElseIf ComboNavigation = "View Data" Then
        Worksheets("ViewData").Activate
    ElseIf ComboNavigation = "Rank Objectives" Then
        Worksheets("RankObjectives").Activate
    ElseIf ComboNavigation = "View Rankings" Then
        Worksheets("ViewRanks").Activate
    Else
        MsgBox ("Please select an option from the list before pressing go.")
    End If
End Sub
