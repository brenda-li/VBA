Option Explicit

Const rho = 3000

Private Sub CommandButton1_Click()
    'Continue Button
    Dim Enjoyment As String
    Dim EnjoymentValue As Double
    Dim i As Integer

    Dim BestSalary, WorstSalary As Double
    Dim BestLocation, WorstLocation As Double
    Dim BestEnjoyment, WorstEnjoyment As Double

    Dim SalaryScale, LocationScale, EnjoymentScale As Double
    Dim SalaryWeight, LocationWeight, EnjoymentWeight As Double

Application.ScreenUpdating = False
'DATA VALIDATION
'Salary Data Entry
If Worksheets("ViewData").Range("H7").value <> "" And Worksheets("ViewData").Range("H7").value >= 0 And Worksheets("ViewData").Range("H7").value <= 100000 Then
If Worksheets("ViewData").Range("I7").value <> "" And Worksheets("ViewData").Range("I7").value >= 0 And Worksheets("ViewData").Range("I7").value <= 100000 Then
If Worksheets("ViewData").Range("J7").value <> "" And Worksheets("ViewData").Range("J7").value >= 0 And Worksheets("ViewData").Range("J7").value <= 100000 Then

'Location Data Entry
If Worksheets("ViewData").Range("H8").value <> "" Then
If Worksheets("ViewData").Range("I8").value <> "" Then
If Worksheets("ViewData").Range("J8").value <> "" Then

'Enjoyment of Work Data Entry
If Worksheets("ViewData").Range("H9").value <> "" Then
If Worksheets("ViewData").Range("I9").value <> "" Then
If Worksheets("ViewData").Range("J9").value <> "" Then

'Location Rank
If Worksheets("ViewRanks").Range("I6").value <> "" And Worksheets("ViewRanks").Range("I6").value >= 0 And Worksheets("ViewRanks").Range("I6").value <= 3 Then
If Worksheets("ViewRanks").Range("I7").value <> "" And Worksheets("ViewRanks").Range("I7").value >= 0 And Worksheets("ViewRanks").Range("I7").value <= 3 Then
If Worksheets("ViewRanks").Range("I8").value <> "" And Worksheets("ViewRanks").Range("I8").value >= 0 And Worksheets("ViewRanks").Range("I8").value <= 3 Then

'Enjoyment Rank
If Worksheets("ViewRanks").Range("L6").value <> "" And Worksheets("ViewRanks").Range("L6").value >= 0 And Worksheets("ViewRanks").Range("L6").value <= 100 Then
If Worksheets("ViewRanks").Range("L7").value <> "" And Worksheets("ViewRanks").Range("L7").value >= 0 And Worksheets("ViewRanks").Range("L7").value <= 100 Then
If Worksheets("ViewRanks").Range("L8").value <> "" And Worksheets("ViewRanks").Range("L8").value >= 0 And Worksheets("ViewRanks").Range("L8").value <= 100 Then
If Worksheets("ViewRanks").Range("L9").value <> "" And Worksheets("ViewRanks").Range("L9").value >= 0 And Worksheets("ViewRanks").Range("L9").value <= 100 Then
If Worksheets("ViewRanks").Range("L10").value <> "" And Worksheets("ViewRanks").Range("L10").value >= 0 And Worksheets("ViewRanks").Range("L10").value <= 100 Then

    If Worksheets("Model").OptionButton1 = True Then
        'Unhide Worksheet
        Worksheets("EvenSwap").Visible = True

        'EVEN SWAP DATA ENTRY
        With Worksheets("EvenSwap")
            'Write Salaries from ViewData Worksheet to EvenSwap Worksheet
            .Range("H7").value = Worksheets("ViewData").Range("H7").value
            .Range("I7").value = Worksheets("ViewData").Range("I7").value
            .Range("J7").value = Worksheets("ViewData").Range("J7").value

            'Write Location Numerical Equivalent to EvenSwap Worksheet
            .Range("H8").value = Worksheets("ViewRanks").Range("I6").value
            .Range("I8").value = Worksheets("ViewRanks").Range("I7").value
            .Range("J8").value = Worksheets("ViewRanks").Range("I8").value

            'Replace Word Enjoyment of Work Levels with Numerical Values on Worksheet
            For i = 1 To 3
                'Read in enjoyment level attribute as boring, okay, good, great, excellent.
                Enjoyment = Worksheets("ViewData").Cells(9, (7 + i)).value

                'Based on enjoyment level, read in the numerical representation to EnjoymentValue
                Select Case Enjoyment
                    Case "Boring":          EnjoymentValue = Worksheets("ViewRanks").Range("L6")
                    Case "Okay":            EnjoymentValue = Worksheets("ViewRanks").Range("L7")
                    Case "Good":            EnjoymentValue = Worksheets("ViewRanks").Range("L8")
                    Case "Great":           EnjoymentValue = Worksheets("ViewRanks").Range("L9")
                    Case "Excellent":       EnjoymentValue = Worksheets("ViewRanks").Range("L10")
                End Select

                'Read in the numerical representation for enjoyment of work to the model for decision making
                .Cells(9, (7 + i)).value = EnjoymentValue
            Next ' i

            'Write rankings to View Objective Ranking user form.
            ObjectiveRanking.LabelLocation1Rank = Worksheets("ViewRanks").Range("I6").value
            ObjectiveRanking.LabelLocation2Rank = Worksheets("ViewRanks").Range("I7").value
            ObjectiveRanking.LabelLocation3Rank = Worksheets("ViewRanks").Range("I8").value

            ObjectiveRanking.LabelBoringRank = Worksheets("ViewRanks").Range("L6").value
            ObjectiveRanking.LabelOkayRank = Worksheets("ViewRanks").Range("L7").value
            ObjectiveRanking.LabelGoodRank = Worksheets("ViewRanks").Range("L8").value
            ObjectiveRanking.LabelGreatRank = Worksheets("ViewRanks").Range("L9").value
            ObjectiveRanking.LabelExcellentRank = Worksheets("ViewRanks").Range("L10").value
        End With

        'NEXT STEP: Show EvenSwap Worksheet
        Worksheets("EvenSwap").Activate
        Worksheets("EvenSwap").Range("A1").Select

    ElseIf Worksheets("Model").OptionButton2 = True Then
        'Unhide Worksheet
        Worksheets("LinearValue").Visible = True

        'LINEAR VALUE DATA ENTRY
        'Read numbers into variable for recalculating salary values.
        BestSalary = Worksheets("ViewData").Range("K7")
        WorstSalary = Worksheets("ViewData").Range("L7")

        'Recalculate Salary Values and write to worksheet
        Worksheets("LinearValue").Range("H7") = 1 - ((Worksheets("ViewData").Range("H7") - BestSalary) / (WorstSalary - BestSalary))
        Worksheets("LinearValue").Range("I7") = 1 - ((Worksheets("ViewData").Range("I7") - BestSalary) / (WorstSalary - BestSalary))
        Worksheets("LinearValue").Range("J7") = 1 - ((Worksheets("ViewData").Range("J7") - BestSalary) / (WorstSalary - BestSalary))

        'Location and Enjoyment Values remain the same because they are already scaled when ranked. Write to Worksheet.
        With Worksheets("LinearValue")
            'Write Location Numerical Equivalent to LinearValue Worksheet
            .Range("H8").value = Worksheets("ViewRanks").Range("I6")
            .Range("I8").value = Worksheets("ViewRanks").Range("I7")
            .Range("J8").value = Worksheets("ViewRanks").Range("I8")

            'Replace Word Enjoyment of Work Levels with Numerical Values on Worksheet
            For i = 1 To 3
                'Read in enjoyment level attribute as boring, okay, good, great, excellent.
                Enjoyment = Worksheets("ViewData").Cells(9, (7 + i)).value

                'Based on enjoyment level, read in the numerical representation to EnjoymentValue
                Select Case Enjoyment
                    Case "Boring":          EnjoymentValue = Worksheets("ViewRanks").Range("L6")
                    Case "Okay":            EnjoymentValue = Worksheets("ViewRanks").Range("L7")
                    Case "Good":            EnjoymentValue = Worksheets("ViewRanks").Range("L8")
                    Case "Great":           EnjoymentValue = Worksheets("ViewRanks").Range("L9")
                    Case "Excellent":       EnjoymentValue = Worksheets("ViewRanks").Range("L10")
                End Select

                'Read in the numerical representation for enjoyment of work to the model for decision making
                .Cells(9, (7 + i)).value = EnjoymentValue
            Next ' i
        End With

        'Calculate Weighted Scores
        If Worksheets("ViewRanks").Range("O6") >= 0 And Worksheets("ViewRanks").Range("O6") <= 100 Then
            If Worksheets("ViewRanks").Range("O7") >= 0 And Worksheets("ViewRanks").Range("O7") <= 100 Then
                If Worksheets("ViewRanks").Range("O8") >= 0 And Worksheets("ViewRanks").Range("O8") <= 100 Then
                    'Write Values to Variables
                    SalaryScale = Worksheets("ViewRanks").Range("O6")
                    LocationScale = Worksheets("ViewRanks").Range("O7")
                    EnjoymentScale = Worksheets("ViewRanks").Range("O8")

                    'Calculate Weights
                    SalaryWeight = SalaryScale / (SalaryScale + LocationScale + EnjoymentScale)
                    LocationWeight = LocationScale / (SalaryScale + LocationScale + EnjoymentScale)
                    EnjoymentWeight = EnjoymentScale / (SalaryScale + LocationScale + EnjoymentScale)

                    'Calulcate Weighted Scores and Write to Linear Value Worksheets
                    If Round(SalaryWeight + LocationWeight + EnjoymentWeight, 0) = 1 Then
                        Worksheets("LinearValue").Range("H10") = Worksheets("LinearValue").Range("H7") * SalaryWeight _
                                                               + Worksheets("LinearValue").Range("H8") * LocationWeight _
                                                               + Worksheets("LinearValue").Range("H9") * EnjoymentWeight
                        Worksheets("LinearValue").Range("I10") = Worksheets("LinearValue").Range("I7") * SalaryWeight _
                                                               + Worksheets("LinearValue").Range("I8") * LocationWeight _
                                                               + Worksheets("LinearValue").Range("I9") * EnjoymentWeight
                        Worksheets("LinearValue").Range("J10") = Worksheets("LinearValue").Range("J7") * SalaryWeight _
                                                               + Worksheets("LinearValue").Range("J8") * LocationWeight _
                                                               + Worksheets("LinearValue").Range("J9") * EnjoymentWeight
                     Else
                         MsgBox ("Weights must sum to 1.")
                     End If

                    'Write Job Option Ranking to Worksheet
                    Worksheets("LinearValue").Range("H11") = WorksheetFunction.Rank(Worksheets("LinearValue").Range("H10"), Worksheets("LinearValue").Range("H10:J10"))
                    Worksheets("LinearValue").Range("I11") = WorksheetFunction.Rank(Worksheets("LinearValue").Range("I10"), Worksheets("LinearValue").Range("H10:J10"))
                    Worksheets("LinearValue").Range("J11") = WorksheetFunction.Rank(Worksheets("LinearValue").Range("J10"), Worksheets("LinearValue").Range("H10:J10"))

                    'Show Sensitivity Analysis Button
                    Worksheets("LinearValue").CommandButton3.Visible = True
                    Worksheets("LinearValue").CommandButton8.Visible = True

                    'Show LinearValue Worksheet and Resize Columns
                    Worksheets("LinearValue").Activate
                    Cells.EntireColumn.AutoFit
                    Worksheets("LinearValue").Range("A1").Select
                Else
                    MsgBox ("Please enter a value from 0 to 100.")
                End If
            Else
                MsgBox ("Please enter a value from 0 to 100.")
            End If
        Else
            MsgBox ("Please enter a value from 0 to 100.")
        End If
    ElseIf Worksheets("Model").OptionButton3 = True Then
        'Unhide Worksheet
        Worksheets("ExponValue").Visible = True

        'EXPONENTIAL VALUE DATA ENTRY
        'Read numbers into variable for recalculating salary values.
        BestSalary = Worksheets("ViewData").Range("K7")
        WorstSalary = Worksheets("ViewData").Range("L7")

        'Recalculate Salary Values and write to worksheet
        Worksheets("ExponValue").Range("H7") = (Exp(-(Worksheets("ViewData").Range("H7") - WorstSalary) / rho) - 1) / (Exp(-(BestSalary - WorstSalary) / rho) - 1)
        Worksheets("ExponValue").Range("I7") = (Exp(-(Worksheets("ViewData").Range("I7") - WorstSalary) / rho) - 1) / (Exp(-(BestSalary - WorstSalary) / rho) - 1)
        Worksheets("ExponValue").Range("J7") = (Exp(-(Worksheets("ViewData").Range("J7") - WorstSalary) / rho) - 1) / (Exp(-(BestSalary - WorstSalary) / rho) - 1)

        'Location and Enjoyment Values remain the same because they are already scaled when ranked. Write to Worksheet.
        With Worksheets("ExponValue")
            'Write Location Numerical Equivalent to ExponValue Worksheet
            .Range("H8").value = Worksheets("ViewRanks").Range("I6")
            .Range("I8").value = Worksheets("ViewRanks").Range("I7")
            .Range("J8").value = Worksheets("ViewRanks").Range("I8")

            'Replace Word Enjoyment of Work Levels with Numerical Values on Worksheet
            For i = 1 To 3
                'Read in enjoyment level attribute as boring, okay, good, great, excellent.
                Enjoyment = Worksheets("ViewData").Cells(9, (7 + i)).value

                'Based on enjoyment level, read in the numerical representation to EnjoymentValue
                Select Case Enjoyment
                    Case "Boring":          EnjoymentValue = Worksheets("ViewRanks").Range("L6")
                    Case "Okay":            EnjoymentValue = Worksheets("ViewRanks").Range("L7")
                    Case "Good":            EnjoymentValue = Worksheets("ViewRanks").Range("L8")
                    Case "Great":           EnjoymentValue = Worksheets("ViewRanks").Range("L9")
                    Case "Excellent":       EnjoymentValue = Worksheets("ViewRanks").Range("L10")
                End Select

                'Read in the numerical representation for enjoyment of work to the model for decision making
                .Cells(9, (7 + i)).value = EnjoymentValue
            Next ' i
        End With

        'Calculate Weighted Scores
        If Worksheets("ViewRanks").Range("O6") >= 0 And Worksheets("ViewRanks").Range("O6") <= 100 Then
            If Worksheets("ViewRanks").Range("O7") >= 0 And Worksheets("ViewRanks").Range("O7") <= 100 Then
                If Worksheets("ViewRanks").Range("O8") >= 0 And Worksheets("ViewRanks").Range("O8") <= 100 Then
                    'Write Values to Variables
                    SalaryScale = Worksheets("ViewRanks").Range("O6")
                    LocationScale = Worksheets("ViewRanks").Range("O7")
                    EnjoymentScale = Worksheets("ViewRanks").Range("O8")

                    'Calculate Weights
                    SalaryWeight = SalaryScale / (SalaryScale + LocationScale + EnjoymentScale)
                    LocationWeight = LocationScale / (SalaryScale + LocationScale + EnjoymentScale)
                    EnjoymentWeight = EnjoymentScale / (SalaryScale + LocationScale + EnjoymentScale)

                    'Calulcate Weighted Scores and Write to Expon Value Worksheets
                    If Round(SalaryWeight + LocationWeight + EnjoymentWeight, 0) = 1 Then
                        Worksheets("ExponValue").Range("H10") = Worksheets("ExponValue").Range("H7") * SalaryWeight _
                                                              + Worksheets("ExponValue").Range("H8") * LocationWeight _
                                                              + Worksheets("ExponValue").Range("H9") * EnjoymentWeight
                        Worksheets("ExponValue").Range("I10") = Worksheets("ExponValue").Range("I7") * SalaryWeight _
                                                              + Worksheets("ExponValue").Range("I8") * LocationWeight _
                                                              + Worksheets("ExponValue").Range("I9") * EnjoymentWeight
                        Worksheets("ExponValue").Range("J10") = Worksheets("ExponValue").Range("J7") * SalaryWeight _
                                                              + Worksheets("ExponValue").Range("J8") * LocationWeight _
                                                              + Worksheets("ExponValue").Range("J9") * EnjoymentWeight
                     Else
                         MsgBox ("Weights must sum to 1.")
                     End If

                    'Write Job Option Ranking to Worksheet
                    Worksheets("ExponValue").Range("H11") = WorksheetFunction.Rank(Worksheets("ExponValue").Range("H10"), Worksheets("ExponValue").Range("H10:J10"))
                    Worksheets("ExponValue").Range("I11") = WorksheetFunction.Rank(Worksheets("ExponValue").Range("I10"), Worksheets("ExponValue").Range("H10:J10"))
                    Worksheets("ExponValue").Range("J11") = WorksheetFunction.Rank(Worksheets("ExponValue").Range("J10"), Worksheets("ExponValue").Range("H10:J10"))

                    'Show Sensitivity Analysis Button
                    Worksheets("ExponValue").CommandButton3.Visible = True
                    Worksheets("ExponValue").CommandButton8.Visible = True

                    'Show ExponValue Worksheet and Resize Columns
                    Worksheets("ExponValue").Activate
                    Cells.EntireColumn.AutoFit
                    Worksheets("ExponValue").Range("A1").Select
                Else
                    MsgBox ("Please enter a value from 0 to 100.")
                End If
            Else
                MsgBox ("Please enter a value from 0 to 100.")
            End If
        Else
            MsgBox ("Please enter a value from 0 to 100.")
        End If
    Else
        MsgBox ("Please return to model worksheet and select a model option.")
        Worksheets("Model").Activate
    End If
'ENJOYMENT RANK
Else
    MsgBox ("Please fill in a value for the excellent scale.")
    RankObjectives.Show
    RankObjectives.MultiPage1.value = 1
End If
Else
    MsgBox ("Please fill in a value for the great scale.")
    RankObjectives.Show
    RankObjectives.MultiPage1.value = 1
End If
Else
    MsgBox ("Please fill in a value for the good scale.")
    RankObjectives.Show
    RankObjectives.MultiPage1.value = 1
End If
Else
    MsgBox ("Please fill in a value for the okay scale.")
    RankObjectives.Show
    RankObjectives.MultiPage1.value = 1
End If
Else
    MsgBox ("Please fill in a value for the boring scale.")
    RankObjectives.Show
    RankObjectives.MultiPage1.value = 1
End If

'LOCATION RANK
Else
    MsgBox ("Please make sure all locations are ranked.")
    RankObjectives.Show
    RankObjectives.MultiPage1.value = 0
End If
Else
    MsgBox ("Please make sure all locations are ranked.")
    RankObjectives.Show
    RankObjectives.MultiPage1.value = 0
End If
Else
    MsgBox ("Please make sure all locations are ranked.")
    RankObjectives.Show
    RankObjectives.MultiPage1.value = 0
End If

'ENJOYMENT DATA ENTRY
Else
    MsgBox ("Please select a value for Job C Enjoyment of Work.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 1
End If
Else
    MsgBox ("Please select a value for Job B Enjoyment of Work.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 1
End If
Else
    MsgBox ("Please select a value for Job A Enjoyment of Work.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 1
End If

'LOCATION DATA ENTRY
Else
    MsgBox ("Please fill in a location for Job C.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 0
End If
Else
    MsgBox ("Please fill in a location for Job B.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 0
End If
Else
    MsgBox ("Please fill in a location for Job A.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 0
End If

'SALARY DATA ENTRY
Else
    MsgBox ("Please fill in the monthly salary for Job C.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 0
End If
Else
    MsgBox ("Please fill in the monthly salary for Job B.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 0
End If
Else
    MsgBox ("Please fill in the monthly salary for Job A.")
    DataEntry.Show
    DataEntry.MultiPage1.value = 0
End If

Application.ScreenUpdating = True

End Sub

Private Sub CommandButton2_Click()
    'Revise Rankings Button
    Worksheets("RankObjectives").Activate
End Sub
