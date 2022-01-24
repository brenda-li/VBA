Option Explicit

Dim JobAScore, JobBScore, JobCScore As Double
Dim SensAnalysisTest As Boolean

Const rho = 3000

Private Sub CommandButton3_Click()
    'Show Recommendation
    Dim JobARank, JobBRank, JobCRank As Integer
    Dim JobOption1, JobOption2, JobOption3 As String
    Dim Salary1, Salary2, Salary3 As Double
    Dim Location1, Location2, Location3 As String
    Dim Enjoyment1, Enjoyment2, Enjoyment3 As String

    Dim ViewDataTable As Range

If (Worksheets("ExponValue").Range("H11") = Worksheets("ExponValue").Range("I11")) And (Worksheets("ExponValue").Range("H11") = Worksheets("ExponValue").Range("J11")) Then
    'All options have the same ranking.
    MsgBox ("All options have the same ranking." & vbLf & vbLf & _
            "Please restart the decision making process and use the navigation to return to the data entry worksheet. Use other values when creating value scales or rating the importance of objectives."), , "No Recommendation"
ElseIf Worksheets("ExponValue").Range("J10") = 0 Then
    'Only Two Options
            Set ViewDataTable = Worksheets("ViewData").Range("H6:J9")

            'Decide which Job Option is related to Rank 1, Rank 2, and Rank 3
            With Worksheets("ExponValue")
                JobOption1 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(1, .Range("H11:J11"), 0))
                JobOption2 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(2, .Range("H11:J11"), 0))
            End With

            'Match job options with salary, location, and enjoyment
            Salary1 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Salary2 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))

            Location1 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Location2 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))

            Enjoyment1 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Enjoyment2 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))

            'Recommendation
            MsgBox ("Based on the exponential value function, you should select a job based on the folowing order: " & vbLf & vbLf & _
                    "1. [ " & JobOption1 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary1, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location1 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment1 & vbLf & _
                    "2. [ " & JobOption2 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary2, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location2 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment2 & vbLf), , "Recommendation"
    Else

If Worksheets("ExponValue").Range("H11") <> Worksheets("ExponValue").Range("I11") Then
    If Worksheets("ExponValue").Range("H11") <> Worksheets("ExponValue").Range("J11") Then
        If Worksheets("ExponValue").Range("I11") <> Worksheets("ExponValue").Range("J11") Then
            'No job options have the same ranking.
            Set ViewDataTable = Worksheets("ViewData").Range("H6:J9")

            'Decide which Job Option is related to Rank 1, Rank 2, and Rank 3
            With Worksheets("ExponValue")
                JobOption1 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(1, .Range("H11:J11"), 0))
                JobOption2 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(2, .Range("H11:J11"), 0))
                JobOption3 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(3, .Range("H11:J11"), 0))
            End With

            'Match job options with salary, location, and enjoyment
            Salary1 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Salary2 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
            Salary3 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

            Location1 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Location2 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
            Location3 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

            Enjoyment1 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Enjoyment2 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
            Enjoyment3 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

            'Recommendation
            MsgBox ("Based on the exponential value function, you should select a job based on the folowing order: " & vbLf & vbLf & _
                    "1. [ " & JobOption1 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary1, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location1 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment1 & vbLf & _
                    "2. [ " & JobOption2 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary2, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location2 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment2 & vbLf & _
                    "3. [ " & JobOption3 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary3, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location3 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment3), , "Recommendation"
        Else
            'Job B and Job C options tied.
            Set ViewDataTable = Worksheets("ViewData").Range("H6:J9")

            'Decide which Job Option is related to Rank 1, Rank 2, and Rank 3
            With Worksheets("ExponValue")
                JobOption1 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(1, .Range("H11:J11"), 0))
                JobOption2 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(2, .Range("H11:J11"), 0))
            End With

            If JobOption1 <> "Job A" And JobOption2 <> "Job A" Then
                JobOption3 = "Job A"
            ElseIf JobOption1 <> "Job B" And JobOption2 <> "Job B" Then
                JobOption3 = "Job B"
            ElseIf JobOption1 <> "Job C" And JobOption2 <> "Job C" Then
                JobOption3 = "Job C"
            End If

            'Match job options with salary, location, and enjoyment
            Salary1 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Salary2 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
            Salary3 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

            Location1 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Location2 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
            Location3 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

            Enjoyment1 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
            Enjoyment2 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
            Enjoyment3 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

            If JobOption1 = "Job A" Then
                'Recommendation
                MsgBox ("Job B and Job C are tied." & vbLf & vbLf & _
                        "Based on the exponential value function, you should select a job based on the folowing order: " & vbLf & vbLf & _
                        "1. [ " & JobOption1 & " ]" & vbLf & vbTab & _
                            "Monthly Salary: " & vbTab & vbTab & format(Salary1, "Currency") & vbLf & vbTab & _
                            "Location: " & vbTab & vbTab & vbTab & Location1 & vbLf & vbTab & _
                            "Work Enjoyment Level: " & vbTab & Enjoyment1 & vbLf & _
                        "2. [ " & JobOption2 & " ]" & vbLf & vbTab & _
                            "Monthly Salary: " & vbTab & vbTab & format(Salary2, "Currency") & vbLf & vbTab & _
                            "Location: " & vbTab & vbTab & vbTab & Location2 & vbLf & vbTab & _
                            "Work Enjoyment Level: " & vbTab & Enjoyment2 & vbLf & _
                        "   [ " & JobOption3 & " ]" & vbLf & vbTab & _
                            "Monthly Salary: " & vbTab & vbTab & format(Salary3, "Currency") & vbLf & vbTab & _
                            "Location: " & vbTab & vbTab & vbTab & Location3 & vbLf & vbTab & _
                            "Work Enjoyment Level: " & vbTab & Enjoyment3), , "Recommendation"
            Else
                'Recommendation
                MsgBox ("Job B and Job C are tied." & vbLf & vbLf & _
                        "Based on the exponential value function, you should select a job based on the folowing order: " & vbLf & vbLf & _
                        "1. [ " & JobOption3 & " ]" & vbLf & vbTab & _
                            "Monthly Salary: " & vbTab & vbTab & format(Salary3, "Currency") & vbLf & vbTab & _
                            "Location: " & vbTab & vbTab & vbTab & Location3 & vbLf & vbTab & _
                            "Work Enjoyment Level: " & vbTab & Enjoyment3 & vbLf & _
                        "   [ " & JobOption1 & " ]" & vbLf & vbTab & _
                            "Monthly Salary: " & vbTab & vbTab & format(Salary1, "Currency") & vbLf & vbTab & _
                            "Location: " & vbTab & vbTab & vbTab & Location1 & vbLf & vbTab & _
                            "Work Enjoyment Level: " & vbTab & Enjoyment1 & vbLf & _
                        "2. [ " & JobOption2 & " ]" & vbLf & vbTab & _
                            "Monthly Salary: " & vbTab & vbTab & format(Salary2, "Currency") & vbLf & vbTab & _
                            "Location: " & vbTab & vbTab & vbTab & Location2 & vbLf & vbTab & _
                            "Work Enjoyment Level: " & vbTab & Enjoyment2), , "Recommendation"
            End If
        End If
    Else
        'Job A and Job C options tied.
        Set ViewDataTable = Worksheets("ViewData").Range("H6:J9")

        'Decide which Job Option is related to Rank 1, Rank 2, and Rank 3
        With Worksheets("ExponValue")
            JobOption1 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(1, .Range("H11:J11"), 0))
            JobOption2 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(2, .Range("H11:J11"), 0))
        End With

        If JobOption1 <> "Job A" And JobOption2 <> "Job A" Then
            JobOption3 = "Job A"
        ElseIf JobOption1 <> "Job B" And JobOption2 <> "Job B" Then
            JobOption3 = "Job B"
        ElseIf JobOption1 <> "Job C" And JobOption2 <> "Job C" Then
            JobOption3 = "Job C"
        End If

        'Match job options with salary, location, and enjoyment
        Salary1 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
        Salary2 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
        Salary3 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

        Location1 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
        Location2 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
        Location3 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

        Enjoyment1 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
        Enjoyment2 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
        Enjoyment3 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

        If JobOption1 = "Job B" Then
            'Recommendation
            MsgBox ("Job A and Job C are tied." & vbLf & vbLf & _
                    "Based on the exponential value function, you should select a job based on the folowing order: " & vbLf & vbLf & _
                    "1. [ " & JobOption1 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary1, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location1 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment1 & vbLf & _
                    "2. [ " & JobOption2 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary2, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location2 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment2 & vbLf & _
                    "   [ " & JobOption3 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary3, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location3 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment3), , "Recommendation"
        Else
            'Recommendation
            MsgBox ("Job A and Job C are tied." & vbLf & vbLf & _
                    "Based on the exponential value function, you should select a job based on the folowing order: " & vbLf & vbLf & _
                    "1. [ " & JobOption3 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary3, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location3 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment3 & vbLf & _
                    "   [ " & JobOption1 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary1, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location1 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment1 & vbLf & _
                    "2. [ " & JobOption2 & " ]" & vbLf & vbTab & _
                        "Monthly Salary: " & vbTab & vbTab & format(Salary2, "Currency") & vbLf & vbTab & _
                        "Location: " & vbTab & vbTab & vbTab & Location2 & vbLf & vbTab & _
                        "Work Enjoyment Level: " & vbTab & Enjoyment2), , "Recommendation"
        End If
    End If
Else
    'Job A and Job B options tied.
    Set ViewDataTable = Worksheets("ViewData").Range("H6:J9")

    'Decide which Job Option is related to Rank 1, Rank 2, and Rank 3
    With Worksheets("ExponValue")
        JobOption1 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(1, .Range("H11:J11"), 0))
        JobOption2 = WorksheetFunction.Index(.Range("H6:J11"), 1, WorksheetFunction.Match(2, .Range("H11:J11"), 0))
    End With

    If JobOption1 <> "Job A" And JobOption2 <> "Job A" Then
        JobOption3 = "Job A"
    ElseIf JobOption1 <> "Job B" And JobOption2 <> "Job B" Then
        JobOption3 = "Job B"
    ElseIf JobOption1 <> "Job C" And JobOption2 <> "Job C" Then
        JobOption3 = "Job C"
    End If

    'Match job options with salary, location, and enjoyment
    Salary1 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
    Salary2 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
    Salary3 = WorksheetFunction.Index(ViewDataTable, 2, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

    Location1 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
    Location2 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
    Location3 = WorksheetFunction.Index(ViewDataTable, 3, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

    Enjoyment1 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption1, Worksheets("ViewData").Range("H6:J6")))
    Enjoyment2 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption2, Worksheets("ViewData").Range("H6:J6")))
    Enjoyment3 = WorksheetFunction.Index(ViewDataTable, 4, WorksheetFunction.Match(JobOption3, Worksheets("ViewData").Range("H6:J6")))

    If JobOption1 = "Job C" Then
        'Recommendation
        MsgBox ("Job A and Job B are tied." & vbLf & vbLf & _
                "Based on the exponential value function, you should select a job based on the folowing order: " & vbLf & vbLf & _
                "1. [ " & JobOption1 & " ]" & vbLf & vbTab & _
                    "Monthly Salary: " & vbTab & vbTab & format(Salary1, "Currency") & vbLf & vbTab & _
                    "Location: " & vbTab & vbTab & vbTab & Location1 & vbLf & vbTab & _
                    "Work Enjoyment Level: " & vbTab & Enjoyment1 & vbLf & _
                "2. [ " & JobOption2 & " ]" & vbLf & vbTab & _
                    "Monthly Salary: " & vbTab & vbTab & format(Salary2, "Currency") & vbLf & vbTab & _
                    "Location: " & vbTab & vbTab & vbTab & Location2 & vbLf & vbTab & _
                    "Work Enjoyment Level: " & vbTab & Enjoyment2 & vbLf & _
                "   [ " & JobOption3 & " ]" & vbLf & vbTab & _
                    "Monthly Salary: " & vbTab & vbTab & format(Salary3, "Currency") & vbLf & vbTab & _
                    "Location: " & vbTab & vbTab & vbTab & Location3 & vbLf & vbTab & _
                    "Work Enjoyment Level: " & vbTab & Enjoyment3), , "Recommendation"
    Else
        'Recommendation
        MsgBox ("Job A and Job B are tied." & vbLf & vbLf & _
                "Based on the exponential value function, you should select a job based on the folowing order: " & vbLf & vbLf & _
                "1. [ " & JobOption3 & " ]" & vbLf & vbTab & _
                    "Monthly Salary: " & vbTab & vbTab & format(Salary3, "Currency") & vbLf & vbTab & _
                    "Location: " & vbTab & vbTab & vbTab & Location3 & vbLf & vbTab & _
                    "Work Enjoyment Level: " & vbTab & Enjoyment3 & vbLf & _
                "   [ " & JobOption1 & " ]" & vbLf & vbTab & _
                    "Monthly Salary: " & vbTab & vbTab & format(Salary1, "Currency") & vbLf & vbTab & _
                    "Location: " & vbTab & vbTab & vbTab & Location1 & vbLf & vbTab & _
                    "Work Enjoyment Level: " & vbTab & Enjoyment1 & vbLf & _
                "2. [ " & JobOption2 & " ]" & vbLf & vbTab & _
                    "Monthly Salary: " & vbTab & vbTab & format(Salary2, "Currency") & vbLf & vbTab & _
                    "Location: " & vbTab & vbTab & vbTab & Location2 & vbLf & vbTab & _
                    "Work Enjoyment Level: " & vbTab & Enjoyment2), , "Recommendation"
    End If
End If

End If

End Sub

Private Sub CommandButton4_Click()
    'Exponential Value Function Help
    DecisionModelHelp.MultiPage1.Pages(0).Visible = False
    DecisionModelHelp.MultiPage1.Pages(1).Visible = False
    DecisionModelHelp.MultiPage1.Pages(2).Visible = False
    DecisionModelHelp.MultiPage1.Pages(4).Visible = False

    DecisionModelHelp.MultiPage1.Pages(3).Visible = True
    DecisionModelHelp.MultiPage1.value = 3

    DecisionModelHelp.Show
End Sub

Private Sub CommandButton8_Click()
    Worksheets("SensAnalysis").Visible = True
    RunSensitivityAnalysis
    Worksheets("SensAnalysis").Range("B3") = "Exponential Value Model"
    Worksheets("SensAnalysis").CommandButton11.Visible = True
End Sub

Private Sub RunSensitivityAnalysis()
    'Data Validation - Within Scale
    If Worksheets("ViewRanks").Range("O6") >= 0 And Worksheets("ViewRanks").Range("O6") <= 100 Then
        If Worksheets("ViewRanks").Range("O7") >= 0 And Worksheets("ViewRanks").Range("O7") <= 100 Then
            If Worksheets("ViewRanks").Range("O8") >= 0 And Worksheets("ViewRanks").Range("O8") <= 100 Then

                Application.ScreenUpdating = False

                'Sensitivity Analysis Button
                SalarySensAnalysis
                SalaryImportanceSensAnalysis
                LocationImportanceSensAnalysis
                EnjoymentImportanceSensAnalysis

                Application.ScreenUpdating = True

                'Show Sensitivity Analysis Graph
                Worksheets("SensAnalysis").Activate
                ActiveWindow.SmallScroll Up:=100

             Else
                MsgBox ("Enjoyment scale value must be from 0 to 100.")
            End If
        Else
           MsgBox ("Location scale value must be from 0 to 100.")
        End If
    Else
        MsgBox ("Salary scale value must be from 0 to 100.")
    End If
End Sub
Private Sub SalarySensAnalysis()
    Dim tempA1, tempB1, tempC1 As Double
    Dim tempA2, tempB2, tempC2 As Double
    Dim i As Integer
    Dim StartSalary, NextSalary As Double
    Dim BestSalary, WorstSalary As Double

    SensAnalysisTest = True

    'Store original salary values from View Data Worksheet.
    tempA1 = Worksheets("ViewData").Range("H7")
    tempB1 = Worksheets("ViewData").Range("I7")
    tempC1 = Worksheets("ViewData").Range("J7")

    'Store original recalculated salaries from Expon Value Worksheet
    tempA2 = Worksheets("ExponValue").Range("H7")
    tempB2 = Worksheets("ExponValue").Range("I7")
    tempC2 = Worksheets("ExponValue").Range("J7")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("A6").value = "Monthly Salary"
    Worksheets("SensAnalysis").Range("B6").value = "Job A"
    Worksheets("SensAnalysis").Range("C6").value = "Job B"
    Worksheets("SensAnalysis").Range("D6").value = "Job C"

    'Set variable monthly salary.
    StartSalary = 0
    NextSalary = StartSalary

    'Fill in Sensitivity Analysis Table for $0 to $10,000 (Monthly Salary).
    For i = 0 To 20
        Worksheets("SensAnalysis").Range("A" & (i + 7)).value = format(NextSalary, "$##,##0")

        'Write test monthly salary value from the Sensitivity Analysis Test Table into View Data Worksheet
        Worksheets("ViewData").Range("H7").value = Worksheets("SensAnalysis").Range("A" & (i + 7)).value
        Worksheets("ViewData").Range("I7").value = Worksheets("SensAnalysis").Range("A" & (i + 7)).value
        Worksheets("ViewData").Range("J7").value = Worksheets("SensAnalysis").Range("A" & (i + 7)).value

        'Read Best and Worst Salaries to Variable
        BestSalary = Worksheets("ViewData").Range("K7")
        WorstSalary = Worksheets("ViewData").Range("L7")

        'Recalculate Salary Values and Write to worksheet
        Worksheets("ExponValue").Range("H7") = (Exp(-(Worksheets("ViewData").Range("H7") - WorstSalary) / rho) - 1) / (Exp(-(BestSalary - WorstSalary) / rho) - 1)
        Worksheets("ExponValue").Range("I7") = (Exp(-(Worksheets("ViewData").Range("I7") - WorstSalary) / rho) - 1) / (Exp(-(BestSalary - WorstSalary) / rho) - 1)
        Worksheets("ExponValue").Range("J7") = (Exp(-(Worksheets("ViewData").Range("J7") - WorstSalary) / rho) - 1) / (Exp(-(BestSalary - WorstSalary) / rho) - 1)

        'EXPON VALUE WEIGHTED SCORE CALCULATION
        ExponValue

        'Write calculated weighted scores into Sensitivity Analysis Table.
        Worksheets("SensAnalysis").Range("B" & (i + 7)).value = JobAScore
        Worksheets("SensAnalysis").Range("C" & (i + 7)).value = JobBScore
        Worksheets("SensAnalysis").Range("D" & (i + 7)).value = JobCScore

        'Vary the salary for next calculation.
        NextSalary = NextSalary + 500
    Next

    'Read the original salary back onto the View Data worksheet.
    Worksheets("ViewData").Range("H7") = tempA1
    Worksheets("ViewData").Range("I7") = tempB1
    Worksheets("ViewData").Range("J7") = tempC1

    'Read the original recalculated salary back onto the Expon Value worksheet.
    Worksheets("ExponValue").Range("H7") = tempA2
    Worksheets("ExponValue").Range("I7") = tempB2
    Worksheets("ExponValue").Range("J7") = tempC2

    SensAnalysisTest = False
End Sub

Private Sub SalaryImportanceSensAnalysis()
    Dim temp As Double
    Dim i As Integer
    Dim StartImportance, NextImportance As Double

    SensAnalysisTest = True

    'Store original salary scale value.
    temp = Worksheets("ViewRanks").Range("O6")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("A30").value = "Salary Importance"
    Worksheets("SensAnalysis").Range("B30").value = "Job A"
    Worksheets("SensAnalysis").Range("C30").value = "Job B"
    Worksheets("SensAnalysis").Range("D30").value = "Job C"

    'Set variable importance ratings.
    StartImportance = 0
    NextImportance = StartImportance

    'Fill in Sensitivity Analysis Table for 0 to 100.
    For i = 0 To 20
        Worksheets("SensAnalysis").Range("A" & (i + 31)).value = NextImportance

        'Write test salary scale value from the Sensitivity Analysis Test Table.
        Worksheets("ViewRanks").Range("O6") = Worksheets("SensAnalysis").Range("A" & (i + 31)).value

        'EXPON VALUE WEIGHTED SCORE CALCULATION
        ExponValue

        'Write calculated weighted scores into Sensitivity Analysis Table.
        Worksheets("SensAnalysis").Range("B" & (i + 31)).value = JobAScore
        Worksheets("SensAnalysis").Range("C" & (i + 31)).value = JobBScore
        Worksheets("SensAnalysis").Range("D" & (i + 31)).value = JobCScore

        'Vary the salary scale value for next calculation.
        NextImportance = NextImportance + 5
    Next

    'Read the original salary scale value back itno the textbox.
    Worksheets("ViewRanks").Range("O6") = temp

    SensAnalysisTest = False
End Sub

Private Sub LocationImportanceSensAnalysis()
    Dim temp As Double
    Dim i As Integer
    Dim StartImportance, NextImportance As Double

    SensAnalysisTest = True

    'Store original location scale value.
    temp = Worksheets("ViewRanks").Range("O7")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("A55").value = "Location Importance"
    Worksheets("SensAnalysis").Range("B55").value = "Job A"
    Worksheets("SensAnalysis").Range("C55").value = "Job B"
    Worksheets("SensAnalysis").Range("D55").value = "Job C"

    'Set variable importance rating.
    StartImportance = 0
    NextImportance = StartImportance

    'Fill in Sensitivity Analysis Table for 0 to 100.
    For i = 0 To 20
        Worksheets("SensAnalysis").Range("A" & (i + 56)).value = NextImportance

        'Read test location scale value from the Sensitivity Analysis Test Table.
        Worksheets("ViewRanks").Range("O7") = Worksheets("SensAnalysis").Range("A" & (i + 56)).value

        'EXPON VALUE WEIGHTED SCORE CALCULATION
        ExponValue

        'Write calculated weighted scores into Sensitivity Analysis Table.
        Worksheets("SensAnalysis").Range("B" & (i + 56)).value = JobAScore
        Worksheets("SensAnalysis").Range("C" & (i + 56)).value = JobBScore
        Worksheets("SensAnalysis").Range("D" & (i + 56)).value = JobCScore

        'Vary the location scale value for next calculation.
        NextImportance = NextImportance + 5
    Next

    'Read the original location scale value back into the textbox.
    Worksheets("ViewRanks").Range("O7") = temp

    SensAnalysisTest = False
End Sub

Private Sub EnjoymentImportanceSensAnalysis()
    Dim temp As Double
    Dim i As Integer
    Dim StartImportance, NextImportance As Double

    SensAnalysisTest = True

    'Store original enjoyment scale value.
    temp = Worksheets("ViewRanks").Range("O8")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("A80").value = "Enjoyment Importance"
    Worksheets("SensAnalysis").Range("B80").value = "Job A"
    Worksheets("SensAnalysis").Range("C80").value = "Job B"
    Worksheets("SensAnalysis").Range("D80").value = "Job C"

    'Set variable importance rating.
    StartImportance = 0
    NextImportance = StartImportance

    'Fill in Sensitivity Analysis Table for 0 to 100.
    For i = 0 To 20
        Worksheets("SensAnalysis").Range("A" & (i + 81)).value = NextImportance

        'Read test enjoyment scale value from the Sensitivity Analysis Test Table.
        Worksheets("ViewRanks").Range("O8") = Worksheets("SensAnalysis").Range("A" & (i + 81)).value

        'EXPON VALUE WEIGHTED SCORE CALCULATION
        ExponValue

        'Write calculated weighted scores into Sensitivity Analysis Table.
        Worksheets("SensAnalysis").Range("B" & (i + 81)).value = JobAScore
        Worksheets("SensAnalysis").Range("C" & (i + 81)).value = JobBScore
        Worksheets("SensAnalysis").Range("D" & (i + 81)).value = JobCScore

        'Vary the enjoyment scale value for next calculation.
        NextImportance = NextImportance + 5
    Next

    'Read the original enjoyment scale value back itno the textbox.
    Worksheets("ViewRanks").Range("O8") = temp

    SensAnalysisTest = False
End Sub

Private Sub ExponValue()
    Dim SalaryScale, LocationScale, EnjoymentScale As Double
    Dim SalaryWeight, LocationWeight, EnjoymentWeight As Double

    'Write Values to Variables
    SalaryScale = Worksheets("ViewRanks").Range("O6")
    LocationScale = Worksheets("ViewRanks").Range("O7")
    EnjoymentScale = Worksheets("ViewRanks").Range("O8")

    'Calculate Weights
    SalaryWeight = SalaryScale / (SalaryScale + LocationScale + EnjoymentScale)
    LocationWeight = LocationScale / (SalaryScale + LocationScale + EnjoymentScale)
    EnjoymentWeight = EnjoymentScale / (SalaryScale + LocationScale + EnjoymentScale)

    'Calulcate Weighted Scores
    If Round(SalaryWeight + LocationWeight + EnjoymentWeight, 0) = 1 Then
        JobAScore = Worksheets("ExponValue").Range("H7") * SalaryWeight _
                  + Worksheets("ExponValue").Range("H8") * LocationWeight _
                  + Worksheets("ExponValue").Range("H9") * EnjoymentWeight
        JobBScore = Worksheets("ExponValue").Range("I7") * SalaryWeight _
                  + Worksheets("ExponValue").Range("I8") * LocationWeight _
                  + Worksheets("ExponValue").Range("I9") * EnjoymentWeight
        JobCScore = Worksheets("ExponValue").Range("J7") * SalaryWeight _
                  + Worksheets("ExponValue").Range("J8") * LocationWeight _
                  + Worksheets("ExponValue").Range("J9") * EnjoymentWeight
    Else
        MsgBox ("Weights must sum to 1.")
    End If
End Sub

Private Sub CommandButton9_Click()
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
    ElseIf ComboNavigation = "Rate Objectives" Then
        Worksheets("ObjectiveWeights").Activate
    ElseIf ComboNavigation = "View Rankings" Then
        Worksheets("ViewRanks").Activate
    Else
        MsgBox ("Please select an option from the list before pressing go.")
    End If
End Sub
