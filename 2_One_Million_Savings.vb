Option Explicit

Public Function Triang(a, b, c) As Double       'Min, Most Likely, Max Value (a, b, c)
'Custom Triang function

Dim URandom As Double

URandom = Rnd()
If URandom <= (b - a) / (c - a) Then
    Triang = a + SquareRoot(URandom * (c - a) * (b - a))
Else
    Triang = c - SquareRoot((1 - URandom) * (c - a) * (c - b))
End If

End Function

Public Function SquareRoot(number) As Double
  SquareRoot = number ^ (1 / 2)
End Function

Private Sub CheckBox1_Click()

'If the CheckBox is checked, show the list box that will show investment detail.
If CheckBox1.Value = True Then
    Worksheets("Investment Data").Range("A1").CurrentRegion.ClearContents
    Worksheets("Input").Range("J2") = "Savings Detail:"
    Sheet1.ListBox1.Visible = True
Else
    Worksheets("Input").Range("J2").ClearContents
    Sheet1.ListBox1.Visible = False
End If

End Sub

Private Sub CommandButton1_Click()

Dim Balance As Double
Dim CurrentAge As Integer
Dim StartSalary As Double
Dim Salary As Double
Dim SalaryContributionAmt As Double
Dim PercentContribution As Double
Dim GrowthRate As Double
Dim MostTriang As Double
Dim MinTriang As Double
Dim MaxTriang As Double
Dim InvestmentIncome As Double
Dim TaxRate As Double
Dim TotalTax As Double
Dim Years As Long

'Read variables from the worksheet into the code.
StartSalary = Worksheets("Input").Range("C4").Value
CurrentAge = Worksheets("Input").Range("C6").Value
PercentContribution = Worksheets("Input").Range("C8").Value
GrowthRate = Worksheets("Input").Range("G4").Value
MostTriang = Worksheets("Input").Range("G7").Value
MinTriang = Worksheets("Input").Range("G8").Value
MaxTriang = Worksheets("Input").Range("G9").Value

'Reset balance and years to 0 for each new run of code.
Balance = 0
Years = 0

'Clear out list box and investment data sheet.
Worksheets("Investment Data").Range("A1").CurrentRegion.ClearContents

'Data Validation
If CurrentAge > 11 And CurrentAge < 101 Then
    If StartSalary > 0 And StartSalary <= 500000 Then
        If PercentContribution >= 0 And PercentContribution <= 1 Then
            If GrowthRate >= 0 And GrowthRate <= 0.2 Then
                If MostTriang >= -0.1 And MostTriang <= 0.4 Then
                    If OptionButton4.Value = True Then
                        'Only check data validation for min/max triang if they decide to use triangular distribution.
                        'Data Validation: Min/Max Triang & Most Triang must be between MinTriang and MaxTriang
                        If MinTriang >= -0.1 And MinTriang <= 0.4 Then
                            If MaxTriang >= -0.1 And MaxTriang <= 0.4 Then
                                If MinTriang >= MostTriang Or MostTriang >= MaxTriang Then
                                    MsgBox ("The most likely return rate must be between the minimum return rate and the maximum return rate.")
                                    Exit Sub
                                End If
                            Else
                                MsgBox ("The maximum expected annual investment return rate must be between -10% and 40%.")
                            End If
                        Else
                            MsgBox ("The minimum expected annual investment return rate must be between -10% and 40%.")
                        End If
                    End If

                    'If CheckBox is checked to display contribution detail, set up investment data worksheet.
                    If CheckBox1.Value = True Then
                        Application.ScreenUpdating = False

                        'Shows investment data detail in new sheet (Investment Data): input data
                        Worksheets("Investment Data").Range("A1") = "Current Age"
                        Worksheets("Investment Data").Range("B1") = CurrentAge
                        Worksheets("Investment Data").Range("A2") = "Salary"
                        Worksheets("Investment Data").Range("B2") = Format(StartSalary, "Currency")
                        Worksheets("Investment Data").Range("A3") = "Contribution Rate"
                        Worksheets("Investment Data").Range("B3") = Format(PercentContribution, "Percent")

                        'Column labels for investment data detail in new sheet (Investment Data)
                        Worksheets("Investment Data").Range("A4") = "Age"
                        Worksheets("Investment Data").Range("B4") = "Contribution"
                        Worksheets("Investment Data").Range("C4") = "Investment Income"
                        Worksheets("Investment Data").Range("D4") = "Loss to Taxes"
                        Worksheets("Investment Data").Range("E4") = "Year-End Balance"
                    End If

                    'Calculation Formula
                    Salary = StartSalary
                    SalaryContributionAmt = PercentContribution * Salary
                    Balance = Balance + InvestmentIncome - TotalTax + SalaryContributionAmt

                    Do Until (Balance >= 1000000) Or (Years > 100)
                        'If CheckBox is checked, fill in investment data detail.
                        If CheckBox1.Value = True Then
                            'Fills in first row in table, year 0 at starting age for the first loop.
                            'Fills in next row in table, next year for each continuous loop.
                            Worksheets("Investment Data").Range("A" & 5 + Years).Value = "Age " & (CurrentAge + Years)
                            Worksheets("Investment Data").Range("B" & 5 + Years).Value = Format(SalaryContributionAmt, "Currency")
                            Worksheets("Investment Data").Range("C" & 5 + Years).Value = Format(InvestmentIncome, "Currency")
                            Worksheets("Investment Data").Range("D" & 5 + Years).Value = Format(TotalTax, "Currency")
                            Worksheets("Investment Data").Range("E" & 5 + Years).Value = Format(Balance, "Currency")
                        End If

                        'Calculation Formula
                        Salary = Salary * (1 + GrowthRate)

                        'TAX RATE: Tax Free Option - Yes(False) or No(True)
                        If OptionButton2.Value = True Then
                            Select Case Salary
                                Case 0 To 30000:            TaxRate = Worksheets("Input").Range("C18").Value + Worksheets("Input").Range("D18").Value
                                Case 30000.01 To 70000:     TaxRate = Worksheets("Input").Range("C19").Value + Worksheets("Input").Range("D19").Value
                                Case 70000.01 To 150000:    TaxRate = Worksheets("Input").Range("C20").Value + Worksheets("Input").Range("D20").Value
                                Case 150000.01 To 250000:   TaxRate = Worksheets("Input").Range("C21").Value + Worksheets("Input").Range("D21").Value
                                Case Is > 250000:           TaxRate = Worksheets("Input").Range("C22").Value + Worksheets("Input").Range("D22").Value
                            End Select
                        Else
                            TaxRate = 0         'Tax Free Option
                        End If

                        'GROWTH RATE: Use Most Likely Rate (OptionButton3 = True) or Triang Function
                        If OptionButton3.Value = True Then
                            InvestmentIncome = Balance * MostTriang
                        Else
                            InvestmentIncome = Balance * Triang(MinTriang, MostTriang, MaxTriang)
                        End If

                        'TAX AMT
                        If InvestmentIncome > 0 Then
                            TotalTax = InvestmentIncome * TaxRate
                        Else
                            TotalTax = 0
                        End If

                        'Calculation Formula
                        SalaryContributionAmt = PercentContribution * Salary
                        Balance = Balance + InvestmentIncome - TotalTax + SalaryContributionAmt
                        Years = Years + 1
                    Loop

                    'If CheckBox checked, fill last row of data in Investment Data Table once you've saved at least $1,000,000.
                    If CheckBox1.Value = True Then
                        Worksheets("Investment Data").Range("A" & 5 + Years).Value = "Age " & (CurrentAge + Years)
                        Worksheets("Investment Data").Range("B" & 5 + Years).Value = Format(SalaryContributionAmt, "Currency")
                        Worksheets("Investment Data").Range("C" & 5 + Years).Value = Format(InvestmentIncome, "Currency")
                        Worksheets("Investment Data").Range("D" & 5 + Years).Value = Format(TotalTax, "Currency")
                        Worksheets("Investment Data").Range("E" & 5 + Years).Value = Format(Balance, "Currency")

                        'Resize Columns to AutoFit
                        Worksheets("Investment Data").Cells.EntireColumn.AutoFit

                        'Scrolls back up to top of Investment Data Table
                        Worksheets("Investment Data").Select
                        Worksheets("Investment Data").Range("A1").Select
                    End If

                    'Returns view to Input Worksheet.
                    Worksheets("Input").Activate

                    'Message Box displayed.
                    If (CurrentAge + Years) < 101 Then
                        MsgBox ("If you make " & Format(StartSalary, "Currency") & " at age " & CurrentAge & " and contribute " _
                            & Format(PercentContribution, "Percent") & " to savings each year," & vbLf & _
                            "you should expect to have $1 million in " & Years & " years when you are " & (CurrentAge + Years) & ".")
                    Else
                        'If CheckBox checked, delete list box contents before showing message box.
                        If CheckBox1.Value = True Then
                            Worksheets("Investment Data").Range("A1").CurrentRegion.ClearContents
                        End If
                        MsgBox ("Given the terms, you will not reach $1 million in savings by age 100.")
                    End If
                Else
                    MsgBox ("The most likely expected annual investment return rate must be between -10% and 40%.")
                End If
            Else
                MsgBox ("Expected annual salary growth rate must be between 0% and 20%.")
            End If
        Else
            MsgBox ("Percent of salary to contribute must be between 0% and 100%.")
        End If
    Else
        MsgBox ("Starting salary needs to be a positive number less than $500,000.")
    End If
Else
    MsgBox ("You must be at least 12 years old.")
End If

Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

Dim temp As Double

'Salary
If Target.Address = "$C$4" Then
    temp = Range("$C$4").Value

    If temp >= 0 And temp <= 500000 Then
        Application.EnableEvents = False
        ScrollBar1.Value = temp
        Application.EnableEvents = True
    Else
        Application.EnableEvents = False
        'Force scroll bar to min if invalid number is entered.
        If temp < 0 Then
            ScrollBar1.Value = 0
        End If

        'Force scroll bar to max if invalid number is entered.
        If temp > 500000 Then
            ScrollBar1.Value = 500000
        End If

        Application.EnableEvents = True
    End If
End If

'Age
If Target.Address = "$C$6" Then
    temp = Range("$C$6").Value

    If temp >= 12 And temp <= 100 Then
        Application.EnableEvents = False
        ScrollBar2.Value = temp
        Application.EnableEvents = True
    Else
        Application.EnableEvents = False
        'Force scroll bar to min if invalid number is entered.
        If temp < 12 Then
            ScrollBar2.Value = 12
        End If

        'Force scroll bar to max if invalid number is entered.
        If temp > 100 Then
            ScrollBar2.Value = 100
        End If

        Application.EnableEvents = True
    End If
End If

'Percent of Salary to Contribute - Linked to Cell P50
If Target.Address = "$C$8" Then
    temp = Range("$C$8").Value

    'Replace formula when a number is entered into the cell.
    If temp >= 0 And temp <= 1 Then
        Application.EnableEvents = False
        ScrollBar3.Value = (temp * 1000)
        Range("$C$8").Formula = "=P50/1000"
        Application.EnableEvents = True
    Else
        Application.EnableEvents = False
        'Force scroll bar to min if invalid number is entered.
        If temp < 0 Then
            ScrollBar3.Value = 0
            Range("$C$8").Formula = "=P50/1000"
        End If

        'Force scroll bar to max if invalid number is entered.
        If temp > 1 Then
            ScrollBar3.Value = 1000
            Range("$C$8").Formula = "=P50/1000"
        End If

        Application.EnableEvents = True
    End If
End If

'Salary Growth Rate - Linked to Cell P51
If Target.Address = "$G$4" Then
    temp = Range("$G$4").Value

    'Replace formula when a number is entered into the cell.
    If temp >= 0 And temp <= 0.2 Then
        Application.EnableEvents = False
        ScrollBar4.Value = (temp * 1000)
        Range("$G$4").Formula = "=P51/1000"
        Application.EnableEvents = True
    Else
        Application.EnableEvents = False
        'Force scroll bar to min if invalid number is entered.
        If temp < 0 Then
            ScrollBar4.Value = 0
            Range("$G$4").Formula = "=P51/1000"
        End If

        'Force scroll bar to max if invalid number is entered.
        If temp > 0.2 Then
            ScrollBar4.Value = 200
            Range("$G$4").Formula = "=P51/1000"
        End If

        Application.EnableEvents = True
    End If
End If

'Most Likely Return Rate - Linked to Cell P52
If Target.Address = "$G$7" Then
    temp = Range("$G$7").Value

    'Replace formula when a number is entered into the cell.
    If temp >= -0.1 And temp <= 0.4 Then
        Application.EnableEvents = False
        ScrollBar5.Value = (temp * 1000) + 100
        Range("$G$7").Formula = "=(P52-100)/1000"
        Application.EnableEvents = True
    Else
        Application.EnableEvents = False
        'Force scroll bar to min if invalid number is entered.
        If temp < -0.1 Then
            ScrollBar5.Value = 0
            Range("$G$7").Formula = "=(P52-100)/1000"
        End If

        'Force scroll bar to max if invalid number is entered.
        If temp > 0.4 Then
            ScrollBar5.Value = 500
            Range("$G$7").Formula = "=(P52-100)/1000"
        End If

        Application.EnableEvents = True
    End If
End If

'Minimum Return Rate - Linked to Cell P53
If Target.Address = "$G$8" Then
    temp = Range("$G$8").Value

    'Replace formula when a number is entered into the cell.
    If temp >= -0.1 And temp <= 0.4 Then
        Application.EnableEvents = False
        ScrollBar6.Value = (temp * 1000) + 100
        Range("$G$8").Formula = "=(P53-100)/1000"
        Application.EnableEvents = True
    Else
        Application.EnableEvents = False
        'Force scroll bar to min if invalid number is entered.
        If temp < -0.1 Then
            ScrollBar6.Value = 0
            Range("$G$8").Formula = "=(P53-100)/1000"
        End If

        'Force scroll bar to max if invalid number is entered.
        If temp > 0.4 Then
            ScrollBar6.Value = 500
            Range("$G$8").Formula = "=(P53-100)/1000"
        End If

        Application.EnableEvents = True
    End If
End If

'Maximum Return Rate - Linked to Cell P54
If Target.Address = "$G$9" Then
    temp = Range("$G$9").Value

    'Replace formula when a number is entered into the cell.
    If temp >= -0.1 And temp <= 0.4 Then
        Application.EnableEvents = False
        ScrollBar7.Value = (temp * 1000) + 100
        Range("$G$9").Formula = "=(P54-100)/1000"
        Application.EnableEvents = True
    Else
        Application.EnableEvents = False
        'Force scroll bar to min if invalid number is entered.
        If temp < -0.1 Then
            ScrollBar7.Value = 0
            Range("$G$9").Formula = "=(P54-100)/1000"
        End If

        'Force scroll bar to max if invalid number is entered.
        If temp > 0.4 Then
            ScrollBar7.Value = 500
            Range("$G$9").Formula = "=(P54-100)/1000"
        End If

        Application.EnableEvents = True
    End If
End If

End Sub
