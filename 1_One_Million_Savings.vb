'Assignment #1
'Inputs: Salary, Current Age, Contribution %, Salary Growth Rate, Investment Return Rate
'Calculate: How long until you're a millionaire?

Private Sub CommandButton1_Click()

Dim Balance As Double
Dim CurrentAge As Integer
Dim StartSalary As Double
Dim Salary As Double
Dim PercentContribution As Double
Dim GrowthRate As Double
Dim ReturnRate As Double
Dim TaxRate As Double
Dim Years As Long

'Read variables from the worksheet into the code.
StartSalary = Worksheets("Sheet1").Range("C4").Value
CurrentAge = Worksheets("Sheet1").Range("C6").Value
PercentContribution = Worksheets("Sheet1").Range("C8").Value
GrowthRate = Worksheets("Sheet1").Range("G4").Value
ReturnRate = Worksheets("Sheet1").Range("G6").Value

'Check input values (data validation for above variables).
If CurrentAge > 0 And CurrentAge < 101 Then
    If StartSalary > 0 Then
        If PercentContribution > 0 And PercentContribution <= 1 Then
            If GrowthRate >= 0 And GrowthRate <= 0.1 Then
                If ReturnRate >= -0.1 And ReturnRate <= 0.4 Then

                    Years = 1                                       'Year 1 is the first time you will be making a contribution.
                    Balance = StartSalary * PercentContribution     'In year 1, add a percentage of your salary into the account.
                    Salary = StartSalary                            'Starting salary becomes the base for salary growth each year.

                    'Set the Do-Loop to run until the balance = $1 million or jump out after 100 years. The investment and salary growth begin in year 2.
                    Do Until (Balance >= 1000000) Or (Years > 100)
                        Years = Years + 1                   'Counter variable that will add a year each time the loop is run.
                        Salary = Salary * (1 + GrowthRate)  'Calculates the new salary for the year.

                        'Based on the current salary, find the total tax rate, which is equal to the federal tax rate plus the state tax rate. Read in tax rates from worksheet.
                        Select Case Salary
                            Case Is <= 30000:           TaxRate = Worksheets("Sheet1").Range("C18").Value + Worksheets("Sheet1").Range("D18").Value
                            Case 30000.01 To 70000:     TaxRate = Worksheets("Sheet1").Range("C19").Value + Worksheets("Sheet1").Range("D19").Value
                            Case 70000.01 To 150000:    TaxRate = Worksheets("Sheet1").Range("C20").Value + Worksheets("Sheet1").Range("D20").Value
                            Case 150000.01 To 250000:   TaxRate = Worksheets("Sheet1").Range("C21").Value + Worksheets("Sheet1").Range("D21").Value
                            Case Is >= 250000.01:       TaxRate = Worksheets("Sheet1").Range("C22").Value + Worksheets("Sheet1").Range("D22").Value
                        End Select

                        'You only pay taxes if the investment income is greater than 0. Depending on the investment income, the formula will calculate the ending balance for the year.
                        If (Balance * ReturnRate) > 0 Then
                            Balance = Balance + (Balance * ReturnRate) * (1 - TaxRate) + (PercentContribution * Salary)
                        Else
                            Balance = Balance + (Balance * ReturnRate) + (PercentContribution * Salary)
                        End If
                    Loop

                    'If the number of years it takes to reach $1 million is reasonable vs. not reasonable, display different message boxes.
                    If (CurrentAge + Years) < 101 Then
                        MsgBox ("If you make " & Format(StartSalary, "Currency") & " at age " & CurrentAge & " and contribute " _
                            & Format(PercentContribution, "Percent") & " to savings each year," & vbLf & _
                            "you should expect to have $1 million in " & Years & " years when you are " & (CurrentAge + Years) & ".")
                    Else
                        MsgBox ("Given the terms, you will not reach $1 million in savings by age 100.")
                    End If
                Else
                    MsgBox ("Expected annual investment return rate must be between -10% and 40%.")
                End If
            Else
                MsgBox ("Expected annual salary growth rate must be between 0% and 10%.")
            End If
        Else
            MsgBox ("Percent of salary to contribute must be between 0% and 100%.")
        End If
    Else
        MsgBox ("Starting salary needs to be a positive value.")
    End If
Else
    MsgBox ("Current age must be positive.")
End If


End Sub
