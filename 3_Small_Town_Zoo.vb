Option Explicit
Dim Email As Double
Dim EmailCoupon As Double
Dim ReturnOwn As Double
Dim ReturnNoCoupon As Double
Dim ReturnCoupon As Double
Dim CouponCost As Double
Dim EmailCost As Double
Dim RevReturnOwn As Double
Dim RevReturnNoCoupon As Double
Dim RevReturnCoupon As Double
Dim NumVisitors As Integer
Public SensAnalysisTest As Boolean
Public ProfitPromotion As Double
Public ProfitNoPromotion As Double

Private Sub CommandButton1_Click()
    'Calculate expected net profit, depending on the data source.
    SensAnalysisTest = False
    CalculateProfit
End Sub

Private Sub CommandButton2_Click()
    RunSensAnalysis
End Sub

Private Sub CommandButton4_Click()
    'User Form for Data Inputs
    DataEntryForm.Show
End Sub

Private Sub CommandButton5_Click()
    'Open Data File
    ImportFile
End Sub

Private Sub CommandButton6_Click()
    'Reset to default values
    Worksheets("Inputs").Range("D2").Value = 1000
    Worksheets("Inputs").Range("D3").Value = Format(0.4, "##0%")
    Worksheets("Inputs").Range("D4").Value = Format(0.5, "##0%")
    Worksheets("Inputs").Range("D5").Value = Format(0.2, "##0%")
    Worksheets("Inputs").Range("D6").Value = Format(0.3, "##0%")
    Worksheets("Inputs").Range("D7").Value = Format(0.4, "##0%")
    Worksheets("Inputs").Range("D8").Value = Format(2, "$##,##0")
    Worksheets("Inputs").Range("D9").Value = Format(200, "$##,##0")
    Worksheets("Inputs").Range("D10").Value = Format(15, "$##,##0")
    Worksheets("Inputs").Range("D11").Value = Format(13, "$##,##0")
    Worksheets("Inputs").Range("D12").Value = Format(9, "$##,##0")
End Sub

Private Sub CalculateProfit()
'Read variables from the worksheet into the code.
NumVisitors = Worksheets("Inputs").Range("D2").Value
Email = Worksheets("Inputs").Range("D3").Value
EmailCoupon = Worksheets("Inputs").Range("D4").Value
ReturnOwn = Worksheets("Inputs").Range("D5").Value
ReturnNoCoupon = Worksheets("Inputs").Range("D6").Value
ReturnCoupon = Worksheets("Inputs").Range("D7").Value
CouponCost = Worksheets("Inputs").Range("D8").Value
EmailCost = Worksheets("Inputs").Range("D9").Value
RevReturnOwn = Worksheets("Inputs").Range("D10").Value
RevReturnNoCoupon = Worksheets("Inputs").Range("D11").Value
RevReturnCoupon = Worksheets("Inputs").Range("D12").Value

'Data Validation - Percentages must be between 0% and 100%.
If NumVisitors >= 0 And NumVisitors <= 10000 Then
    If Email >= 0 And Email <= 1 Then
        If EmailCoupon >= 0 And EmailCoupon <= 1 Then
            If ReturnOwn >= 0 And ReturnOwn <= 1 Then
                If ReturnNoCoupon >= 0 And ReturnNoCoupon <= 1 Then
                    If ReturnCoupon >= 0 And ReturnCoupon <= 1 Then
                        If CouponCost >= 0 And CouponCost <= 1000 Then
                            If EmailCost >= 0 And EmailCost <= 10000 Then
                                If RevReturnOwn >= 0 And RevReturnOwn <= 1000 Then
                                    If RevReturnNoCoupon >= 0 And RevReturnNoCoupon <= 1000 Then
                                        If RevReturnCoupon >= 0 And RevReturnCoupon <= 1000 Then
                                            'Profit Calculation with Promotion = Revenue.No Email--Return + Revenue.Email--NoCoupon--Return + Profit.Email--Coupon--Return - EmailCost
                                            ProfitPromotion = NumVisitors * (1 - Email) * ReturnOwn * RevReturnOwn _
                                                + NumVisitors * Email * (1 - EmailCoupon) * ReturnNoCoupon * RevReturnNoCoupon _
                                                + NumVisitors * Email * EmailCoupon * ReturnCoupon * (RevReturnCoupon - CouponCost) _
                                                - EmailCost

                                            'Profit Calculation without Promotion
                                            ProfitNoPromotion = NumVisitors * ReturnOwn * RevReturnOwn

                                            'Only run this chunk of code when user is not running sensitivity analysis.
                                            If Not (SensAnalysisTest) Then
                                                'Recommendation Message Box - Email or Not?
                                                If ProfitPromotion > ProfitNoPromotion Then
                                                    'Email Recommendation
                                                    MsgBox ("Expected profit with email: " & Format(ProfitPromotion, "$##,##0") _
                                                        & vbLf & "Expected profit without email: " _
                                                        & Format(ProfitNoPromotion, "$##,##0") & vbLf & vbLf _
                                                        & "Recommendation: Small Town Zoo should send the promotion email.") _
                                                        , , "Recommendation"
                                                ElseIf ProfitPromotion < ProfitNoPromotion Then
                                                    'No Email Recommendation
                                                    MsgBox ("Expected profit with email: " & Format(ProfitPromotion, "$##,##0") _
                                                        & vbLf & "Expected profit without email: " _
                                                        & Format(ProfitNoPromotion, "$##,##0") & vbLf & vbLf _
                                                        & "Recommendation: Small Town Zoo should not send the promotion email.") _
                                                        , , "Recommendation"
                                                Else
                                                    'Same Profit
                                                    MsgBox ("There is no difference between sending or not sending the promotion email. Expected profit with or without email is: " & Format(ProfitPromotion, "$##,##0"))
                                                End If
                                            End If
                                        Else
                                            MsgBox ("Average revenue per person returning with coupon needs to be between $0 and $1000.")
                                        End If
                                    Else
                                        MsgBox ("Average revenue per person returning without a coupon needs to be between $0 and $1000.")
                                    End If
                                Else
                                    MsgBox ("Average revenue per person returning on their own needs to be between $0 and $1000.")
                                End If
                            Else
                                MsgBox ("Cost per email campaign needs to be between $0 and $10,000.")
                            End If
                        Else
                            MsgBox ("Cost per coupon needs to be between $0 and $1000.")
                        End If
                    Else
                        MsgBox ("% of Email Recipients With Coupon Who Return must be between 0% and 100%.")
                    End If
                Else
                    MsgBox ("% of Email Recipients Without Coupon Who Return must be between 0% and 100%.")
                End If
            Else
                MsgBox ("% of Visitors Who Return On Their Own must be between 0% and 100%.")
            End If
        Else
            MsgBox ("% of Email Recipients that Receive Coupon must be between 0% and 100%.")
        End If
    Else
        MsgBox ("% of First-Time Visitors Emailed must be between 0% and 100%.")
    End If
Else
    MsgBox ("Number of visitors needs to be between 0 and 10,000 visitors.")
End If

End Sub

Private Sub EmailSensAnalysis()
    Dim temp As Double
    Dim i As Integer
    Dim StartPercent As Double
    Dim NextPercent As Double

    SensAnalysisTest = True

    'Store original Email value
    temp = Worksheets("Inputs").Range("D3")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("A34").Value = "Email"
    Worksheets("SensAnalysis").Range("B34").Value = "Email Profit"
    Worksheets("SensAnalysis").Range("C34").Value = "No Email Profit"

    'Set variable percentages.
    StartPercent = 0
    NextPercent = StartPercent

    'Fill in Sensitivity Analysis Table for 0% to 100%.
    For i = 1 To 21
        Worksheets("SensAnalysis").Range("A" & (i + 34)).Value = Format(NextPercent, "##0%")

        'Read off test Email value from the Sensitivity Analysis test table.
        Worksheets("Inputs").Range("D3").Value = Worksheets("SensAnalysis").Range("A" & (i + 34)).Value

        'Use test Email value in profit formula.
        CalculateProfit

        'Write calculated profit for email and no email to sensitivity analysis table.
        Worksheets("SensAnalysis").Range("B" & (i + 34)).Value = Format(ProfitPromotion, "$##,##0")
        Worksheets("SensAnalysis").Range("C" & (i + 34)).Value = Format(ProfitNoPromotion, "$##,##0")

        'Vary the percentage for next calculation.
        NextPercent = NextPercent + 0.05
    Next

    'Read the original Email value back onto the worksheet.
    Worksheets("Inputs").Range("D3") = temp
End Sub

Private Sub EmailCouponSensAnalysis()
    Dim temp As Double
    Dim i As Integer
    Dim StartPercent As Double
    Dim NextPercent As Double

    SensAnalysisTest = True

    'Store original EmailCoupon value
    temp = Worksheets("Inputs").Range("D4")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("E34").Value = "EmailCoupon"
    Worksheets("SensAnalysis").Range("F34").Value = "Email Profit"
    Worksheets("SensAnalysis").Range("G34").Value = "No Email Profit"

    'Set variable percentages.
    StartPercent = 0
    NextPercent = StartPercent

    'Fill in Sensitivity Analysis Table for 0% to 100%.
    For i = 1 To 21
        Worksheets("SensAnalysis").Range("E" & (i + 34)).Value = Format(NextPercent, "##0%")

        'Read off test EmailCoupon value from the Sensitivity Analysis test table.
        Worksheets("Inputs").Range("D4").Value = Worksheets("SensAnalysis").Range("E" & (i + 34)).Value

        'Use test EmailCoupon value in profit formula.
        CalculateProfit

        'Write calculated profit for email and no email to sensitivity analysis table.
        Worksheets("SensAnalysis").Range("F" & (i + 34)).Value = Format(ProfitPromotion, "$##,##0")
        Worksheets("SensAnalysis").Range("G" & (i + 34)).Value = Format(ProfitNoPromotion, "$##,##0")

        'Vary the percentage for next calculation.
        NextPercent = NextPercent + 0.05
    Next

    'Read the original EmailCoupon value back onto the worksheet.
    Worksheets("Inputs").Range("D4") = temp
End Sub
Private Sub ReturnOwnSensAnalysis()
    Dim temp As Double
    Dim i As Integer
    Dim StartPercent As Double
    Dim NextPercent As Double

    SensAnalysisTest = True

    'Store original ReturnOwn value
    temp = Worksheets("Inputs").Range("D5")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("I34").Value = "ReturnOwn"
    Worksheets("SensAnalysis").Range("J34").Value = "Email Profit"
    Worksheets("SensAnalysis").Range("K34").Value = "No Email Profit"

    'Set variable percentages.
    StartPercent = 0
    NextPercent = StartPercent

    'Fill in Sensitivity Analysis Table for 0% to 100%.
    For i = 1 To 21
        Worksheets("SensAnalysis").Range("I" & (i + 34)).Value = Format(NextPercent, "##0%")

        'Read off test ReturnOwn value from the Sensitivity Analysis test table.
        Worksheets("Inputs").Range("D5").Value = Worksheets("SensAnalysis").Range("I" & (i + 34)).Value

        'Use test ReturnOwn value in profit formula.
        CalculateProfit

        'Write calculated profit for email and no email to sensitivity analysis table.
        Worksheets("SensAnalysis").Range("J" & (i + 34)).Value = Format(ProfitPromotion, "$##,##0")
        Worksheets("SensAnalysis").Range("K" & (i + 34)).Value = Format(ProfitNoPromotion, "$##,##0")

        'Vary the percentage for next calculation.
        NextPercent = NextPercent + 0.05
    Next

    'Read the original ReturnOwn value back onto the worksheet.
    Worksheets("Inputs").Range("D5") = temp
End Sub

Private Sub ReturnNoCouponSensAnalysis()
    Dim temp As Double
    Dim i As Integer
    Dim StartPercent As Double
    Dim NextPercent As Double

    SensAnalysisTest = True

    'Store original ReturnNoCoupon value
    temp = Worksheets("Inputs").Range("D6")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("M34").Value = "ReturnNoCoupon"
    Worksheets("SensAnalysis").Range("N34").Value = "Email Profit"
    Worksheets("SensAnalysis").Range("O34").Value = "No Email Profit"

    'Set variable percentages.
    StartPercent = 0
    NextPercent = StartPercent

    'Fill in Sensitivity Analysis Table for 0% to 100%.
    For i = 1 To 21
        Worksheets("SensAnalysis").Range("M" & (i + 34)).Value = Format(NextPercent, "##0%")

        'Read off test ReturnNoCoupon value from the Sensitivity Analysis test table.
        Worksheets("Inputs").Range("D6").Value = Worksheets("SensAnalysis").Range("M" & (i + 34)).Value

        'Use test ReturnNoCoupon value in profit formula.
        CalculateProfit

        'Write calculated profit for email and no email to sensitivity analysis table.
        Worksheets("SensAnalysis").Range("N" & (i + 34)).Value = Format(ProfitPromotion, "$##,##0")
        Worksheets("SensAnalysis").Range("O" & (i + 34)).Value = Format(ProfitNoPromotion, "$##,##0")

        'Vary the percentage for next calculation.
        NextPercent = NextPercent + 0.05
    Next

    'Read the original ReturnNoCoupon value back onto the worksheet.
    Worksheets("Inputs").Range("D6") = temp
End Sub

Private Sub ReturnCouponSensAnalysis()
    Dim temp As Double
    Dim i As Integer
    Dim StartPercent As Double
    Dim NextPercent As Double

    SensAnalysisTest = True

    'Store original ReturnCoupon value
    temp = Worksheets("Inputs").Range("D7")

    'Sensitivity Table Labels
    Worksheets("SensAnalysis").Range("Q34").Value = "ReturnCoupon"
    Worksheets("SensAnalysis").Range("R34").Value = "Email Profit"
    Worksheets("SensAnalysis").Range("S34").Value = "No Email Profit"

    'Set variable percentages.
    StartPercent = 0
    NextPercent = StartPercent

    'Fill in Sensitivity Analysis Table for 0% to 100%.
    For i = 1 To 21
        Worksheets("SensAnalysis").Range("Q" & (i + 34)).Value = Format(NextPercent, "##0%")

        'Read off test ReturnCoupon value from the Sensitivity Analysis test table.
        Worksheets("Inputs").Range("D7").Value = Worksheets("SensAnalysis").Range("Q" & (i + 34)).Value

        'Use test ReturnCoupon value in profit formula.
        CalculateProfit

        'Write calculated profit for email and no email to sensitivity analysis table.
        Worksheets("SensAnalysis").Range("R" & (i + 34)).Value = Format(ProfitPromotion, "$##,##0")
        Worksheets("SensAnalysis").Range("S" & (i + 34)).Value = Format(ProfitNoPromotion, "$##,##0")

        'Vary the percentage for next calculation.
        NextPercent = NextPercent + 0.05
    Next

    'Read the original ReturnCoupon value back onto the worksheet.
    Worksheets("Inputs").Range("D7") = temp
End Sub

Private Sub RunSensAnalysis()
    'Run and View Sensitivity Analysis
    Application.ScreenUpdating = False

    'Call Sensitivity Analysis Sub Functions
    EmailSensAnalysis
    EmailCouponSensAnalysis
    ReturnOwnSensAnalysis
    ReturnNoCouponSensAnalysis
    ReturnCouponSensAnalysis

    'Resize graphs on SensAnalysis sheet.
    Worksheets("SensAnalysis").Activate
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveSheet.ChartObjects("Chart 1").Activate
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes.Range(Array("Chart 1", "Chart 2")).Select
    ActiveSheet.Shapes.Range(Array("Chart 1", "Chart 2", "Chart 3")).Select
    ActiveSheet.Shapes.Range(Array("Chart 1", "Chart 2", "Chart 3", "Chart 4")).Select
    ActiveSheet.Shapes.Range(Array("Chart 1", "Chart 2", "Chart 3", "Chart 4", "Chart 5")).Select
    Selection.ShapeRange.Height = 216
    Selection.ShapeRange.Width = 360
    ActiveSheet.Range("C21").Select

    'Return to Inputs worksheet view.
    Worksheets("Inputs").Activate

    Application.ScreenUpdating = True

    'Too much code, chart won't automatically update until you press next. - Spoke with Prof. Montano about this issue, but no clear resolution.
    MsgBox ("To view the sensitivity analysis graphs, please press okay and then press the NEXT button in the user form to continue.")

    'Pop up sensivitiy analysis graphs.
    SensAnalysisGraphs.Show
End Sub
