Option Explicit
Private Sub CommandButton1_Click()
    'Continue Button >> Data Entry

    If OptionButton1 = True Or OptionButton2 = True Or OptionButton3 = True Then
        'Continue to Data Entry
        Worksheets("DataEntry").Visible = True
        Worksheets("DataEntry").Activate
    Else
        MsgBox ("Please select a model option before continuing.")
    End If

End Sub

Private Sub CommandButton2_Click()
    'Even Swap Explanation
    DecisionModelHelp.MultiPage1.Pages(0).Visible = False
    DecisionModelHelp.MultiPage1.Pages(2).Visible = False
    DecisionModelHelp.MultiPage1.Pages(3).Visible = False
    DecisionModelHelp.MultiPage1.Pages(4).Visible = False

    DecisionModelHelp.MultiPage1.Pages(1).Visible = True
    DecisionModelHelp.MultiPage1.value = 1

    DecisionModelHelp.Show
End Sub

Private Sub CommandButton3_Click()
    'Linear Model Explanation
    DecisionModelHelp.MultiPage1.Pages(0).Visible = False
    DecisionModelHelp.MultiPage1.Pages(1).Visible = False
    DecisionModelHelp.MultiPage1.Pages(3).Visible = False
    DecisionModelHelp.MultiPage1.Pages(4).Visible = False

    DecisionModelHelp.MultiPage1.Pages(2).Visible = True
    DecisionModelHelp.MultiPage1.value = 2

    DecisionModelHelp.Show
End Sub

Private Sub CommandButton4_Click()
    'Exponential Model Explanation
    DecisionModelHelp.MultiPage1.Pages(0).Visible = False
    DecisionModelHelp.MultiPage1.Pages(1).Visible = False
    DecisionModelHelp.MultiPage1.Pages(2).Visible = False
    DecisionModelHelp.MultiPage1.Pages(4).Visible = False

    DecisionModelHelp.MultiPage1.Pages(3).Visible = True
    DecisionModelHelp.MultiPage1.value = 3

    DecisionModelHelp.Show
End Sub
