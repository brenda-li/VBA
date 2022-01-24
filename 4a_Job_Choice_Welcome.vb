Private Sub CommandButton1_Click()
    'Continue Button >> Select a Model
    Worksheets("Model").Visible = True
    Worksheets("Model").Activate
End Sub

Private Sub CommandButton2_Click()
    'Unhide All Worksheets

Application.ScreenUpdating = False

    Worksheets("Model").Visible = True
    Worksheets("DataEntry").Visible = True
    Worksheets("ViewData").Visible = True
    Worksheets("RankObjectives").Visible = True
    Worksheets("ObjectiveWeights").Visible = True
    Worksheets("ViewRanks").Visible = True
    Worksheets("EvenSwap").Visible = True
    Worksheets("LinearValue").Visible = True
    Worksheets("ExponValue").Visible = True
    Worksheets("SensAnalysis").Visible = True

Application.ScreenUpdating = True

End Sub

Private Sub CommandButton3_Click()
    DecisionModelHelp.MultiPage1.Pages(0).Visible = True
    DecisionModelHelp.MultiPage1.Pages(1).Visible = True
    DecisionModelHelp.MultiPage1.Pages(2).Visible = True
    DecisionModelHelp.MultiPage1.Pages(3).Visible = True
    DecisionModelHelp.MultiPage1.Pages(4).Visible = True

    DecisionModelHelp.MultiPage1.value = 0

    DecisionModelHelp.Show
End Sub
