Private Sub CommandButton1_Click()
    'Unhide Worksheet
    Worksheets("RankObjectives").Visible = True
    'Continue Button
    Worksheets("RankObjectives").Activate
End Sub

Private Sub CommandButton2_Click()
    'Return to Data Entry Button
    Worksheets("DataEntry").Activate
End Sub
