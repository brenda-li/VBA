Option Explicit

Private Sub CommandButton1_Click()
    ActiveWindow.SmallScroll Down:=4
End Sub

Private Sub CommandButton2_Click()
    ActiveWindow.SmallScroll Down:=27
End Sub

Private Sub CommandButton3_Click()
    ActiveWindow.SmallScroll Down:=52
End Sub

Private Sub CommandButton4_Click()
    ActiveWindow.SmallScroll Down:=78
End Sub

Private Sub CommandButton5_Click()
    ActiveWindow.SmallScroll Up:=90
End Sub

Private Sub CommandButton6_Click()
    ActiveWindow.SmallScroll Up:=90
End Sub

Private Sub CommandButton7_Click()
    ActiveWindow.SmallScroll Up:=90
End Sub

Private Sub CommandButton8_Click()
    ActiveWindow.SmallScroll Up:=90
End Sub

Private Sub CommandButton9_Click()
    DecisionModelHelp.MultiPage1.Pages(0).Visible = False
    DecisionModelHelp.MultiPage1.Pages(1).Visible = False
    DecisionModelHelp.MultiPage1.Pages(2).Visible = False
    DecisionModelHelp.MultiPage1.Pages(3).Visible = False

    DecisionModelHelp.MultiPage1.Pages(4).Visible = True
    DecisionModelHelp.MultiPage1.value = 4

    DecisionModelHelp.Show
End Sub

Private Sub CommandButton10_Click()
    'Return to Linear Value Model
    Worksheets("LinearValue").Activate
End Sub

Private Sub CommandButton11_Click()
    'Return to Exponential Value Model
    Worksheets("ExponValue").Activate
End Sub
