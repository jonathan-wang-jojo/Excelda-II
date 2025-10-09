'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "Sheet6"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    ' Run game from current state
    Call Start
End Sub

Private Sub CommandButton2_Click()
    ' Reset all managers
    Call ResetAllManagers
End Sub

Private Sub CommandButton3_Click()
ActiveSheet.Shapes("LinkUp1").Visible = True
ActiveSheet.Shapes("LinkUp2").Visible = True

ActiveSheet.Shapes("LinkDown1").Visible = True
ActiveSheet.Shapes("LinkDown2").Visible = True

ActiveSheet.Shapes("LinkRight1").Visible = True
ActiveSheet.Shapes("LinkRight2").Visible = True

ActiveSheet.Shapes("LinkLeft1").Visible = True
ActiveSheet.Shapes("LinkLeft2").Visible = True

     
ActiveSheet.Shapes("LinkUp1").Top = 100
ActiveSheet.Shapes("LinkUp1").Left = 100
ActiveSheet.Shapes("LinkUp2").Top = ActiveSheet.Shapes("LinkUp1").Top
ActiveSheet.Shapes("LinkUp2").Left = ActiveSheet.Shapes("LinkUp1").Left + 31

ActiveSheet.Shapes("LinkDown1").Top = ActiveSheet.Shapes("LinkUp1").Top + 31
ActiveSheet.Shapes("LinkDown1").Left = ActiveSheet.Shapes("LinkUp1").Left
ActiveSheet.Shapes("LinkDown2").Top = ActiveSheet.Shapes("LinkUp2").Top + 31
ActiveSheet.Shapes("LinkDown2").Left = ActiveSheet.Shapes("LinkUp2").Left

ActiveSheet.Shapes("LinkLeft1").Top = ActiveSheet.Shapes("LinkDown1").Top + 31
ActiveSheet.Shapes("LinkLeft1").Left = ActiveSheet.Shapes("LinkDown1").Left
ActiveSheet.Shapes("LinkLeft2").Top = ActiveSheet.Shapes("LinkDown2").Top + 31
ActiveSheet.Shapes("LinkLeft2").Left = ActiveSheet.Shapes("LinkDown2").Left

ActiveSheet.Shapes("LinkRight1").Top = ActiveSheet.Shapes("LinkLeft1").Top + 31
ActiveSheet.Shapes("LinkRight1").Left = ActiveSheet.Shapes("LinkLeft1").Left
ActiveSheet.Shapes("LinkRight2").Top = ActiveSheet.Shapes("LinkLeft2").Top + 31
ActiveSheet.Shapes("LinkRight2").Left = ActiveSheet.Shapes("LinkLeft2").Left

ActiveSheet.Shapes("skeletonDown1").Visible = True
ActiveSheet.Shapes("skeletonDown2").Visible = True

End Sub

Private Sub CommandButton4_Click()

DevelopmentConsole.Show


End Sub

Private Sub CommandButton5_Click()

FindPic.Show

End Sub
