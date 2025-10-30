Option Explicit

' Sheet6 (Game1)
' Link Game Scene

Private Sub btnStartGame_Click()
    ' Reset all managers and start new game
    Engine_GameLoop.ConfigureGameSheet ""
    Engine_GameLoop.StartNewGame
End Sub

Private Sub btnContinue_Click()
    ' Continue game from current state
    Engine_GameLoop.ConfigureGameSheet ""
    Engine_GameLoop.ContinueGame
End Sub

'###################################################################################
'                       DEPRECATED - DEBUG FUNCTIONALITY
'###################################################################################
' NOTE: The functions below are development/debugging tools and should be removed
'       or moved to a dedicated debug sheet. Not for production use.
'###################################################################################

Private Sub btnDebugAlignSprites_Click()
    ' DEPRECATED: Manual sprite alignment for debugging
    ' TODO: Move to DevelopmentConsole or remove entirely
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

Private Sub btnDevConsole_Click()
    ' DEPRECATED: Move to title screens only
    DevelopmentConsole.Show
End Sub

Private Sub btnFindPic_Click()
    ' DEPRECATED: Move to title screens only
    FindPic.Show
End Sub
