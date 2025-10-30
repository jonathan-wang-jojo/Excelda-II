Option Explicit

' Sheet4 (Title)
' Link Game Title Screen

Private Sub btnDevConsole_Click()
    DevelopmentConsole.Show
End Sub

Private Sub btnStartGame_Click()
    ' Activate starting screen and begin game
    Sheets("Game1").Activate
    Engine_GameLoop.ConfigureGameSheet ""
    Engine_GameLoop.StartNewGame
End Sub

Private Sub btnContinue_Click()
    ' Continue game from current state
    Sheets("Game1").Activate
    Engine_GameLoop.ConfigureGameSheet ""
    Engine_GameLoop.ContinueGame
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Reserved - no behavior yet
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Reserved - no behavior yet
End Sub