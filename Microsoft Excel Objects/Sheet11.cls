Option Explicit

' Code for Maze Title Sheet
' Sheet11 (Maze Title)

Private Sub btnDevConsole_Click()
    DevelopmentConsole.Show
End Sub

Private Sub btnStartGame_Click()
    ' Activate maze game sheet and start new game
    Sheets("Maze").Activate
    Engine_GameLoop.StartNewGameOnSheet "Maze", ""
End Sub

Private Sub btnContinue_Click()
    ' Activate maze game sheet and continue from current state
    Sheets("Maze").Activate
    Engine_GameLoop.ContinueGameOnSheet "Maze"
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    ' Reserved - no behavior yet
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    ' Reserved - no behavior yet
End Sub