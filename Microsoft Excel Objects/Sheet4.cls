'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "Sheet4"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    DevelopmentConsole.Show
End Sub

Private Sub StartGameButton_Click()
    ' Activate starting screen and begin game
    Sheets("Game1").Activate
    AA_GameLoop.StartNewGame
End Sub

Private Sub ContinueButton_Click()
    ' Continue game from current state
    Sheets("Game1").Activate
    AA_GameLoop.ContinueGame
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Reserved - no behavior yet
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Reserved - no behavior yet
End Sub
