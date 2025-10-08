VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
DevelopmentConsole.Show

End Sub

Private Sub StartGameButton_Click()

Sheets("Game1").Activate
ActiveSheet.Range("CO595").Activate
Call resetALL
Call alignScreen
Set LinkSprite = ActiveSheet.Pictures("LinkDown1")
Call calculateScreenLocation(1, "D")

On Error GoTo endPoint

mySub = currentScreen 'global
Application.Run mySub

Call runGame

Exit Sub
endPoint:
MsgBox "Screen setup macro not found: " & mySub




End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
'ActiveCell.Value = ""
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

'ActiveSheet.Range("A1").Activate
'ActiveCell.Value = ""
End Sub
