VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DialogueForm 
   ClientHeight    =   3828
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7608
   OleObjectBlob   =   "DialogueForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DialogueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CloseButton_Click()

DialogueForm.Hide

End Sub

Private Sub UserForm_Activate()

DialogueBox.Value = Sheets("Data").Range("C42").Value

DialogueBox.SetFocus

DialogueBox.SelStart = 0
DialogueBox.SelLength = 0

Me.Repaint
DoEvents

Me.StartUpPosition = 0
Me.Top = Application.Top + Application.Height - Me.Height
Me.Left = Application.Left + Application.Width - Me.Width


End Sub
