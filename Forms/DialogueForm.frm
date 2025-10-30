Private Sub CloseButton_Click()

DialogueForm.Hide

End Sub

Private Sub UserForm_Activate()

Dim dataSheet As Worksheet
Set dataSheet = GameRegistryInstance().GetGameDataSheet()
If dataSheet Is Nothing Then Set dataSheet = Sheets("Data")  ' Fallback

DialogueBox.Value = dataSheet.Range("C42").Value

DialogueBox.SetFocus

DialogueBox.SelStart = 0
DialogueBox.SelLength = 0

Me.Repaint
DoEvents

Me.StartUpPosition = 0
Me.Top = Application.Top + Application.Height - Me.Height
Me.Left = Application.Left + Application.Width - Me.Width


End Sub
