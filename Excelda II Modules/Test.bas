' Attribute VB_Name = "Test"
Sub resetALL()

'reset all enemies through EnemyManager
Call ResetAllEnemies

'hide all pictures
Dim sh As Shape

For Each sh In Sheets("Game1").Shapes

If sh.Type = msoPicture Then
    sh.Visible = False
End If
Next sh

'reset link
ActiveSheet.Shapes("LinkDown1").Visible = True
ActiveSheet.Shapes("Linkdown1").Top = ActiveCell.Top
ActiveSheet.Shapes("Linkdown1").Left = ActiveCell.Left

ActiveSheet.Shapes("LinkDown2").Top = ActiveSheet.Shapes("LinkDown1").Top
ActiveSheet.Shapes("LinkDown2").Left = ActiveSheet.Shapes("LinkDown1").Left

ActiveSheet.Shapes("LinkUp1").Top = ActiveSheet.Shapes("LinkDown1").Top
ActiveSheet.Shapes("LinkUp1").Left = ActiveSheet.Shapes("LinkDown1").Left
ActiveSheet.Shapes("LinkUp2").Top = ActiveSheet.Shapes("LinkDown1").Top
ActiveSheet.Shapes("LinkUp2").Left = ActiveSheet.Shapes("LinkDown1").Left

ActiveSheet.Shapes("LinkRight1").Top = ActiveSheet.Shapes("LinkDown1").Top
ActiveSheet.Shapes("LinkRight1").Left = ActiveSheet.Shapes("LinkDown1").Left
ActiveSheet.Shapes("LinkRight2").Top = ActiveSheet.Shapes("LinkDown1").Top
ActiveSheet.Shapes("LinkRight2").Left = ActiveSheet.Shapes("LinkDown1").Left

ActiveSheet.Shapes("LinkLeft1").Top = ActiveSheet.Shapes("LinkDown1").Top
ActiveSheet.Shapes("LinkLeft1").Left = ActiveSheet.Shapes("LinkDown1").Left
ActiveSheet.Shapes("LinkLeft2").Top = ActiveSheet.Shapes("LinkDown1").Top
ActiveSheet.Shapes("LinkLeft2").Left = ActiveSheet.Shapes("LinkDown1").Left

'reset the data sheet
Sheets("Data").Range("C6").Value = "0"
Sheets("Data").Range("C7").Value = ""
Sheets("Data").Range("C26").Value = ""
Sheets("Data").Range("C27").Value = ""
Sheets("Data").Range("Z2:Z500").Value = ""
Sheets("Data").Range("AB1:AB500").Value = ""


End Sub




