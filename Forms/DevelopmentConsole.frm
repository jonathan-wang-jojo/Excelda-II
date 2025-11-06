Option Explicit

Private Sub ActionBox_Change()

Dim actionBox As Variant
Set actionBox = Me.ActionBox

If actionBox.Value <> "Relocate" Then
    RelocateBox.Value = ""
    RelocateBox.Locked = True
Else
    RelocateBox.Locked = False
End If

End Sub

Private Sub coordsButton_Click()

Dim myTop As Double, myLeft As Double, myCell As String

myTop = ActiveSheet.Shapes(NameBox.Value).Top
myLeft = ActiveSheet.Shapes(NameBox.Value).Left
myCell = ActiveSheet.Shapes(NameBox.Value).TopLeftCell.Address


MsgBox ("Top = " & myTop & Chr(10) _
 & "left = " & myLeft & Chr(10) _
 & "Cell ref = " & myCell)
 
End Sub

Private Sub EnemyBox_Change()

If EnemyBox.Value = "None" Then
    EnemyLocationBox.Value = ""
    EnemyLocationBox.Locked = True
Else
    EnemyLocationBox.Locked = False
End If

End Sub



Private Sub EnemyLocationBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

Dim myVal As Long

myVal = Len(EnemyLocationBox.Value)

If myVal <> 4 Then

    MsgBox "The unique reference must be four characters long"
    
    EnemyLocationBox.Value = vbNullString

End If

End Sub

Private Sub FindButton_Click()

Dim myPic As String

myPic = NameBox.Value

On Error GoTo errortrap


ActiveSheet.Shapes(myPic).Top = 50
ActiveSheet.Shapes(myPic).Left = 50

ActiveSheet.Shapes(myPic).Visible = True

MsgBox "Done!"

Exit Sub

errortrap:

MsgBox ("Image not found")

End Sub



Private Sub GenerateCodeButton_Click()

Dim myAction As String, myEnemy As String, myScroll As String
Dim myDirection As String, myRelocate As String, myEnemyRelocate As String

myAction = vbNullString
myEnemy = vbNullString
myScroll = vbNullString
myDirection = vbNullString
myRelocate = vbNullString
myEnemyRelocate = vbNullString


If ActionBox.Value = vbNullString Then
    MsgBox "Please select an action (or 'none')"
    Exit Sub
End If

If ActionBox.Value = "Relocate" And RelocateBox.Value = vbNullString Then
    MsgBox "Please type a cell reference into the 'Relocate to' box"
    Exit Sub
End If

If EnemyBox.Value = vbNullString Then
    MsgBox "Please select an enemy (or 'none')"
    Exit Sub
End If

If EnemyBox.Value <> "None" And EnemyLocationBox.Value = vbNullString Then
    MsgBox "Please enter an enemy location"
    Exit Sub
End If

If ScreenScrollYes.Value = False And ScreenScrollNo.Value = False Then
    MsgBox ("Please select whether you want the screen to scroll when the code is triggered")
    Exit Sub
End If

If DirectionDown.Value = False And DirectionUp.Value = False And DirectionLeft.Value = False And DirectionRight.Value = False Then
    If ScreenScrollYes.Value = False And ScreenScrollNo.Value = False Then
        MsgBox ("Please select the direction you wish the screen to scroll")
        Exit Sub
    End If
End If





Select Case ActionBox.Value

    Case Is = "None"
        myAction = "XX"
        myRelocate = "XX"
    Case Is = "Relocate"
        myAction = "RL"
        myRelocate = RelocateBox.Value
        GoTo displayCode:
    Case Is = "Fall"
        myAction = "FL"
        myRelocate = RelocateBox.Value
    Case Is = "Push"
        myAction = "PU"
        myRelocate = "XXXX"
    
    'case is = ### ADD MORE CASES HERE ###
    
    Case Else
        MsgBox ("Unknown action selected - unable to continue")
        Exit Sub
        
End Select

Select Case EnemyBox.Value

    Case Is = "None"
        myEnemy = "XXXXXX"
        myEnemyRelocate = "XXXX"
    Case Is = "Sandcrab1F1"
        myEnemy = "ETSC01"
        myEnemyRelocate = EnemyLocationBox.Value
    Case Is = "Sandcrab2F1"
        myEnemy = "ETSC02"
        myEnemyRelocate = EnemyLocationBox.Value
    Case Is = "Octorok1F1"
        myEnemy = "ETOC01"
        myEnemyRelocate = EnemyLocationBox.Value
    Case Is = "Octorok2F1"
        myEnemy = "ETOC02"
        myEnemyRelocate = EnemyLocationBox.Value
    Case Is = "Sandspinner1F1"
        myEnemy = "ETSS01"
        myEnemyRelocate = EnemyLocationBox.Value
    Case Is = "Sandspinner2F1"
        myEnemy = "ETSS02"
        myEnemyRelocate = EnemyLocationBox.Value
     Case Is = "Moblin1F1"
        myEnemy = "ETMB01"
        myEnemyRelocate = EnemyLocationBox.Value
    Case Is = "Moblin2F1"
        myEnemy = "ETMB02"
        myEnemyRelocate = EnemyLocationBox.Value
    Case Is = "Moblin3F1"
        myEnemy = "ETMB03"
        myEnemyRelocate = EnemyLocationBox.Value
        
    'case is = ### ADD MORE CASES HERE ###
    
    Case Else
        MsgBox ("Unknown enemy selected - unable to continue")
        Exit Sub
        
End Select


If ScreenScrollYes.Value = True Then
    myScroll = "S"
End If

If ScreenScrollNo.Value = True Then
    myScroll = "X"
End If

If DirectionDown.Value = False And DirectionUp.Value = False And DirectionLeft.Value = False And DirectionRight.Value = False Then

    myScroll = myScroll & "X"
    myDirection = "X"

Else

    If DirectionDown.Value = True Then

        myScroll = myScroll & "1"
        myDirection = "D"

    End If

    If DirectionUp.Value = True Then

        myScroll = myScroll & "1"
        myDirection = "U"
    
    End If


    If DirectionLeft.Value = True Then

       myScroll = myScroll & "2"
        myDirection = "L"
    
    End If

    If DirectionRight.Value = True Then

        myScroll = myScroll & "2"
        myDirection = "R"

    End If

End If

displayCode:

CodeBox.Value = myScroll & myAction & myRelocate & myEnemy & myDirection & myEnemyRelocate


End Sub







Private Sub Givebutton_Click()

If SlotButton1.Value = True Then

    Sheets("Data").Range("C27").Value = giveWhatBox.Value
    MsgBox giveWhatBox.Value & " assigned to D"

ElseIf SlotButton2.Value = True Then

    Sheets("Data").Range("C26").Value = giveWhatBox.Value
    MsgBox giveWhatBox.Value & " assigned to C"
Else

    MsgBox "Please select which button to assign to (D or C)"
    Exit Sub
End If



End Sub

Private Sub RenameButton_Click()

Dim myFrom As String
Dim myTo As String

On Error GoTo errortrap

myFrom = NameBox.Value
myTo = NewNameBox.Value

ActiveSheet.Pictures(myFrom).Name = myTo

MsgBox ("Done")
Exit Sub

errortrap:

MsgBox ("Image not found")

End Sub

Private Sub ScreenScrollNo_Click()

DirectionDown.Value = False
DirectionUp.Value = False
DirectionLeft.Value = False
DirectionRight.Value = False


End Sub

Private Sub ShowAllButton_Click()

Dim sh As Shape
'this macro is hiding all the images on active sheet.

For Each sh In Sheets("Game1").Shapes

If sh.Type = msoPicture Then
    sh.Visible = True
End If

Next sh



'
'ActiveSheet.Shapes("LinkFall1").Visible = True
'ActiveSheet.Shapes("LinkFall2").Visible = True
'ActiveSheet.Shapes("LinkFall3").Visible = True

'ActiveSheet.Shapes("LinkUp1").Visible = True
'ActiveSheet.Shapes("LinkUp2").Visible = True

'ActiveSheet.Shapes("LinkDown1").Visible = True
'ActiveSheet.Shapes("LinkDown2").Visible = True

'ActiveSheet.Shapes("LinkRight1").Visible = True
'ActiveSheet.Shapes("LinkRight2").Visible = True

'ActiveSheet.Shapes("LinkLeft1").Visible = True
'ActiveSheet.Shapes("LinkLeft2").Visible = True

'ActiveSheet.Shapes("LinkUp1").Top = 100
'ActiveSheet.Shapes("LinkUp1").Left = 100

'ActiveSheet.Shapes("LinkUp2").Top = ActiveSheet.Shapes("LinkUp1").Top
'ActiveSheet.Shapes("LinkUp2").Left = ActiveSheet.Shapes("LinkUp1").Left + 31

'ActiveSheet.Shapes("LinkDown1").Top = ActiveSheet.Shapes("LinkUp1").Top + 31
'ActiveSheet.Shapes("LinkDown1").Left = ActiveSheet.Shapes("LinkUp1").Left
'ActiveSheet.Shapes("LinkDown2").Top = ActiveSheet.Shapes("LinkUp2").Top + 31
'ActiveSheet.Shapes("LinkDown2").Left = ActiveSheet.Shapes("LinkUp2").Left

'ActiveSheet.Shapes("LinkLeft1").Top = ActiveSheet.Shapes("LinkDown1").Top + 31
'ActiveSheet.Shapes("LinkLeft1").Left = ActiveSheet.Shapes("LinkDown1").Left
'ActiveSheet.Shapes("LinkLeft2").Top = ActiveSheet.Shapes("LinkDown2").Top + 31
'ActiveSheet.Shapes("LinkLeft2").Left = ActiveSheet.Shapes("LinkDown2").Left

'ActiveSheet.Shapes("LinkRight1").Top = ActiveSheet.Shapes("LinkLeft1").Top + 31
'ActiveSheet.Shapes("LinkRight1").Left = ActiveSheet.Shapes("LinkLeft1").Left
'ActiveSheet.Shapes("LinkRight2").Top = ActiveSheet.Shapes("LinkLeft2").Top + 31
'ActiveSheet.Shapes("LinkRight2").Left = ActiveSheet.Shapes("LinkLeft2").Left

'ActiveSheet.Shapes("skeletonDown1").Visible = True
'ActiveSheet.Shapes("skeletonDown2").Visible = True

'ActiveSheet.Shapes("Octorok1F1").Visible = True
'ActiveSheet.Shapes("Octorok1F2").Visible = True

'ActiveSheet.Shapes("Octorok2F1").Visible = True
'ActiveSheet.Shapes("Octorok2F2").Visible = True

'ActiveSheet.Shapes("Sandcrab1F1").Visible = True
'ActiveSheet.Shapes("Sandcrab1F2").Visible = True

'ActiveSheet.Shapes("Sandcrab2F1").Visible = True
'ActiveSheet.Shapes("Sandcrab2F2").Visible = True

'ActiveSheet.Shapes("Gordo1F1").Visible = True
'ActiveSheet.Shapes("Gordo1F2").Visible = True

'ActiveSheet.Shapes("Gordo2F1").Visible = True
'ActiveSheet.Shapes("Gordo2F2").Visible = True
'
'ActiveSheet.Shapes("Gordo3F1").Visible = True
'ActiveSheet.Shapes("Gordo3F2").Visible = True

MsgBox ("All images shown")

End Sub



Private Sub StartButton_Click()

End Sub

Private Sub TitleScreenButton_Click()

Sheets("Title").Activate

End Sub
