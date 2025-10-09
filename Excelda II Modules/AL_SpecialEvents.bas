'Attribute VB_Name = "AL_SpecialEvents"

Sub SpecialEventTrigger(EventCode)


Dim eventID
eventID = Mid(EventCode, 9, 4)

mySub = "specialEvent" & eventID

Application.Run mySub

End Sub


Sub specialEvent0001()
'Links gets his sword
'XXXXXXSE0001XX

Dim Owl1, Owl2, frameCount, startCell
frameCount = 0
Set Owl1 = ActiveSheet.Shapes("Owl1")
Set Owl2 = ActiveSheet.Shapes("Owl2")

'Start at DY491
startCell = ActiveSheet.Range("DU487").Address
Owl1.Top = Range(startCell).Top
Owl1.Left = Range(startCell).Left
Owl2.Top = Range(startCell).Top
Owl2.Left = Range(startCell).Left

Owl1.Visible = True

For N = 1 To 30
    frameCount = frameCount + 1

    If frameCount = 3 Then
        Owl2.Visible = True
        Owl1.Visible = False
    End If

    If frameCount = 6 Then
        Owl1.Visible = True
        Owl2.Visible = False
        frameCount = 0
    End If

    Owl1.Top = Owl1.Top + 3
    Owl1.Left = Owl1.Left + 7
    Owl2.Top = Owl1.Top + 3
    Owl2.Left = Owl1.Left + 7
    Range("A1").Copy Range("A2")
    Sleep 25
    
Next N

Sheets("Data").Range("C42").Value = Sheets("Data").Range("T5").Value

DialogueForm.Show

frameCount = 0

For N = 1 To 30
    frameCount = frameCount + 1

    If frameCount = 3 Then
        Owl2.Visible = True
        Owl1.Visible = False
    End If

    If frameCount = 6 Then
        Owl1.Visible = True
        Owl2.Visible = False
        frameCount = 0
    End If

    Owl1.Top = Owl1.Top - 3
    Owl1.Left = Owl1.Left - 7
    Owl2.Top = Owl1.Top - 3
    Owl2.Left = Owl1.Left - 7
    Range("A1").Copy Range("A2")
    Sleep 25
    
Next N

Owl1.Visible = False
Owl2.Visible = False

'animate link getting the sword


ActiveSheet.Shapes("LinkWin").Top = LinkSprite.Top
ActiveSheet.Shapes("LinkWin").Left = LinkSprite.Left
LinkSprite.Visible = False
ActiveSheet.Shapes("LinkWin").Visible = True
ActiveSheet.Shapes("SwordUp").Top = ActiveSheet.Shapes("LinkWin").Top - 45
ActiveSheet.Shapes("SwordUp").Left = ActiveSheet.Shapes("LinkWin").Left + 20
Range("A1").Copy Range("A2")
Sleep 2000


ActiveSheet.Shapes("SwordUp").Visible = False
Call swordSpin

LinkSprite.Visible = True
ActiveSheet.Shapes("LinkWin").Visible = False
Range("A1").Copy Range("A2")

Sheets("Data").Range("C42").Value = Sheets("Data").Range("T6").Value
DialogueForm.Show

ActiveSheet.Range("EW507:EW514").Value = ""
ActiveSheet.Range("EW507:FG507").Value = ""
ActiveSheet.Range("FG507:FG514").Value = ""

Sheets("Data").Range("Z4").Value = "Y"
Sheets("Data").Range("C27").Value = "Sword"

CItem = Sheets("Data").Range("C26").Value
DItem = Sheets("Data").Range("C27").Value

End Sub

Sub specialEvent0002()

Call getHeartPiece("2")

End Sub

'Link can't leave without his shield
Sub specialEvent0003()

    Sheets("Data").Range("C42").Value = Sheets("Data").Range("T9").Value
    DialogueForm.Show
    
    
    ActiveSheet.Shapes("LinkUp1").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkUp1").Left = LinkSpriteLeft
    LinkSprite.Visible = False
    Set LinkSprite = ActiveSheet.Shapes("LinkUp1")
    LinkSprite.Visible = True
    LinkSprite.Top = LinkSprite.Top - 40
    Range("A1").Copy Range("A2")
    
    linkCellAddress = LinkSprite.TopLeftCell.Address
    LinkSpriteTop = LinkSprite.Top
    LinkSpriteLeft = LinkSprite.Left
    CodeCell = ""
    Sheets("Data").Range("C18").Value = linkCellAddress
    


End Sub

'Link gets his shield
Sub specialEvent0004()

'MsgBox "called"

Select Case GetAsyncKeyState(67)

    'If it's pressed
    Case Is <> 0
        
        GoTo Triggerpoint
        
    'If it isn't
    Case Is = 0
        

End Select

Select Case GetAsyncKeyState(68)

    'If it's pressed
    Case Is <> 0
        
        GoTo Triggerpoint
        
    'If it isn't
    Case Is = 0


End Select


Exit Sub

Triggerpoint:

        Sheets("Data").Range("C42").Value = Sheets("Data").Range("T10").Value
        DialogueForm.Show
    
        ActiveSheet.Shapes("LinkWin").Top = LinkSprite.Top
        ActiveSheet.Shapes("LinkWin").Left = LinkSprite.Left
        LinkSprite.Visible = False

        ActiveSheet.Shapes("LinkWin").Visible = True
        ActiveSheet.Shapes("LinkShieldDown").Top = ActiveSheet.Shapes("LinkWin").Top - 45
        ActiveSheet.Shapes("LinkShieldDown").Left = ActiveSheet.Shapes("LinkWin").Left
        ActiveSheet.Shapes("LinkShieldDown").Visible = True

        Range("A1").Copy Range("A2")
        Sleep 2000
        
        Sheets("Data").Range("C42").Value = Sheets("Data").Range("T11").Value
        DialogueForm.Show

        ActiveSheet.Shapes("LinkWin").Visible = False
        ActiveSheet.Shapes("LinkShieldDown").Visible = False
        LinkSprite.Visible = True

        Sheets("Data").Range("C42").Value = Sheets("Data").Range("T12").Value
        DialogueForm.Show
        
        Sheets("Data").Range("Z3").Value = "Y"
        Sheets("Data").Range("C26").Value = "Shield"
        
        CItem = Sheets("Data").Range("C26").Value
        DItem = Sheets("Data").Range("C27").Value
        
        ActiveSheet.Range(ActiveSheet.Range("DC595"), ActiveSheet.Range("DC595").Offset(11, 9)).Value = ""
        ActiveSheet.Range(ActiveSheet.Range("CQ613"), ActiveSheet.Range("CQ613").Offset(1, 7)).Value = ""
        

End Sub
