Attribute VB_Name = "AG_LinkActions"
'##
'
'
'
'##
Sub Falling()

Sheets("Data").Range("C10").Value = "Y"
    
Dim location

location = Mid(CodeCell, 5, 4)

If location = "XXXX" Then
    location = Sheets("Data").Range("C8").Value
End If

Dim a, b, c

Select Case moveDir

    Case Is = "U"
        ActiveSheet.Shapes("LinkFall1").Top = LinkSprite.Top - 15
        ActiveSheet.Shapes("LinkFall1").Left = LinkSprite.Left
    Case Is = "D"
        ActiveSheet.Shapes("LinkFall1").Top = LinkSprite.Top + 50
        ActiveSheet.Shapes("LinkFall1").Left = LinkSprite.Left
    Case Is = "L"
        ActiveSheet.Shapes("LinkFall1").Top = LinkSprite.Top
        ActiveSheet.Shapes("LinkFall1").Left = LinkSprite.Left - 20
    Case Is = "R"
        ActiveSheet.Shapes("LinkFall1").Top = LinkSprite.Top
        ActiveSheet.Shapes("LinkFall1").Left = LinkSprite.Left + 20

End Select

ActiveSheet.Shapes("LinkFall2").Top = ActiveSheet.Shapes("LinkFall1").Top
ActiveSheet.Shapes("LinkFall2").Left = ActiveSheet.Shapes("LinkFall1").Left

ActiveSheet.Shapes("LinkFall3").Top = ActiveSheet.Shapes("LinkFall1").Top
ActiveSheet.Shapes("LinkFall3").Left = ActiveSheet.Shapes("LinkFall1").Left

    LinkSprite.Visible = False
    LinkSprite.Visible = False

    ActiveSheet.Shapes("LinkFall1").Visible = True
    Range("A1").Copy Range("A2")
    For a = 1 To 30
        Sleep 10
    Next a
    ActiveSheet.Shapes("LinkFall1").Visible = False
    
    
    ActiveSheet.Shapes("LinkFall2").Visible = True
    Range("A1").Copy Range("A2")
    For b = 1 To 30
        Sleep 10
    Next b
    ActiveSheet.Shapes("LinkFall2").Visible = False
    
    ActiveSheet.Shapes("LinkFall3").Visible = True
    Range("A1").Copy Range("A2")
    For c = 1 To 30
        Sleep 10
    Next c
    ActiveSheet.Shapes("LinkFall3").Visible = False
    
    'MsgBox (Sheets("Data").Range("C8").Value)
    
    Call Relocate(location)
    
    Sheets("Data").Range("C10").Value = "N"


End Sub


'##
'
'
'
'##
Sub JumpDown()

Sheets("Data").Range("C10").Value = "Y"

'Reset the 'prevent re-scrolling' timer
Sheets("Data").Range("C6").Value = "0"

Dim jumpCol, jumpRow, jumpTo, jumpCell, jumpVal, EnemyVal
'specify the row to move down to

jumpCol = Range(linkCellAddress).Column
jumpRow = Mid(CodeCell, 5, 3)

jumpTo = Cells(jumpRow, jumpCol).Address

'place the shadow
ActiveSheet.Shapes("LinkShadow").Top = Range(jumpTo).Top
ActiveSheet.Shapes("LinkShadow").Left = Range(jumpTo).Left
ActiveSheet.Shapes("LinkShadow").Left = ActiveSheet.Shapes("LinkShadow").Left - 5
ActiveSheet.Shapes("LinkShadow").Top = ActiveSheet.Shapes("LinkShadow").Top + 5
ActiveSheet.Shapes("LinkShadow").Visible = True

'align the jumping sprites
ActiveSheet.Shapes("LinkJump1").Top = LinkSprite.Top + 10
ActiveSheet.Shapes("LinkJump1").Left = LinkSprite.Left

ActiveSheet.Shapes("LinkJump2").Top = ActiveSheet.Shapes("LinkJump1").Top + 30
ActiveSheet.Shapes("LinkJump2").Left = LinkSprite.Left

ActiveSheet.Shapes("LinkJump3").Top = ActiveSheet.Shapes("LinkJump2").Top + 30
ActiveSheet.Shapes("LinkJump3").Left = LinkSprite.Left

LinkSprite.Visible = False
LinkSprite.Visible = False

ActiveSheet.Shapes("LinkJump1").Visible = True
ActiveSheet.Shapes("LinkJump2").Visible = False
ActiveSheet.Shapes("LinkJump3").Visible = False

'-----------------------------------------------------
'Begin somersault
For a = 1 To 10
ActiveSheet.Shapes("LinkJump1").Top = ActiveSheet.Shapes("LinkJump1").Top + 2
LinkSprite.Top = ActiveSheet.Shapes("LinkJump1").Top
linkCellAddress = LinkSprite.TopLeftCell.Address
jumpCell = ActiveSheet.Shapes("LinkJump1").TopLeftCell.Address
LinkSprite.Top = ActiveSheet.Shapes("LinkJump1").Top
jumpVal = Range(jumpCell).Value
EnemyVal = Mid(jumpVal, 7, 2)
jumpVal = Left(jumpVal, 2)


Select Case jumpVal

    Case Is = "S1"
        Call myScroll(1)
    Case Is = "S2"
        Call myScroll(2)
    Case Else
        ' Do nothing
End Select


Sleep 10
Range("A1").Copy Range("A2")
Next a


'-----------------------------------------------------
ActiveSheet.Shapes("LinkJump1").Visible = False
ActiveSheet.Shapes("LinkJump2").Visible = True

For a = 1 To 10
ActiveSheet.Shapes("LinkJump2").Top = ActiveSheet.Shapes("LinkJump2").Top + 2
LinkSprite.Top = ActiveSheet.Shapes("LinkJump2").Top
linkCellAddress = LinkSprite.TopLeftCell.Address
jumpCell = ActiveSheet.Shapes("LinkJump2").TopLeftCell.Address
jumpVal = Range(jumpCell).Value
EnemyVal = Mid(jumpVal, 7, 2)
jumpVal = Left(jumpVal, 2)


Select Case jumpVal

    Case Is = "S1"
        Call myScroll(1)
    Case Is = "S2"
        Call myScroll(2)
    Case Else
        ' Do nothing
End Select


Sleep 10
Range("A1").Copy Range("A2")
Next a

'-----------------------------------------------------
ActiveSheet.Shapes("LinkJump2").Visible = False
ActiveSheet.Shapes("LinkJump3").Visible = True


For a = 1 To 10
ActiveSheet.Shapes("LinkJump3").Top = ActiveSheet.Shapes("LinkJump3").Top + 2
LinkSprite.Top = ActiveSheet.Shapes("LinkJump3").Top
linkCellAddress = LinkSprite.TopLeftCell.Address
jumpCell = ActiveSheet.Shapes("LinkJump3").TopLeftCell.Address
jumpVal = Range(jumpCell).Value
EnemyVal = Mid(jumpVal, 7, 2)
jumpVal = Left(jumpVal, 2)


Select Case jumpVal

    Case Is = "S1"
        Call myScroll(1)
    Case Is = "S2"
        Call myScroll(2)
    Case Else
        ' Do nothing
End Select


Sleep 10
Range("A1").Copy Range("A2")
Next a

'-----------------------------------------------------
ActiveSheet.Shapes("LinkJump3").Visible = False

LinkSprite.Top = ActiveSheet.Shapes("LinkJump3").Top

LinkSprite.Visible = True
linkCellAddress = LinkSprite.TopLeftCell.Address
CodeCell = ""


' Fall to bottom after somersault
Do Until LinkSprite.Top >= Range(jumpTo).Top - 30

LinkSprite.Top = LinkSprite.Top + 4
LinkSpriteLeft = LinkSprite.Left
LinkSpriteTop = LinkSprite.Top
linkCellAddress = LinkSprite.TopLeftCell.Address

jumpCell = LinkSprite.TopLeftCell.Address
jumpVal = Range(jumpCell).Value
EnemyVal = Mid(jumpVal, 7, 2)
jumpVal = Left(jumpVal, 2)

Select Case jumpVal

    Case Is = "S1"
        Call myScroll(1)
    Case Is = "S2"
        Call myScroll(2)
    Case Else
        ' Do nothing
End Select

Sleep 10
Range("A1").Copy Range("A2")
Loop

ActiveSheet.Shapes("LinkShadow").Visible = False

Sheets("Data").Range("C10").Value = "N"
End Sub

'####################################################################################
'#
'#    Sword stuff
'#
'####################################################################################

Sub swordSwipe(Indicator)


Dim keypressed

If Indicator = 2 Then
    keypressed = DPress
    
ElseIf Indicator = 1 Then
    keypressed = CPress
Else
    keypressed = 0
End If

'MsgBox ("Indicator = " & Indicator & ".  Keypressed = " & keypressed)

Select Case lastDir

    Case Is = "L"
        ActiveSheet.Shapes("SwordUp").Top = LinkSprite.Top - 30
        ActiveSheet.Shapes("SwordUp").Left = LinkSprite.Left - 10
        
        ActiveSheet.Shapes("SwordSwipeUpLeft").Top = LinkSprite.Top - 30
        ActiveSheet.Shapes("SwordSwipeUpLeft").Left = LinkSprite.Left - 50
        
        ActiveSheet.Shapes("SwordLeft").Top = LinkSprite.Top
        ActiveSheet.Shapes("SwordLeft").Left = LinkSprite.Left - 50
        
        Set SwordFrame1 = ActiveSheet.Shapes("SwordUp")
        Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeUpLeft")
        Set SwordFrame3 = ActiveSheet.Shapes("SwordLeft")
        
    Case Is = "R"
        'MsgBox "aligning sword"
        ActiveSheet.Shapes("SwordUp").Top = LinkSprite.Top - 30
        ActiveSheet.Shapes("SwordUp").Left = LinkSprite.Left + 30
        
        ActiveSheet.Shapes("SwordSwipeUpRight").Top = LinkSprite.Top - 30
        ActiveSheet.Shapes("SwordSwipeUpRight").Left = LinkSprite.Left + 45
        
        ActiveSheet.Shapes("SwordRight").Top = LinkSprite.Top
        ActiveSheet.Shapes("SwordRight").Left = LinkSprite.Left + 45
        
        Set SwordFrame1 = ActiveSheet.Shapes("SwordUp")
        Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeUpRight")
        Set SwordFrame3 = ActiveSheet.Shapes("SwordRight")
        
    Case Is = "U"
        ActiveSheet.Shapes("SwordUp").Top = LinkSprite.Top - 45
        ActiveSheet.Shapes("SwordUp").Left = LinkSprite.Left + 5
        
        ActiveSheet.Shapes("SwordSwipeUpRight").Top = LinkSprite.Top - 45
        ActiveSheet.Shapes("SwordSwipeUpRight").Left = LinkSprite.Left + 25
        
        ActiveSheet.Shapes("SwordRight").Top = LinkSprite.Top - 15
        ActiveSheet.Shapes("SwordRight").Left = LinkSprite.Left + 35
        
        Set SwordFrame1 = ActiveSheet.Shapes("SwordRight")
        Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeUpRight")
        Set SwordFrame3 = ActiveSheet.Shapes("SwordUp")
        
    Case Is = "D"
        ActiveSheet.Shapes("SwordLeft").Top = LinkSprite.Top
        ActiveSheet.Shapes("SwordLeft").Left = LinkSprite.Left - 50
        
        ActiveSheet.Shapes("SwordSwipeDownLeft").Top = LinkSprite.Top + 30
        ActiveSheet.Shapes("SwordSwipeDownLeft").Left = LinkSprite.Left - 45
        
        ActiveSheet.Shapes("SwordDown").Top = LinkSprite.Top + 40
        ActiveSheet.Shapes("SwordDown").Left = LinkSprite.Left - 25
        
        Set SwordFrame1 = ActiveSheet.Shapes("SwordLeft")
        Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeDownLeft")
        Set SwordFrame3 = ActiveSheet.Shapes("SwordDown")
    
    Case Is = "LD"
    
        ActiveSheet.Shapes("SwordLeft").Top = LinkSprite.Top
        ActiveSheet.Shapes("SwordLeft").Left = LinkSprite.Left - 50
        
        ActiveSheet.Shapes("SwordSwipeDownLeft").Top = LinkSprite.Top + 30
        ActiveSheet.Shapes("SwordSwipeDownLeft").Left = LinkSprite.Left - 45
        
        ActiveSheet.Shapes("SwordDown").Top = LinkSprite.Top + 40
        ActiveSheet.Shapes("SwordDown").Left = LinkSprite.Left - 25
        
        Set SwordFrame1 = ActiveSheet.Shapes("SwordLeft")
        Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeDownLeft")
        Set SwordFrame3 = ActiveSheet.Shapes("SwordDown")
        
    Case Is = "RD"
    
        ActiveSheet.Shapes("SwordLeft").Top = LinkSprite.Top
        ActiveSheet.Shapes("SwordLeft").Left = LinkSprite.Left - 50
        
        ActiveSheet.Shapes("SwordSwipeDownLeft").Top = LinkSprite.Top + 30
        ActiveSheet.Shapes("SwordSwipeDownLeft").Left = LinkSprite.Left - 45
        
        ActiveSheet.Shapes("SwordDown").Top = LinkSprite.Top + 40
        ActiveSheet.Shapes("SwordDown").Left = LinkSprite.Left - 25
        
        Set SwordFrame1 = ActiveSheet.Shapes("SwordLeft")
        Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeDownLeft")
        Set SwordFrame3 = ActiveSheet.Shapes("SwordDown")
  
    Case Is = "RU"
    
        ActiveSheet.Shapes("SwordUp").Top = LinkSprite.Top - 45
        ActiveSheet.Shapes("SwordUp").Left = LinkSprite.Left + 5
        
        ActiveSheet.Shapes("SwordSwipeUpRight").Top = LinkSprite.Top - 45
        ActiveSheet.Shapes("SwordSwipeUpRight").Left = LinkSprite.Left + 25
        
        ActiveSheet.Shapes("SwordRight").Top = LinkSprite.Top - 15
        ActiveSheet.Shapes("SwordRight").Left = LinkSprite.Left + 35
        
        Set SwordFrame1 = ActiveSheet.Shapes("SwordRight")
        Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeUpRight")
        Set SwordFrame3 = ActiveSheet.Shapes("SwordUp")
        
    Case Is = "LU"
    
        ActiveSheet.Shapes("SwordUp").Top = LinkSprite.Top - 45
        ActiveSheet.Shapes("SwordUp").Left = LinkSprite.Left + 5
        
        ActiveSheet.Shapes("SwordSwipeUpRight").Top = LinkSprite.Top - 45
        ActiveSheet.Shapes("SwordSwipeUpRight").Left = LinkSprite.Left + 25
        
        ActiveSheet.Shapes("SwordRight").Top = LinkSprite.Top - 15
        ActiveSheet.Shapes("SwordRight").Left = LinkSprite.Left + 35
        
        Set SwordFrame1 = ActiveSheet.Shapes("SwordRight")
        Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeUpRight")
        Set SwordFrame3 = ActiveSheet.Shapes("SwordUp")
        
End Select

Select Case keypressed

    Case Is <= 1
        SwordFrame1.Visible = True
        Range("A1").Copy Range("A2")
        Sleep 25
        
        'Call didSwordHit(SwordFrame1, RNDenemyFrame1_1)
        'Call didSwordHit(SwordFrame1, RNDenemyFrame1_2)
        
        'Call didSwordHit(SwordFrame1, RNDenemyFrame2_1)
        'Call didSwordHit(SwordFrame1, RNDenemyFrame2_2)
        
        'Call didSwordHit(SwordFrame1, RNDenemyFrame3_1)
        'Call didSwordHit(SwordFrame1, RNDenemyFrame3_2)
        
        'Call didSwordHit(SwordFrame1, RNDenemyFrame4_1)
        'Call didSwordHit(SwordFrame1, RNDenemyFrame4_2)
                
        SwordFrame1.Visible = False
        SwordFrame2.Visible = True
        Range("A1").Copy Range("A2")
        Sleep 25

        'Call didSwordHit(SwordFrame2, RNDenemyFrame1_1)
        'Call didSwordHit(SwordFrame2, RNDenemyFrame1_2)
        
        'Call didSwordHit(SwordFrame2, RNDenemyFrame2_1)
        'Call didSwordHit(SwordFrame2, RNDenemyFrame2_2)
        
        'Call didSwordHit(SwordFrame2, RNDenemyFrame3_1)
        'Call didSwordHit(SwordFrame2, RNDenemyFrame3_2)
        
        'Call didSwordHit(SwordFrame2, RNDenemyFrame4_1)
        'Call didSwordHit(SwordFrame2, RNDenemyFrame4_2)


        SwordFrame2.Visible = False
        SwordFrame3.Visible = True
        Range("A1").Copy Range("A2")
        Sleep 25
        
        Call didSwordHit(SwordFrame3, RNDenemyFrame1_1)
        Call didSwordHit(SwordFrame3, RNDenemyFrame1_2)
        
        Call didSwordHit(SwordFrame3, RNDenemyFrame2_1)
        Call didSwordHit(SwordFrame3, RNDenemyFrame2_2)
        
        Call didSwordHit(SwordFrame3, RNDenemyFrame3_1)
        Call didSwordHit(SwordFrame3, RNDenemyFrame3_2)
        
        Call didSwordHit(SwordFrame3, RNDenemyFrame4_1)
        Call didSwordHit(SwordFrame3, RNDenemyFrame4_2)
        
        Call swordHitBush(SwordFrame3)

        
        
        SwordFrame3.Visible = False
        
        Case Is <= 20
        
        'do nothing
        
        
        Case Is > 20
        SwordFrame3.Visible = True
        
        Call didSwordHit(SwordFrame3, RNDenemyFrame1_1)
        Call didSwordHit(SwordFrame3, RNDenemyFrame1_2)
        
        Call didSwordHit(SwordFrame3, RNDenemyFrame2_1)
        Call didSwordHit(SwordFrame3, RNDenemyFrame2_2)
        
        Call didSwordHit(SwordFrame3, RNDenemyFrame3_1)
        Call didSwordHit(SwordFrame3, RNDenemyFrame3_2)
        
        Call didSwordHit(SwordFrame3, RNDenemyFrame4_1)
        Call didSwordHit(SwordFrame3, RNDenemyFrame4_2)
        
        Call swordHitBush(SwordFrame3)
        'MsgBox SwordHit
        
        Case Else
        


End Select



End Sub

Sub swordSpin()

Dim spinFrame1, spinFrame2, SpinFrame3, spinFrame4, spinFrame5, spinFrame6, spinFrame7, spinFrame8
Dim linkSpin1, linkSpin2, linkSpin3, linkSpin4

SwordFrame1.Visible = False
SwordFrame2.Visible = False
SwordFrame3.Visible = False

ActiveSheet.Shapes("LinkLeft1").Visible = False
ActiveSheet.Shapes("LinkLeft2").Visible = False

ActiveSheet.Shapes("LinkRight1").Visible = False
ActiveSheet.Shapes("LinkRight2").Visible = False

ActiveSheet.Shapes("LinkUp1").Visible = False
ActiveSheet.Shapes("LinkUp2").Visible = False

ActiveSheet.Shapes("LinkDown1").Visible = False
ActiveSheet.Shapes("LinkDown2").Visible = False

'MsgBox "spinning! " & SwordFrame1.Name

'align the frames
    ActiveSheet.Shapes("SwordUp").Top = LinkSprite.Top - 30
    ActiveSheet.Shapes("SwordUp").Left = LinkSprite.Left
    
    ActiveSheet.Shapes("SwordRight").Top = LinkSprite.Top
    ActiveSheet.Shapes("SwordRight").Left = LinkSprite.Left + 35

    ActiveSheet.Shapes("SwordLeft").Top = LinkSprite.Top
    ActiveSheet.Shapes("SwordLeft").Left = LinkSprite.Left - 50

    ActiveSheet.Shapes("SwordDown").Top = LinkSprite.Top + 40
    ActiveSheet.Shapes("SwordDown").Left = LinkSprite.Left - 25
    
    ActiveSheet.Shapes("SwordSwipeUpLeft").Top = LinkSprite.Top - 30
    ActiveSheet.Shapes("SwordSwipeUpLeft").Left = LinkSprite.Left - 50

    ActiveSheet.Shapes("SwordSwipeUpRight").Top = LinkSprite.Top - 45
    ActiveSheet.Shapes("SwordSwipeUpRight").Left = LinkSprite.Left + 25
    
    ActiveSheet.Shapes("SwordSwipeDownRight").Top = LinkSprite.Top + 45
    ActiveSheet.Shapes("SwordSwipeDownRight").Left = LinkSprite.Left + 35

    ActiveSheet.Shapes("SwordSwipeDownLeft").Top = LinkSprite.Top + 30
    ActiveSheet.Shapes("SwordSwipeDownLeft").Left = LinkSprite.Left - 45
        
Select Case lastDir

    Case Is = "L"

        Set spinFrame1 = ActiveSheet.Shapes("swordLeft")
        Set spinFrame2 = ActiveSheet.Shapes("swordSwipeDownLeft")
        Set SpinFrame3 = ActiveSheet.Shapes("swordDown")
        Set spinFrame4 = ActiveSheet.Shapes("swordSwipeDownRight")
        Set spinFrame5 = ActiveSheet.Shapes("swordRight")
        Set spinFrame6 = ActiveSheet.Shapes("swordSwipeUpRight")
        Set spinFrame7 = ActiveSheet.Shapes("swordUp")
        Set spinFrame8 = ActiveSheet.Shapes("swordSwipeUpLeft")
        
        Set linkSpin1 = ActiveSheet.Shapes("LinkLeft1")
        Set linkSpin2 = ActiveSheet.Shapes("LinkDown1")
        Set linkSpin3 = ActiveSheet.Shapes("LinkRight1")
        Set linkSpin4 = ActiveSheet.Shapes("LinkUp1")

    Case Is = "R"
        Set spinFrame1 = ActiveSheet.Shapes("swordRight")
        Set spinFrame2 = ActiveSheet.Shapes("swordSwipeDownRight")
        Set SpinFrame3 = ActiveSheet.Shapes("swordDown")
        Set spinFrame4 = ActiveSheet.Shapes("swordSwipeDownLeft")
        Set spinFrame5 = ActiveSheet.Shapes("swordLeft")
        Set spinFrame6 = ActiveSheet.Shapes("swordSwipeUpLeft")
        Set spinFrame7 = ActiveSheet.Shapes("swordUp")
        Set spinFrame8 = ActiveSheet.Shapes("swordSwipeUpRight")
        
        Set linkSpin1 = ActiveSheet.Shapes("LinkRight1")
        Set linkSpin2 = ActiveSheet.Shapes("LinkDown1")
        Set linkSpin3 = ActiveSheet.Shapes("LinkLeft1")
        Set linkSpin4 = ActiveSheet.Shapes("LinkUp1")
        
    Case Is = "RU"
        Set spinFrame1 = ActiveSheet.Shapes("swordRight")
        Set spinFrame2 = ActiveSheet.Shapes("swordSwipeDownRight")
        Set SpinFrame3 = ActiveSheet.Shapes("swordDown")
        Set spinFrame4 = ActiveSheet.Shapes("swordSwipeDownLeft")
        Set spinFrame5 = ActiveSheet.Shapes("swordLeft")
        Set spinFrame6 = ActiveSheet.Shapes("swordSwipeUpLeft")
        Set spinFrame7 = ActiveSheet.Shapes("swordUp")
        Set spinFrame8 = ActiveSheet.Shapes("swordSwipeUpRight")
        
        Set linkSpin1 = ActiveSheet.Shapes("LinkRight1")
        Set linkSpin2 = ActiveSheet.Shapes("LinkDown1")
        Set linkSpin3 = ActiveSheet.Shapes("LinkLeft1")
        Set linkSpin4 = ActiveSheet.Shapes("LinkUp1")
        
    Case Is = "LU"
        Set spinFrame1 = ActiveSheet.Shapes("swordRight")
        Set spinFrame2 = ActiveSheet.Shapes("swordSwipeDownRight")
        Set SpinFrame3 = ActiveSheet.Shapes("swordDown")
        Set spinFrame4 = ActiveSheet.Shapes("swordSwipeDownLeft")
        Set spinFrame5 = ActiveSheet.Shapes("swordLeft")
        Set spinFrame6 = ActiveSheet.Shapes("swordSwipeUpLeft")
        Set spinFrame7 = ActiveSheet.Shapes("swordUp")
        Set spinFrame8 = ActiveSheet.Shapes("swordSwipeUpRight")
        
        Set linkSpin1 = ActiveSheet.Shapes("LinkRight1")
        Set linkSpin2 = ActiveSheet.Shapes("LinkDown1")
        Set linkSpin3 = ActiveSheet.Shapes("LinkLeft1")
        Set linkSpin4 = ActiveSheet.Shapes("LinkUp1")
        
    Case Is = "U"
        Set spinFrame1 = ActiveSheet.Shapes("swordUp")
        Set spinFrame2 = ActiveSheet.Shapes("swordSwipeUpLeft")
        Set SpinFrame3 = ActiveSheet.Shapes("swordLeft")
        Set spinFrame4 = ActiveSheet.Shapes("swordSwipeDownLeft")
        Set spinFrame5 = ActiveSheet.Shapes("swordDown")
        Set spinFrame6 = ActiveSheet.Shapes("swordSwipeDownRight")
        Set spinFrame7 = ActiveSheet.Shapes("swordRight")
        Set spinFrame8 = ActiveSheet.Shapes("swordSwipeUpRight")
        
        Set linkSpin1 = ActiveSheet.Shapes("LinkUp1")
        Set linkSpin2 = ActiveSheet.Shapes("LinkLeft1")
        Set linkSpin3 = ActiveSheet.Shapes("LinkDown1")
        Set linkSpin4 = ActiveSheet.Shapes("LinkRight1")
        
    Case Is = "D"
        Set spinFrame1 = ActiveSheet.Shapes("swordDown")
        Set spinFrame2 = ActiveSheet.Shapes("swordSwipeDownRight")
        Set SpinFrame3 = ActiveSheet.Shapes("swordRight")
        Set spinFrame4 = ActiveSheet.Shapes("swordSwipeUpRight")
        Set spinFrame5 = ActiveSheet.Shapes("swordUp")
        Set spinFrame6 = ActiveSheet.Shapes("swordSwipeUpLeft")
        Set spinFrame7 = ActiveSheet.Shapes("swordLeft")
        Set spinFrame8 = ActiveSheet.Shapes("swordSwipeDownLeft")
        
        Set linkSpin1 = ActiveSheet.Shapes("LinkDown1")
        Set linkSpin2 = ActiveSheet.Shapes("LinkRight1")
        Set linkSpin3 = ActiveSheet.Shapes("LinkUp1")
        Set linkSpin4 = ActiveSheet.Shapes("LinkLeft1")
        
    Case Is = "LD"
        Set spinFrame1 = ActiveSheet.Shapes("swordDown")
        Set spinFrame2 = ActiveSheet.Shapes("swordSwipeDownRight")
        Set SpinFrame3 = ActiveSheet.Shapes("swordRight")
        Set spinFrame4 = ActiveSheet.Shapes("swordSwipeUpRight")
        Set spinFrame5 = ActiveSheet.Shapes("swordUp")
        Set spinFrame6 = ActiveSheet.Shapes("swordSwipeUpLeft")
        Set spinFrame7 = ActiveSheet.Shapes("swordLeft")
        Set spinFrame8 = ActiveSheet.Shapes("swordSwipeDownLeft")
        
        Set linkSpin1 = ActiveSheet.Shapes("LinkDown1")
        Set linkSpin2 = ActiveSheet.Shapes("LinkRight1")
        Set linkSpin3 = ActiveSheet.Shapes("LinkUp1")
        Set linkSpin4 = ActiveSheet.Shapes("LinkLeft1")

    Case Is = "RD"
        Set spinFrame1 = ActiveSheet.Shapes("swordDown")
        Set spinFrame2 = ActiveSheet.Shapes("swordSwipeDownRight")
        Set SpinFrame3 = ActiveSheet.Shapes("swordRight")
        Set spinFrame4 = ActiveSheet.Shapes("swordSwipeUpRight")
        Set spinFrame5 = ActiveSheet.Shapes("swordUp")
        Set spinFrame6 = ActiveSheet.Shapes("swordSwipeUpLeft")
        Set spinFrame7 = ActiveSheet.Shapes("swordLeft")
        Set spinFrame8 = ActiveSheet.Shapes("swordSwipeDownLeft")
        
        Set linkSpin1 = ActiveSheet.Shapes("LinkDown1")
        Set linkSpin2 = ActiveSheet.Shapes("LinkRight1")
        Set linkSpin3 = ActiveSheet.Shapes("LinkUp1")
        Set linkSpin4 = ActiveSheet.Shapes("LinkLeft1")
End Select

'animate the spin

    spinFrame1.Visible = True
    linkSpin1.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25
    
    spinFrame1.Visible = False
    spinFrame2.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25

    spinFrame2.Visible = False
    SpinFrame3.Visible = True
    linkSpin1.Visible = False
    linkSpin2.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25
    
    SpinFrame3.Visible = False
    spinFrame4.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25

    spinFrame4.Visible = False
    spinFrame5.Visible = True
    linkSpin2.Visible = False
    linkSpin3.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25
    
    spinFrame5.Visible = False
    spinFrame6.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25

    spinFrame6.Visible = False
    spinFrame7.Visible = True
    linkSpin3.Visible = False
    linkSpin4.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25
    
    spinFrame7.Visible = False
    spinFrame8.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25
    
    spinFrame8.Visible = False
    spinFrame1.Visible = True
    linkSpin4.Visible = False
    linkSpin1.Visible = True
    Range("A1").Copy Range("A2")
    Sleep 25

    spinFrame1.Visible = False


spinFrame1 = ""
spinFrame2 = ""
SpinFrame3 = ""
spinFrame4 = ""
spinFrame5 = ""
spinFrame6 = ""
spinFrame7 = ""
spinFrame8 = ""

linkSpin1 = ""
linkSpin2 = ""
linkSpin3 = ""
linkSpin4 = ""




End Sub


Sub showShield()
Sheets("Data").Range("C28").Value = "Y"

Select Case moveDir
    
    Case Is = "D"
        Set shieldSprite = ActiveSheet.Shapes("LinkShieldDown")
        shieldSprite.Top = LinkSprite.Top
        shieldSprite.Left = LinkSprite.Left
        shieldSprite.Visible = True
        ActiveSheet.Shapes("LinkShieldUp").Visible = False
        ActiveSheet.Shapes("LinkShieldLeft").Visible = False
        ActiveSheet.Shapes("LinkShieldRight").Visible = False
        
    Case Is = "LD"
        Set shieldSprite = ActiveSheet.Shapes("LinkShieldDown")
        shieldSprite.Top = LinkSprite.Top
        shieldSprite.Left = LinkSprite.Left
        shieldSprite.Visible = True
        ActiveSheet.Shapes("LinkShieldUp").Visible = False
        ActiveSheet.Shapes("LinkShieldLeft").Visible = False
        ActiveSheet.Shapes("LinkShieldRight").Visible = False
        
    Case Is = "RD"
        Set shieldSprite = ActiveSheet.Shapes("LinkShieldDown")
        shieldSprite.Top = LinkSprite.Top
        shieldSprite.Left = LinkSprite.Left
        shieldSprite.Visible = True
        ActiveSheet.Shapes("LinkShieldUp").Visible = False
        ActiveSheet.Shapes("LinkShieldLeft").Visible = False
        ActiveSheet.Shapes("LinkShieldRight").Visible = False
        
     Case Is = "U"
        Set shieldSprite = ActiveSheet.Shapes("LinkShieldUp")
        shieldSprite.Top = LinkSprite.Top
        shieldSprite.Left = LinkSprite.Left
        shieldSprite.Visible = True
        ActiveSheet.Shapes("LinkShieldDown").Visible = False
        ActiveSheet.Shapes("LinkShieldLeft").Visible = False
        ActiveSheet.Shapes("LinkShieldRight").Visible = False
        
     Case Is = "RU"
        Set shieldSprite = ActiveSheet.Shapes("LinkShieldUp")
        shieldSprite.Top = LinkSprite.Top
        shieldSprite.Left = LinkSprite.Left
        shieldSprite.Visible = True
        ActiveSheet.Shapes("LinkShieldDown").Visible = False
        ActiveSheet.Shapes("LinkShieldLeft").Visible = False
        ActiveSheet.Shapes("LinkShieldRight").Visible = False
        
     Case Is = "LU"
        Set shieldSprite = ActiveSheet.Shapes("LinkShieldUp")
        shieldSprite.Top = LinkSprite.Top
        shieldSprite.Left = LinkSprite.Left
        shieldSprite.Visible = True
        ActiveSheet.Shapes("LinkShieldDown").Visible = False
        ActiveSheet.Shapes("LinkShieldLeft").Visible = False
        ActiveSheet.Shapes("LinkShieldRight").Visible = False
        
     Case Is = "L"
        Set shieldSprite = ActiveSheet.Shapes("LinkShieldLeft")
        shieldSprite.Top = LinkSprite.Top
        shieldSprite.Left = LinkSprite.Left
        shieldSprite.Visible = True
        ActiveSheet.Shapes("LinkShieldUp").Visible = False
        ActiveSheet.Shapes("LinkShieldDown").Visible = False
        ActiveSheet.Shapes("LinkShieldRight").Visible = False
        
     Case Is = "R"
        Set shieldSprite = ActiveSheet.Shapes("LinkShieldRight")
        shieldSprite.Top = LinkSprite.Top
        shieldSprite.Left = LinkSprite.Left
        shieldSprite.Visible = True
        ActiveSheet.Shapes("LinkShieldUp").Visible = False
        ActiveSheet.Shapes("LinkShieldLeft").Visible = False
        ActiveSheet.Shapes("LinkShieldDown").Visible = False
    Case Else


End Select

End Sub

