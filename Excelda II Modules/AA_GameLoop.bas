Attribute VB_Name = "AA_GameLoop"
#If VBA7 Then
    Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Long
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


'---- Player state globals ----
Global LinkSprite As Shape
Global LinkSpriteTop As Double, LinkSpriteLeft As Double
Global moveDir As String, lastDir As String 'either 'U, UL, UR, D, DL, DR, L, R
Global LinkSpriteFrame As Integer 'either 1, or 2
Global LinkMove As Integer 'sheets("Data").range("C19").value
Global CItem, DItem As String
Global CPress As Long, DPress As Long

'---- Game timers ----
Global screenSetUpTimer As Long

'---- Misc sprites ----
Global SwordFrame1 As Shape, SwordFrame2 As Shape, SwordFrame3 As Shape
Global shieldSprite As Shape

'---- Misc ----
Global currentScreen As String
Global linkCellAddress As String
Global CodeCell As String
Global gameSpeed As Long


'###################################################################################
'
'
'
'
'
'###################################################################################

Sub runGame()

CItem = Sheets("Data").Range("C26").Value
DItem = Sheets("Data").Range("C27").Value
CPress = 0
DPress = 0

Set SwordFrame1 = ActiveSheet.Shapes("SwordLeft")
Set SwordFrame2 = ActiveSheet.Shapes("SwordSwipeDownLeft")
Set SwordFrame3 = ActiveSheet.Shapes("SwordDown")
Set shieldSprite = ActiveSheet.Shapes("LinkShieldDown")

screenSetUpTimer = 0

startSub:



'initial values
Set LinkSprite = ActiveSheet.Shapes("LinkDown2")

        
LinkMove = Sheets("Data").Range("C19").Value

gameSpeed = Sheets("Data").Range("C4").Value


'############################################################################
startLoop: '#################################################################

'MsgBox ("linkSprite = " + LinkSprite.Name)

If Sheets("Data").Range("C20").Value >= 5 Then
    LinkSpriteFrame = 1
Else
    LinkSpriteFrame = 2
End If

LinkSpriteLeft = LinkSprite.Left
LinkSpriteTop = LinkSprite.Top

'-------- Quit to title if 'Q' was pressed -----
If GetAsyncKeyState(81) <> 0 Then
    Application.CutCopyMode = False
    Sheets("Title").Activate
    ActiveSheet.Range("A1").Select
    GoTo endLoop
End If
'--------------------------------------

'MsgBox screenSetUpTimer

If screenSetUpTimer > 0 Then
    screenSetUpTimer = screenSetUpTimer - 1
End If

'check to see if an enemy collision had occurred
If RNDBounceback <> "" Then
    Call BounceBack(LinkSprite, ActiveSheet.Shapes(CollidedWith))
    GoTo afterMove
End If



'account for falling/jumping
Select Case Sheets("Data").Range("C9").Value

    Case Is = "Y"
    GoTo afterMove
    Case Else

End Select





'assign up to 2 four-way movement values
'Right #################################
Select Case GetAsyncKeyState(37)

Case Is <> 0


Sheets("Data").Range("C21").Value = Sheets("Data").Range("C21").Value + "L"

Case Is = 0

End Select

'Left ##############################

Select Case GetAsyncKeyState(39)

Case Is <> 0

Sheets("Data").Range("C21").Value = Sheets("Data").Range("C21").Value + "R"

Case Is = 0

'down ###########################
End Select

Select Case GetAsyncKeyState(40)

Case Is <> 0

Sheets("Data").Range("C21").Value = Sheets("Data").Range("C21").Value + "D"

Case Is = 0

End Select

'Up ###########################################
Select Case GetAsyncKeyState(38)

Case Is <> 0

Sheets("Data").Range("C21").Value = Sheets("Data").Range("C21").Value + "U"

Case Is = 0

End Select

'------------------------------------------------------------------------
moveDir = Sheets("Data").Range("C21").Value

If moveDir <> "" Then
    lastDir = moveDir
End If

Select Case moveDir

    Case Is = "U"
        If LinkSpriteFrame = 1 Then
            Set LinkSprite = ActiveSheet.Shapes("LinkUp1")
        Else
            Set LinkSprite = ActiveSheet.Shapes("LinkUp2")
        End If
        
        LinkSpriteTop = LinkSpriteTop - LinkMove
        
    Case Is = "D"
        If LinkSpriteFrame = 1 Then
            Set LinkSprite = ActiveSheet.Shapes("LinkDown1")
        Else
            Set LinkSprite = ActiveSheet.Shapes("LinkDown2")
        End If
        LinkSpriteTop = LinkSpriteTop + LinkMove

    Case Is = "R"
        If LinkSpriteFrame = 1 Then
            Set LinkSprite = ActiveSheet.Shapes("LinkRight1")
        Else
            Set LinkSprite = ActiveSheet.Shapes("LinkRight2")
        End If
        LinkSpriteLeft = LinkSpriteLeft + LinkMove
        
    Case Is = "L"
        If LinkSpriteFrame = 1 Then
            Set LinkSprite = ActiveSheet.Shapes("LinkLeft1")
        Else
            Set LinkSprite = ActiveSheet.Shapes("LinkLeft2")
        End If
        LinkSpriteLeft = LinkSpriteLeft - LinkMove
        
    Case Is = "LU"
    If LinkSpriteFrame = 1 Then
        'Set LinkSprite = ActiveSheet.Shapes("LinkUpLeft1")
        Set LinkSprite = ActiveSheet.Shapes("LinkUp1")
    Else
         'Set LinkSprite = ActiveSheet.Shapes("LinkUpLeft2")
         Set LinkSprite = ActiveSheet.Shapes("LinkUp2")
    End If
        LinkSpriteLeft = LinkSpriteLeft - LinkMove
        LinkSpriteTop = LinkSpriteTop - LinkMove
        
    Case Is = "UL"
    If LinkSpriteFrame = 1 Then
        'Set LinkSprite = ActiveSheet.Shapes("LinkUpLeft1")
        Set LinkSprite = ActiveSheet.Shapes("LinkUp1")
    Else
         'Set LinkSprite = ActiveSheet.Shapes("LinkUpLeft2")
         Set LinkSprite = ActiveSheet.Shapes("LinkUp2")
    End If
        LinkSpriteLeft = LinkSpriteLeft - LinkMove
        LinkSpriteTop = LinkSpriteTop - LinkMove
        
    Case Is = "RU"
        If LinkSpriteFrame = 1 Then
            'Set LinkSprite = ActiveSheet.Shapes("LinkUpRight1")
            Set LinkSprite = ActiveSheet.Shapes("LinkUp1")
        Else
            'Set LinkSprite = ActiveSheet.Shapes("LinkUpRight2")
            Set LinkSprite = ActiveSheet.Shapes("LinkUp2")
        End If
        LinkSpriteLeft = LinkSpriteLeft + LinkMove
        LinkSpriteTop = LinkSpriteTop - LinkMove
        
    Case Is = "UR"
        If LinkSpriteFrame = 1 Then
            'Set LinkSprite = ActiveSheet.Shapes("LinkUpRight1")
            Set LinkSprite = ActiveSheet.Shapes("LinkUp1")
        Else
            'Set LinkSprite = ActiveSheet.Shapes("LinkUpRight2")
            Set LinkSprite = ActiveSheet.Shapes("LinkUp2")
        End If
        LinkSpriteLeft = LinkSpriteLeft + LinkMove
        LinkSpriteTop = LinkSpriteTop - LinkMove
        
    Case Is = "LD"
        If LinkSpriteFrame = 1 Then
            'Set LinkSprite = ActiveSheet.Shapes("LinkDownLeft2")
            Set LinkSprite = ActiveSheet.Shapes("LinkDown2")
        Else
            'Set LinkSprite = ActiveSheet.Shapes("LinkDownLeft1")
            Set LinkSprite = ActiveSheet.Shapes("LinkDown1")
        End If
        LinkSpriteLeft = LinkSpriteLeft - LinkMove
        LinkSpriteTop = LinkSpriteTop + LinkMove
        
    Case Is = "DL"
        If LinkSpriteFrame = 1 Then
            'Set LinkSprite = ActiveSheet.Shapes("LinkDownLeft2")
            Set LinkSprite = ActiveSheet.Shapes("LinkDown2")
        Else
            'Set LinkSprite = ActiveSheet.Shapes("LinkDownLeft1")
            Set LinkSprite = ActiveSheet.Shapes("LinkDown1")
        End If
        LinkSpriteLeft = LinkSpriteLeft - LinkMove
        LinkSpriteTop = LinkSpriteTop + LinkMove
        
    Case Is = "RD"
        If LinkSpriteFrame = 1 Then
            'Set LinkSprite = ActiveSheet.Shapes("LinkDownRight1")
            Set LinkSprite = ActiveSheet.Shapes("LinkDown1")
        Else
            'Set LinkSprite = ActiveSheet.Shapes("LinkDownRight2")
            Set LinkSprite = ActiveSheet.Shapes("LinkDown2")
        End If
        LinkSpriteLeft = LinkSpriteLeft + LinkMove
        LinkSpriteTop = LinkSpriteTop + LinkMove
        
    Case Is = "DR"
        If LinkSpriteFrame = 1 Then
            'Set LinkSprite = ActiveSheet.Shapes("LinkDownRight1")
            Set LinkSprite = ActiveSheet.Shapes("LinkDown1")
        Else
            'Set LinkSprite = ActiveSheet.Shapes("LinkDownRight2")
            Set LinkSprite = ActiveSheet.Shapes("LinkDown2")
        End If
        LinkSpriteLeft = LinkSpriteLeft + LinkMove
        LinkSpriteTop = LinkSpriteTop + LinkMove
        
    Case Else
        'MsgBox ("Gameloop:  No MoveDir - Linksprite = " & LinkSprite.Name)
End Select


'###########################################################################
' C and D keys (actions) ###################################################
'###########################################################################


'C pressed ###########################################
Select Case GetAsyncKeyState(67)

'If it's pressed
Case Is <> 0
    Sheets("Data").Range("C24").Value = "Y"
    CPress = CPress + 1
    
    Select Case CItem
    
        Case Is = "Sword"
            Call swordSwipe(1)
            
        Case Is = "Shield"
            Call showShield
            
        Case Else
            'Insert more items here
    End Select

Case Is = 0

    Select Case CItem
        'If the sword has stopped being pressed
        Case Is = "Sword"
            
            Select Case CPress
                Case Is >= 20
                    SwordFrame1.Visible = False
                    SwordFrame2.Visible = False
                    SwordFrame2.Visible = False
                    Call swordSpin
                    
                Case Else
                    SwordFrame1.Visible = False
                    SwordFrame2.Visible = False
                    SwordFrame2.Visible = False

            End Select
            
        Case Is = "Shield"
            shieldSprite.Visible = False
            Sheets("Data").Range("C28").Value = ""
        
        Case Else

    End Select
    
    'Reset the flags
    Sheets("Data").Range("C24").Value = ""
    CPress = 0
    
End Select




'D pressed ###########################################
Select Case GetAsyncKeyState(68)

'If it's pressed
Case Is <> 0
    Sheets("Data").Range("C25").Value = "Y"
    DPress = DPress + 1

    Select Case DItem
    
        Case Is = "Sword"
            Call swordSwipe(2)
            
        Case Is = "Shield"
            Call showShield
            
        Case Else
            'Insert more items here
    End Select

Case Is = 0

    Select Case DItem
        'If the sword has stopped being pressed
        Case Is = "Sword"
            
            Select Case DPress
                Case Is >= 20
                    SwordFrame1.Visible = False
                    SwordFrame2.Visible = False
                    SwordFrame2.Visible = False
                    Call swordSpin
                    
                Case Else
                    SwordFrame1.Visible = False
                    SwordFrame2.Visible = False
                    SwordFrame2.Visible = False

            End Select
            
        Case Is = "Shield"
            shieldSprite.Visible = False
            Sheets("Data").Range("C28").Value = ""
        
        Case Else

    End Select
    
    'Reset the flags
    Sheets("Data").Range("C25").Value = ""
    DPress = 0
    
End Select


'###########################################################################
'--------------- trigger stuff -------------------------------
'###########################################################################

'MsgBox LinkSprite.Name
linkCellAddress = LinkSprite.TopLeftCell.Address
'MsgBox ("Gameloop: linkcelladdress = " & linkCellAddress)

CodeCell = Range(linkCellAddress).Offset(3, 2).Value
'MsgBox ("Gameloop: Codecell = " & CodeCell)

Sheets("Data").Range("C18").Value = linkCellAddress

If CodeCell <> "" Then
    
    'Sheets("Data").Range("C8").Value = Range(linkCellAddress).Offset(3, 2).Value
    
    Dim ScrollIndicator
        ScrollIndicator = Left(CodeCell, 1)
        
    Dim scrollDir
        scrollDir = Mid(CodeCell, 2, 1)

    Dim FallIndicator
        FallIndicator = Mid(CodeCell, 3, 2)

    Dim ActionIndicator
        'ActionIndicator = Range(LinkCellAddress).Value
        ActionIndicator = Mid(CodeCell, 7, 2)
   
 'MsgBox ("ScrollInd = " & ScrollIndicator & Chr(10) & "FallInd = " & FallIndicator & Chr(10) & "ActionInd = " & ActionIndicator)
 
    
    Select Case ScrollIndicator

        Case Is = "S"
            Call myScroll(scrollDir)
            'MsgBox "scrolling"
        Case Else

    End Select
  
    Select Case FallIndicator

        Case Is = "FL"
            Call Falling
            
        Case Is = "JD"
            Call JumpDown
    Case Else

    End Select
   

    Select Case ActionIndicator

        Case Is = "RL"
            'MsgBox "Relocating"
            Call Relocate(CodeCell)
            GoTo startSub

        Case Is = "ET"
            'MsgBox "Enemy triggering"
            Call EnemyTrigger(CodeCell)
        
        Case Is = "SE"
            'MsgBox "Enemy triggering"
            Call SpecialEventTrigger(CodeCell)
        Case Else
            'Do nothing
    End Select



End If

'##################### Enemies #####################################
'-------------------------------------------------------------------


If RNDenemyName1 <> "" Then

    Call enemyCollision(LinkSprite, RNDenemyName1)
    If RNDenemyHit1 > 0 Then
        Call enemyBounceBack(1)
        GoTo endEnemy1
    End If
    Call RNDEnemyMove(1)
endEnemy1:

End If



If RNDenemyName2 <> "" Then
    Call enemyCollision(LinkSprite, RNDenemyName2)
    If RNDenemyHit2 > 0 Then
        Call enemyBounceBack(2)
        GoTo endEnemy2
    End If
    Call RNDEnemyMove(2)
endEnemy2:
End If



If RNDenemyName3 <> "" Then
    Call enemyCollision(LinkSprite, RNDenemyName3)
    
    If RNDenemyHit3 > 0 Then
        Call enemyBounceBack(3)
        GoTo endEnemy3
    End If
    Call RNDEnemyMove(3)
endEnemy3:
End If

If RNDenemyName4 <> "" Then
    Call enemyCollision(LinkSprite, RNDenemyName4)
    
    If RNDenemyHit4 > 0 Then
        Call enemyBounceBack(4)
        GoTo endEnemy4
    End If
    Call RNDEnemyMove(4)
endEnemy4:
End If
' #### add more enemies (if required) here ####


'-------------------------------------------------------------------

' ################### Link collision detection #####################

If moveDir = "D" And Range(linkCellAddress).Offset(4, 3).Value = "B" Then
    GoTo afterMove
End If

If moveDir = "U" And Range(linkCellAddress).Offset(0, 3).Value = "B" Then
    GoTo afterMove
End If

If moveDir = "L" And Range(linkCellAddress).Offset(4, 0).Value = "B" Then
    GoTo afterMove
End If

If moveDir = "R" And Range(linkCellAddress).Offset(1, 2).Value = "B" Then
    GoTo afterMove
End If

If moveDir = "R" And Range(linkCellAddress).Offset(4, 4).Value = "B" Then
    GoTo afterMove
End If

If moveDir = "RU" And Range(linkCellAddress).Offset(0, 3).Value = "B" Then
    GoTo afterMove
End If

If moveDir = "LU" And Range(linkCellAddress).Value = "B" Then
    GoTo afterMove
End If

If moveDir = "RD" And Range(linkCellAddress).Offset(4, 3).Value = "B" Then
    GoTo afterMove
End If

If moveDir = "LD" And Range(linkCellAddress).Offset(4, 0).Value = "B" Then
    GoTo afterMove
End If



'MsgBox (LinkSprite.Name)

Select Case LinkSprite.Name

'N,S,E,W #################################################

'Right

    Case Is = "LinkRight1"
    
        ActiveSheet.Shapes("LinkUp1").Visible = False
        ActiveSheet.Shapes("LinkUp2").Visible = False
        ActiveSheet.Shapes("LinkDown1").Visible = False
        ActiveSheet.Shapes("LinkDown2").Visible = False
        ActiveSheet.Shapes("LinkLeft1").Visible = False
        ActiveSheet.Shapes("LinkLeft2").Visible = False
        ActiveSheet.Shapes("LinkRight1").Visible = True
        ActiveSheet.Shapes("LinkRight2").Visible = False
        


    Case Is = "LinkRight2"
        ActiveSheet.Shapes("LinkUp1").Visible = False
        ActiveSheet.Shapes("LinkUp2").Visible = False
        ActiveSheet.Shapes("LinkDown1").Visible = False
        ActiveSheet.Shapes("LinkDown2").Visible = False
        ActiveSheet.Shapes("LinkLeft1").Visible = False
        ActiveSheet.Shapes("LinkLeft2").Visible = False
        ActiveSheet.Shapes("LinkRight1").Visible = False
        ActiveSheet.Shapes("LinkRight2").Visible = True
        
                      
'left
    Case Is = "LinkLeft1"
        ActiveSheet.Shapes("LinkUp1").Visible = False
        ActiveSheet.Shapes("LinkUp2").Visible = False
        ActiveSheet.Shapes("LinkDown1").Visible = False
        ActiveSheet.Shapes("LinkDown2").Visible = False
        ActiveSheet.Shapes("LinkLeft1").Visible = True
        ActiveSheet.Shapes("LinkLeft2").Visible = False
        ActiveSheet.Shapes("LinkRight1").Visible = False
        ActiveSheet.Shapes("LinkRight2").Visible = False
        
        
    Case Is = "LinkLeft2"
        ActiveSheet.Shapes("LinkUp1").Visible = False
        ActiveSheet.Shapes("LinkUp2").Visible = False
        ActiveSheet.Shapes("LinkDown1").Visible = False
        ActiveSheet.Shapes("LinkDown2").Visible = False
        ActiveSheet.Shapes("LinkLeft1").Visible = False
        ActiveSheet.Shapes("LinkLeft2").Visible = True
        ActiveSheet.Shapes("LinkRight1").Visible = False
        ActiveSheet.Shapes("LinkRight2").Visible = False
        
        
'down
    Case Is = "LinkDown1"
        ActiveSheet.Shapes("LinkUp1").Visible = False
        ActiveSheet.Shapes("LinkUp2").Visible = False
        ActiveSheet.Shapes("LinkDown1").Visible = True
        ActiveSheet.Shapes("LinkDown2").Visible = False
        ActiveSheet.Shapes("LinkLeft1").Visible = False
        ActiveSheet.Shapes("LinkLeft2").Visible = False
        ActiveSheet.Shapes("LinkRight1").Visible = False
        ActiveSheet.Shapes("LinkRight2").Visible = False
        

    Case Is = "LinkDown2"
        ActiveSheet.Shapes("LinkUp1").Visible = False
        ActiveSheet.Shapes("LinkUp2").Visible = False
        ActiveSheet.Shapes("LinkDown1").Visible = False
        ActiveSheet.Shapes("LinkDown2").Visible = True
        ActiveSheet.Shapes("LinkLeft1").Visible = False
        ActiveSheet.Shapes("LinkLeft2").Visible = False
        ActiveSheet.Shapes("LinkRight1").Visible = False
        ActiveSheet.Shapes("LinkRight2").Visible = False
        
        
'Up -------------------------------------------------------
    Case Is = "LinkUp1"
        ActiveSheet.Shapes("LinkUp1").Visible = True
        ActiveSheet.Shapes("LinkUp2").Visible = False
        ActiveSheet.Shapes("LinkDown1").Visible = False
        ActiveSheet.Shapes("LinkDown2").Visible = False
        ActiveSheet.Shapes("LinkLeft1").Visible = False
        ActiveSheet.Shapes("LinkLeft2").Visible = False
        ActiveSheet.Shapes("LinkRight1").Visible = False
        ActiveSheet.Shapes("LinkRight2").Visible = False
        

    Case Is = "LinkUp2"
        ActiveSheet.Shapes("LinkUp1").Visible = False
        ActiveSheet.Shapes("LinkUp2").Visible = True
        ActiveSheet.Shapes("LinkDown1").Visible = False
        ActiveSheet.Shapes("LinkDown2").Visible = False
        ActiveSheet.Shapes("LinkLeft1").Visible = False
        ActiveSheet.Shapes("LinkLeft2").Visible = False
        ActiveSheet.Shapes("LinkRight1").Visible = False
        ActiveSheet.Shapes("LinkRight2").Visible = False
        
    Case Else
        'do nothing
End Select


'LinkAlign:

'make sure all images are in the same place at all times
'Call alignLink
    ActiveSheet.Shapes("LinkUp1").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkUp1").Left = LinkSpriteLeft
    ActiveSheet.Shapes("LinkUp2").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkUp2").Left = LinkSpriteLeft
    
    ActiveSheet.Shapes("LinkDown1").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkDown1").Left = LinkSpriteLeft
    ActiveSheet.Shapes("LinkDown2").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkDown2").Left = LinkSpriteLeft

    ActiveSheet.Shapes("LinkRight1").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkRight1").Left = LinkSpriteLeft
    ActiveSheet.Shapes("LinkRight2").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkRight2").Left = LinkSpriteLeft
        
    ActiveSheet.Shapes("LinkLeft1").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkLeft1").Left = LinkSpriteLeft
    ActiveSheet.Shapes("LinkLeft2").Top = LinkSpriteTop
    ActiveSheet.Shapes("LinkLeft2").Left = LinkSpriteLeft
    


'--------------------------------------------------------------------
'reset the animation counter

If Sheets("Data").Range("C20").Value >= 10 Then
    Sheets("Data").Range("C20").Value = 0
Else
    Sheets("Data").Range("C20").Value = Sheets("Data").Range("C20").Value + 1
End If

afterMove:

'-------------------------------
Range("A1").Copy Range("A2")
Sheets("Data").Range("C21").Value = ""
Sleep (gameSpeed)


'scroll timer
'If Sheets("Data").Range("C6").Value <> 0 Then
'    Sheets("Data").Range("C6").Value = Sheets("Data").Range("C6").Value - 1
'Else
'    Sheets("Data").Range("C7").Value = "X"
'End If

Application.CutCopyMode = False
GoTo startLoop
'-------------------------------

endLoop:


End Sub

Sub Relocate(location)

If location = Sheets("Data").Range("C8").Value Then
    
    Range(location).Activate

    Select Case Sheets("Data").Range("C9").Value
    
    Case Is = "U"
        ActiveCell.Offset(-1, 0).Activate
    Case Is = "D"
        ActiveCell.Offset(1, 0).Activate
    Case Is = "L"
        ActiveCell.Offset(0, -1).Activate
    Case Is = "R"
        ActiveCell.Offset(0, 2).Activate
    Case Else
    
    End Select

Else

    Dim cellAdd
    cellAdd = location
    cellAdd = Right(cellAdd, 4)

    Cells.Find(What:=cellAdd, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        True, SearchFormat:=False).Activate
        
End If

'MsgBox "relocating all link images"

    ActiveSheet.Shapes("LinkRight1").Top = ActiveCell.Top
    ActiveSheet.Shapes("LinkRight1").Left = ActiveCell.Left
        
    ActiveSheet.Shapes("LinkRight2").Top = ActiveCell.Top
    ActiveSheet.Shapes("LinkRight2").Left = ActiveCell.Left
        
    ActiveSheet.Shapes("LinkLeft1").Top = ActiveCell.Top
    ActiveSheet.Shapes("LinkLeft1").Left = ActiveCell.Left
    
    ActiveSheet.Shapes("LinkLeft2").Top = ActiveCell.Top
    ActiveSheet.Shapes("LinkLeft2").Left = ActiveCell.Left
    
    ActiveSheet.Shapes("LinkDown1").Top = ActiveCell.Top
    ActiveSheet.Shapes("LinkDown1").Left = ActiveCell.Left
    
    ActiveSheet.Shapes("LinkDown2").Top = ActiveCell.Top
    ActiveSheet.Shapes("LinkDown2").Left = ActiveCell.Left
    
    ActiveSheet.Shapes("LinkUp1").Top = ActiveCell.Top
    ActiveSheet.Shapes("LinkUp1").Left = ActiveCell.Left
    
    ActiveSheet.Shapes("LinkUp2").Top = ActiveCell.Top
    ActiveSheet.Shapes("LinkUp2").Left = ActiveCell.Left
    
    LinkSprite.Top = ActiveCell.Top
    LinkSprite.Left = ActiveCell.Left
    
    LinkSpriteLeft = LinkSprite.Left
    LinkSpriteTop = LinkSprite.Top
    
    linkCellAddress = LinkSprite.TopLeftCell.Address
    CodeCell = ""
    
    'Sheets("Data").Range("C8").Value = linkCellAddress
    
Call alignScreen

Range("A1").Copy Range("A2")

Call calculateScreenLocation("1", "D")

On Error GoTo endPoint

mySub = currentScreen 'global
Application.Run mySub

Exit Sub

endPoint:
MsgBox "Screen setup macro not found: " & mySub


End Sub