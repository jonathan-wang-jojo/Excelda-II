Option Explicit

' Win32 API Declarations
Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Module-level variables
Private m_ScreenSetUpTimer As Long
Private m_GameState As GameState
Private m_SpriteManager As SpriteManager

' Action state type
Private Type ActionState
    CItem As String
    DItem As String
    CPress As Long
    DPress As Long
End Type

Private m_Actions As ActionState

' References to action-related sprites
Private Type ActionSprites
    SwordFrame1 As Object
    SwordFrame2 As Object
    SwordFrame3 As Object
    ShieldSprite As Object
End Type

Private m_ActionSprites As ActionSprites


'###################################################################################
'
'
'
'
'
'###################################################################################

'###################################################################################
'                              Main Game Loop
'###################################################################################

Public Sub RunGame()
    InitializeGame
    GameLoop
End Sub

Private Sub InitializeGame()
    ' Initialize GameState singleton
    Set m_GameState = New GameState
    Set m_SpriteManager = New SpriteManager
    
    ' Initialize action state
    With m_Actions
        .CItem = Sheets(SHEET_DATA).Range(RANGE_C_ITEM).Value
        .DItem = Sheets(SHEET_DATA).Range(RANGE_D_ITEM).Value
        .CPress = 0
        .DPress = 0
    End With
    
    ' Initialize action sprites
    With m_ActionSprites
        Set .SwordFrame1 = ActiveSheet.Shapes("SwordLeft")
        Set .SwordFrame2 = ActiveSheet.Shapes("SwordSwipeDownLeft")
        Set .SwordFrame3 = ActiveSheet.Shapes("SwordDown")
        Set .ShieldSprite = ActiveSheet.Shapes("LinkShieldDown")
    End With
    
    ' Initialize screen setup timer
    m_ScreenSetUpTimer = 0
    
    ' Keep these global for now as other modules depend on them
    CItem = m_Actions.CItem
    DItem = m_Actions.DItem
    Set SwordFrame1 = m_ActionSprites.SwordFrame1
    Set SwordFrame2 = m_ActionSprites.SwordFrame2
    Set SwordFrame3 = m_ActionSprites.SwordFrame3
    Set shieldSprite = m_ActionSprites.ShieldSprite
End Sub


Private Sub GameLoop()
    Do
        If CheckQuitGame Then Exit Sub
        
        UpdateTimers
        
        ' Handle special states
        If HandleCollisionState Then GoTo ContinueLoop
        If HandleFallingState Then GoTo ContinueLoop
        
        ' Handle input and movement
        HandleMovementInput
        UpdateSpriteFrames
        HandleActionInput
        
ContinueLoop:
        UpdateSpritePositions
        SleepAndSync
    Loop
End Sub

'###################################################################################
'                              Input Handling
'###################################################################################

Private Sub HandleMovementInput()
    Dim newDir As String
    
    ' Check movement keys
    If GetAsyncKeyState(KEY_LEFT) <> 0 Then newDir = newDir & "L"
    If GetAsyncKeyState(KEY_RIGHT) <> 0 Then newDir = newDir & "R"
    If GetAsyncKeyState(KEY_DOWN) <> 0 Then newDir = newDir & "D"
    If GetAsyncKeyState(KEY_UP) <> 0 Then newDir = newDir & "U"
    
    ' Update movement state
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = newDir
    
    ' Update global state for compatibility
    moveDir = newDir
    If newDir <> "" Then lastDir = newDir
    
    ' Update GameState
    m_GameState.MoveDir = newDir
End Sub

Private Sub HandleActionInput()
    HandleActionKey KEY_C, m_Actions.CItem, m_Actions.CPress, RANGE_ACTION_C
    HandleActionKey KEY_D, m_Actions.DItem, m_Actions.DPress, RANGE_ACTION_D
    
    ' Update global state for compatibility
    CPress = m_Actions.CPress
    DPress = m_Actions.DPress
End Sub

Private Sub HandleActionKey(ByVal keyCode As Integer, ByVal item As String, ByRef pressCounter As Long, ByVal flagCell As String)
    If GetAsyncKeyState(keyCode) <> 0 Then
        Sheets(SHEET_DATA).Range(flagCell).Value = "Y"
        pressCounter = pressCounter + 1
        
        Select Case item
            Case "Sword"
                Call SwordSwipe(IIf(keyCode = KEY_C, 1, 2))
            Case "Shield"
                Call ShowShield
        End Select
    Else
        If item = "Sword" Then
            If pressCounter >= 20 Then Call SwordSpin
            With m_ActionSprites
                .SwordFrame1.Visible = False
                .SwordFrame2.Visible = False
                .SwordFrame3.Visible = False
            End With
        ElseIf item = "Shield" Then
            m_ActionSprites.ShieldSprite.Visible = False
            Sheets(SHEET_DATA).Range(RANGE_SHIELD_STATE).Value = ""
        End If
        
        Sheets(SHEET_DATA).Range(flagCell).Value = ""
        pressCounter = 0
    End If
End Sub

'###################################################################################
'                              Sprite Management
'###################################################################################

Private Sub UpdateSpriteFrames()
    Dim currentFrame As Integer
    currentFrame = IIf(Sheets(SHEET_DATA).Range(RANGE_FRAME_COUNT).Value >= 5, 1, 2)
    
    ' Update sprite frame and position based on movement
    With m_SpriteManager
        .UpdateFrame m_GameState.MoveDir
        
        Select Case m_GameState.MoveDir
            Case "U": .Top = .Top - m_GameState.MoveSpeed
            Case "D": .Top = .Top + m_GameState.MoveSpeed
            Case "R": .Left = .Left + m_GameState.MoveSpeed
            Case "L": .Left = .Left - m_GameState.MoveSpeed
        End Select
        
        ' Update global state for compatibility
        Set LinkSprite = .LinkSprite
        LinkSpriteTop = .Top
        LinkSpriteLeft = .Left
    End With
End Sub

Private Sub UpdateSpritePositions()
    ' Update sprite positions through SpriteManager
    m_SpriteManager.UpdatePosition
    
    ' Update global state for compatibility
    LinkSpriteTop = m_SpriteManager.Top
    LinkSpriteLeft = m_SpriteManager.Left
    
    Call AlignLinkSprites(LinkSpriteLeft, LinkSpriteTop)
End Sub

'###################################################################################
'                              Helper Functions
'###################################################################################

Private Function CheckQuitGame() As Boolean
    If GetAsyncKeyState(KEY_Q) <> 0 Then
        Application.CutCopyMode = False
        Sheets(SHEET_TITLE).Activate
        ActiveSheet.Range("A1").Select
        CheckQuitGame = True
    End If
End Function

Private Sub UpdateTimers()
    If m_ScreenSetUpTimer > 0 Then m_ScreenSetUpTimer = m_ScreenSetUpTimer - 1
    screenSetUpTimer = m_ScreenSetUpTimer ' Update global for compatibility
End Sub

Private Function HandleCollisionState() As Boolean
    If RNDBounceback <> "" Then
        Call BounceBack(m_SpriteManager.LinkSprite, ActiveSheet.Shapes(CollidedWith))
        HandleCollisionState = True
    End If
End Function

Private Function HandleFallingState() As Boolean
    HandleFallingState = (Sheets(SHEET_DATA).Range(RANGE_FALLING).Value = "Y")
    m_GameState.IsFalling = HandleFallingState
End Function

Private Sub SleepAndSync()
    Range("A1").Copy Range("A2")
    Sleep m_GameState.GameSpeed
    Application.CutCopyMode = False
End Sub
        
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


'###################################################################################
'                              Trigger System
'###################################################################################

Private Sub HandleTriggers()
    ' Update cell references
    Dim currentCellAddress As String
    currentCellAddress = m_SpriteManager.LinkSprite.TopLeftCell.Address
    
    ' Update global state for compatibility
    linkCellAddress = currentCellAddress
    m_GameState.LinkCellAddress = currentCellAddress
    
    ' Store current location
    Sheets(SHEET_DATA).Range("C18").Value = currentCellAddress
    
    ' Get and process code cell
    Dim codeCellValue As String
    codeCellValue = Range(currentCellAddress).Offset(3, 2).Value
    
    ' Update global and GameState
    CodeCell = codeCellValue
    m_GameState.CodeCell = codeCellValue
    
    ' Process triggers if code cell has content
    If codeCellValue <> "" Then
        Dim ScrollIndicator As String
        Dim ScrollDir As String
        Dim FallIndicator As String
        Dim ActionIndicator As String
        
        ScrollIndicator = Left(codeCellValue, 1)
        ScrollDir = Mid(codeCellValue, 2, 1)
        FallIndicator = Mid(codeCellValue, 3, 2)
        ActionIndicator = Mid(codeCellValue, 7, 2)
 
    
        ' Handle scroll triggers
        If ScrollIndicator = "S" Then
            Call myScroll(ScrollDir)
        End If
        
        ' Handle movement triggers
        Select Case FallIndicator
            Case "FL"
                Call Falling
            Case "JD"
                Call JumpDown
        End Select
        
        ' Handle special actions
        Select Case ActionIndicator
            Case "RL"
                Call Relocate(codeCellValue)
                Exit Sub  ' Replaces GoTo startSub
                
            Case "ET"
                Call EnemyTrigger(codeCellValue)
                
            Case "SE"
                Call SpecialEventTrigger(codeCellValue)
        End Select
    End If



End If

'###################################################################################
'                              Enemy Management
'###################################################################################

Private Sub HandleEnemies()
    HandleEnemy 1
    HandleEnemy 2
    HandleEnemy 3
    HandleEnemy 4
End Sub

Private Sub HandleEnemy(ByVal enemyIndex As Integer)
    ' Get enemy name based on index
    Dim enemyName As String
    Select Case enemyIndex
        Case 1: enemyName = RNDenemyName1
        Case 2: enemyName = RNDenemyName2
        Case 3: enemyName = RNDenemyName3
        Case 4: enemyName = RNDenemyName4
        Case Else: Exit Sub
    End Select
    
    ' Skip if enemy doesn't exist
    If enemyName = "" Then Exit Sub
    
    ' Process enemy
    Call enemyCollision(m_SpriteManager.LinkSprite, enemyName)
    
    ' Check hit state
    Dim isHit As Boolean
    Select Case enemyIndex
        Case 1: isHit = (RNDenemyHit1 > 0)
        Case 2: isHit = (RNDenemyHit2 > 0)
        Case 3: isHit = (RNDenemyHit3 > 0)
        Case 4: isHit = (RNDenemyHit4 > 0)
    End Select
    
    ' Handle hit or movement
    If isHit Then
        Call enemyBounceBack(enemyIndex)
    Else
        Call RNDEnemyMove(enemyIndex)
    End If
End Sub


'-------------------------------------------------------------------

'###################################################################################
'                              Collision Detection
'###################################################################################

Private Function CheckCollision() As Boolean
    Dim baseCell As Range
    Set baseCell = Range(m_GameState.LinkCellAddress)
    
    Select Case m_GameState.MoveDir
        Case "D"
            CheckCollision = (baseCell.Offset(4, 3).Value = "B")
            
        Case "U"
            CheckCollision = (baseCell.Offset(0, 3).Value = "B")
            
        Case "L"
            CheckCollision = (baseCell.Offset(4, 0).Value = "B")
            
        Case "R"
            CheckCollision = (baseCell.Offset(1, 2).Value = "B") Or _
                           (baseCell.Offset(4, 4).Value = "B")
            
        Case "RU"
            CheckCollision = (baseCell.Offset(0, 3).Value = "B")
            
        Case "LU"
            CheckCollision = (baseCell.Value = "B")
            
        Case "RD"
            CheckCollision = (baseCell.Offset(4, 3).Value = "B")
            
        Case "LD"
            CheckCollision = (baseCell.Offset(4, 0).Value = "B")
            
    End Select
End Function



'MsgBox (LinkSprite.Name)

'###################################################################################
'                              Sprite Visibility Management
'###################################################################################

Private Sub UpdateSpriteVisibility()
    Dim spriteName As String
    spriteName = m_SpriteManager.LinkSprite.Name
    
    ' Hide all sprites first
    Dim directions As Variant
    directions = Array("Up", "Down", "Left", "Right")
    Dim frames As Variant
    frames = Array("1", "2")
    
    Dim dir As Variant, frame As Variant
    For Each dir In directions
        For Each frame In frames
            ActiveSheet.Shapes("Link" & dir & frame).Visible = False
        Next frame
    Next dir
    
    ' Show only the active sprite
    m_SpriteManager.LinkSprite.Visible = True
    
    ' Update animation counter
    UpdateAnimationCounter
End Sub

Private Sub UpdateAnimationCounter()
    Dim currentCount As Long
    currentCount = Sheets(SHEET_DATA).Range(RANGE_FRAME_COUNT).Value
    
    If currentCount >= 10 Then
        Sheets(SHEET_DATA).Range(RANGE_FRAME_COUNT).Value = 0
    Else
        Sheets(SHEET_DATA).Range(RANGE_FRAME_COUNT).Value = currentCount + 1
    End If
End Sub


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