Option Explicit

'###################################################################################
'                              EXCELDA II - MAIN GAME LOOP
'###################################################################################

' Win32 API Declarations
Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Module-level variables
Private m_GameState As GameState
Private m_SpriteManager As SpriteManager
Private m_ActionManager As ActionManager
Private m_EnemyManager As EnemyManager

'###################################################################################
'                              Main Game Loop
'###################################################################################

Public Sub RunGame()
    On Error GoTo ErrorHandler
    
    InitializeGame
    ' Call calculateScreenLocation(linkDirection, linkCell)

    GameLoop
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in RunGame: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Game Error"
    Application.CutCopyMode = False
    Sheets(SHEET_TITLE).Activate
End Sub

Private Sub InitializeGame()
    On Error GoTo InitializeError
    
    ' Initialize focused managers
    Set m_GameState = GameStateInstance()
    Set m_SpriteManager = SpriteManagerInstance()
    Set m_ActionManager = ActionManagerInstance()
    Set m_EnemyManager = EnemyManagerInstance()
    
    Exit Sub
    
InitializeError:
    MsgBox "Error initializing game: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Initialization Error"
End Sub

Private Sub GameLoop()
    On Error GoTo GameLoopError
    
    Do
        ' Check for quit condition
        If CheckQuitGame Then Exit Do
        
        ' Update game timers
        UpdateTimers
        
        ' Handle special game states
        If HandleCollisionState Then GoTo ContinueLoop
        If HandleFallingState Then GoTo ContinueLoop
        
        ' Process game logic
        HandleMovementInput
        UpdateSpriteFrames
        HandleActionInput
        HandleTriggers
        HandleEnemies

        ' Check for collisions after movement
        If CheckCollision Then GoTo ContinueLoop

        ' Update visual elements
        UpdateSpriteVisibility

ContinueLoop:
        ' Update positions and sync
        UpdateSpritePositions
        SleepAndSync
        
        ' Yield control to Excel
        DoEvents
    Loop
    
    Exit Sub
    
GameLoopError:
    MsgBox "Error in GameLoop: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Game Loop Error"
    Application.CutCopyMode = False
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
    
    ' Update GameState
    m_GameState.MoveDir = newDir
End Sub

Private Sub HandleActionInput()
    m_ActionManager.HandleActionKey KEY_C, m_ActionManager.CItem, m_ActionManager.CPress, RANGE_ACTION_C
    m_ActionManager.HandleActionKey KEY_D, m_ActionManager.DItem, m_ActionManager.DPress, RANGE_ACTION_D
End Sub

'###################################################################################
'                              Sprite Management
'###################################################################################

Private Sub UpdateSpriteFrames()
    On Error GoTo SpriteFrameError
    
    ' Update sprite frame and position through SpriteManager
    m_SpriteManager.UpdateFrame m_GameState.MoveDir, m_GameState.MoveSpeed
    
    Exit Sub
    
SpriteFrameError:
    MsgBox "Error in UpdateSpriteFrames: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Sprite Error"
End Sub

Private Sub UpdateSpritePositions()
    ' Update sprite positions through SpriteManager
    m_SpriteManager.UpdatePosition
End Sub

'###################################################################################
'                              Helper Functions
'###################################################################################

Private Function CheckQuitGame() As Boolean
    If GetAsyncKeyState(KEY_Q) <> 0 Then
        Call DestroyAllManagers

        Application.CutCopyMode = False
        Sheets(SHEET_TITLE).Activate
        ActiveSheet.Range("A1").Select
        CheckQuitGame = True
    End If
End Function

Private Sub UpdateTimers()
    If m_GameState.ScreenSetUpTimer > 0 Then m_GameState.ScreenSetUpTimer = m_GameState.ScreenSetUpTimer - 1
End Sub

Private Function HandleCollisionState() As Boolean
    If m_GameState.RNDBounceback <> "" Then
        Call BounceBack(m_SpriteManager.LinkSprite, ActiveSheet.Shapes(m_GameState.CollidedWith))
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

'###################################################################################
'                              Trigger System
'###################################################################################

Private Sub HandleTriggers()
    ' Update cell references
    Dim currentCellAddress As String
    currentCellAddress = m_SpriteManager.LinkSprite.TopLeftCell.Address
    
    ' Update GameState
    m_GameState.LinkCellAddress = currentCellAddress
    
    ' Store current location
    Sheets(SHEET_DATA).Range("C18").Value = currentCellAddress
    
    ' Get and process code cell
    Dim codeCellValue As String
    codeCellValue = Range(currentCellAddress).Offset(3, 2).Value
    
    ' Update GameState
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
End Sub

'###################################################################################
'                              Enemy Management
'###################################################################################

Private Sub HandleEnemies()
    m_EnemyManager.HandleEnemy 1, m_SpriteManager.LinkSprite
    m_EnemyManager.HandleEnemy 2, m_SpriteManager.LinkSprite
    m_EnemyManager.HandleEnemy 3, m_SpriteManager.LinkSprite
    m_EnemyManager.HandleEnemy 4, m_SpriteManager.LinkSprite
End Sub

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

'###################################################################################
'                              Sprite Visibility Management
'###################################################################################

Private Sub UpdateSpriteVisibility()
    ' Update sprite visibility through SpriteManager
    m_SpriteManager.UpdateVisibility
    
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

Sub Relocate(ByVal location As String)
    On Error GoTo RelocateError
    
    Dim targetCell As Range
    Dim cellAdd As String
    
    ' Determine target cell
    If location = Sheets(SHEET_DATA).Range("C8").Value Then
        Set targetCell = Range(location)
        
        ' Apply offset based on direction
        Select Case Sheets(SHEET_DATA).Range("C9").Value
            Case "U"
                Set targetCell = targetCell.Offset(-1, 0)
            Case "D"
                Set targetCell = targetCell.Offset(1, 0)
            Case "L"
                Set targetCell = targetCell.Offset(0, -1)
            Case "R"
                Set targetCell = targetCell.Offset(0, 2)
        End Select
    Else
        ' Find cell by searching for the last 4 characters
        cellAdd = Right(location, 4)
        Set targetCell = Cells.Find(What:=cellAdd, After:=ActiveCell, LookIn:=xlFormulas, _
                                   LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                   MatchCase:=True, SearchFormat:=False)
        
        If targetCell Is Nothing Then
            MsgBox "Target cell not found: " & cellAdd, vbCritical, "Relocate Error"
            Exit Sub
        End If
    End If
    
    ' Update all sprite positions
    m_SpriteManager.AlignLinkSprites targetCell.Left, targetCell.Top
    
    ' Update GameState
    m_SpriteManager.LinkSprite.Top = targetCell.Top
    m_SpriteManager.LinkSprite.Left = targetCell.Left
    m_SpriteManager.LinkSpriteLeft = targetCell.Left
    m_SpriteManager.LinkSpriteTop = targetCell.Top
    m_GameState.LinkCellAddress = m_SpriteManager.LinkSprite.TopLeftCell.Address
    m_GameState.CodeCell = ""
    
    ' Screen alignment and setup
    Call alignScreen
    Range("A1").Copy Range("A2")
    Call calculateScreenLocation("1", "D")
    
    ' Run screen setup macro
    On Error GoTo ScreenSetupError
    Dim mySub As String
    mySub = m_GameState.CurrentScreen
    Application.Run mySub
    
    Exit Sub
    
ScreenSetupError:
    MsgBox "Screen setup macro not found: " & mySub, vbCritical, "Screen Setup Error"
    Exit Sub
    
RelocateError:
    MsgBox "Error in Relocate: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Relocate Error"
End Sub