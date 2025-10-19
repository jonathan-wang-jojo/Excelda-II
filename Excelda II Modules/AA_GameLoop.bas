'Attribute VB_Name = "AA_GameLoop"
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
Private m_PreviousScreenUpdating As Boolean
Private m_PreviousEnableEvents As Boolean
Private m_PreviousDisplayStatusBar As Boolean
Private m_PreviousCalculation As XlCalculation
Private m_InGameMode As Boolean

'###################################################################################
'                              Main Game Loop
'###################################################################################
' Standard game pattern: Start ? Update ? Cleanup

Public Sub Start()
    ' Standard game entry point - call this to begin a new game
    On Error GoTo ErrorHandler
    
    ' Setup and start
    Call StartGame
    Call UpdateLoop
    
    Exit Sub
    
ErrorHandler:
    Call ExitGameMode
    RestoreExcelNavigation
    MsgBox "Game Error: " & Err.Description, vbCritical
    Sheets(SHEET_TITLE).Activate
End Sub

Private Sub StartGame()
    ' Initialize everything needed for a new game
    On Error GoTo ErrorHandler
    
    ' Reset managers
    Call ResetAllManagers
    
    ' Get manager instances
    Set m_GameState = GameStateInstance()
    Set m_SpriteManager = SpriteManagerInstance()
    Set m_ActionManager = ActionManagerInstance()
    Set m_EnemyManager = EnemyManagerInstance()
    
    ' Setup starting state
    Dim screen As String
    screen = ActiveSheet.Name
    If screen = SHEET_TITLE Then screen = SHEET_GAME
    Sheets(screen).Activate

    EnterGameMode
    DisableExcelNavigation
    Application.ScreenUpdating = True
    
    Dim direction As String
    direction = Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value
    If direction = "" Then direction = "D"
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = direction
    
    ' Find Link sprite
    Dim spriteName As String
    spriteName = FindLinkSprite(screen)
    If spriteName = "" Then Err.Raise vbObjectError + 1, "StartGame", "Link sprite not found"
    
    ' Initialize sprite manager
    m_SpriteManager.Initialize screen, spriteName
    m_SpriteManager.UpdateVisibility
    m_ActionManager.Initialize screen
    
    ' Set game state
    m_GameState.RefreshFromDataSheet
    m_GameState.CurrentScreen = screen
    m_GameState.MoveDir = direction
    
    ' Sync state
    m_GameState.LinkCellAddress = m_SpriteManager.LinkSprite.TopLeftCell.Address
    Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = m_GameState.LinkCellAddress
    
    ' Align view and run screen setup
    Call alignScreen
    On Error Resume Next
    Call calculateScreenLocation("", direction)
    If m_GameState.CurrentScreen <> "" Then Application.Run m_GameState.CurrentScreen
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Exit Sub
    
ErrorHandler:
    RestoreExcelNavigation
    Call ExitGameMode
    MsgBox "Start Error: " & Err.Description, vbCritical
    Sheets(SHEET_TITLE).Activate
End Sub

Private Sub UpdateLoop()
    ' Main game loop - runs every frame
    On Error GoTo ErrorHandler
    Dim frameDelay As Long
    
    Do
        ' Quit check
        If IsQuitRequested() Then Exit Do
        Dim deltaSeconds As Double
        deltaSeconds = m_GameState.BeginFrame()
        
        ' Update game state
        m_GameState.RefreshFromDataSheet
        Call Update(deltaSeconds)
        Application.ScreenUpdating = True
        DoEvents
        If IsQuitRequested() Then Exit Do
        
        ' Sleep for frame timing
        frameDelay = m_GameState.GameSpeed
        If frameDelay <= 0 Then
            frameDelay = CLng(Val(Sheets(SHEET_DATA).Range(RANGE_GAME_SPEED).Value))
            If frameDelay <= 0 Then frameDelay = DEFAULT_GAME_SPEED
            m_GameState.GameSpeed = frameDelay
        End If
        Sleep frameDelay
        If IsQuitRequested() Then Exit Do
        DoEvents
        If IsQuitRequested() Then Exit Do
        Application.CutCopyMode = False
        Application.ScreenUpdating = False
    Loop
    
    ' Cleanup
    Call DestroyAllManagers
    Call ExitGameMode
    RestoreExcelNavigation
    Sheets(SHEET_TITLE).Activate
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Update Error: " & Err.Description, vbCritical
    Call DestroyAllManagers
    Call ExitGameMode
    RestoreExcelNavigation
    Sheets(SHEET_TITLE).Activate
End Sub

Private Sub Update(ByVal deltaSeconds As Double)
    ' Core game update - called every frame
    
    ' Update timers
    If m_GameState.ScreenSetUpTimer > 0 Then
        m_GameState.ScreenSetUpTimer = m_GameState.ScreenSetUpTimer - 1
    End If
    
    ' Handle collision bounce
    If m_GameState.RNDBounceback > 0 Then
        Dim bounceSpeed As Long
        bounceSpeed = m_GameState.RNDBounceback
        m_SpriteManager.ApplyLinkBounce bounceSpeed
        m_GameState.RNDBounceback = 0
    End If
    
    ' Check falling state
    m_GameState.IsFalling = (Sheets(SHEET_DATA).Range(RANGE_FALLING).Value = "Y")
    
    ' Handle input and update
    Call HandleInput
    Call HandleTriggers
    Call HandleEnemies
    Call UpdateSprites(deltaSeconds)
End Sub

Private Sub HandleInput()
    ' Process player input
    Dim newDir As String
    newDir = ""
    Dim currentCell As Range
    On Error Resume Next
    Set currentCell = m_SpriteManager.LinkSprite.TopLeftCell
    On Error GoTo 0
    
    Dim moveUp As Boolean, moveDown As Boolean, moveLeft As Boolean, moveRight As Boolean
    moveUp = IsKeyPressed(KEY_UP)
    moveDown = IsKeyPressed(KEY_DOWN)
    moveLeft = IsKeyPressed(KEY_LEFT)
    moveRight = IsKeyPressed(KEY_RIGHT)
    
    If moveUp And moveDown Then
        moveUp = False
        moveDown = False
    End If
    If moveLeft And moveRight Then
        moveLeft = False
        moveRight = False
    End If
    
    If moveUp Then newDir = newDir & "U"
    If moveDown Then newDir = newDir & "D"
    If moveLeft Then newDir = newDir & "L"
    If moveRight Then newDir = newDir & "R"
    
    ' Block movement if collision detected
    If newDir <> "" And Not currentCell Is Nothing Then
        If DirectionBlocked(newDir, currentCell) Then newDir = ""
    End If
    
    ' Update direction
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = newDir
    m_GameState.MoveDir = newDir
    
    ' Process actions
    m_ActionManager.ProcessAction KEY_C
    m_ActionManager.ProcessAction KEY_D
End Sub

Private Sub UpdateSprites(ByVal deltaSeconds As Double)
    ' Update sprite frames and positions
    Dim movementDir As String
    movementDir = m_GameState.MoveDir
    Dim facingDir As String
    If movementDir = "" Then
        facingDir = m_GameState.LastDir
    Else
        facingDir = movementDir
    End If
    m_SpriteManager.UpdateFrame movementDir, facingDir, m_GameState.MoveSpeed, deltaSeconds
    m_SpriteManager.UpdatePosition
    m_SpriteManager.UpdateVisibility
    On Error Resume Next
    Dim linkCell As Range
    Set linkCell = m_SpriteManager.LinkSprite.TopLeftCell
    If Not linkCell Is Nothing Then
        m_GameState.LinkCellAddress = linkCell.Address
        Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = m_GameState.LinkCellAddress
    End If
    On Error GoTo 0
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = ""
End Sub

Private Function DirectionBlocked(ByVal direction As String, ByVal baseCell As Range) As Boolean
    On Error Resume Next
    If baseCell Is Nothing Then Exit Function
    Dim blocked As Boolean
    
    If InStr(direction, "D") > 0 Then
        blocked = blocked Or (baseCell.Offset(4, 3).Value = "B")
    End If
    If InStr(direction, "U") > 0 Then
        blocked = blocked Or (baseCell.Offset(0, 3).Value = "B")
    End If
    If InStr(direction, "L") > 0 Then
        blocked = blocked Or (baseCell.Offset(4, 0).Value = "B")
    End If
    If InStr(direction, "R") > 0 Then
        blocked = blocked Or (baseCell.Offset(1, 2).Value = "B") Or _
                             (baseCell.Offset(4, 4).Value = "B")
    End If
    If InStr(direction, "R") > 0 And InStr(direction, "U") > 0 Then
        blocked = blocked Or (baseCell.Offset(0, 3).Value = "B")
    End If
    If InStr(direction, "L") > 0 And InStr(direction, "U") > 0 Then
        blocked = blocked Or (baseCell.Value = "B")
    End If
    If InStr(direction, "R") > 0 And InStr(direction, "D") > 0 Then
        blocked = blocked Or (baseCell.Offset(4, 3).Value = "B")
    End If
    If InStr(direction, "L") > 0 And InStr(direction, "D") > 0 Then
        blocked = blocked Or (baseCell.Offset(4, 0).Value = "B")
    End If
    DirectionBlocked = blocked
    On Error GoTo 0
End Function

'###################################################################################
'                              Helper Functions
'###################################################################################

Private Function IsQuitRequested() As Boolean
    Dim keyState As Long
    keyState = GetAsyncKeyState(KEY_Q)
    IsQuitRequested = ((keyState And &H8000&) <> 0)
End Function

Private Function IsKeyPressed(ByVal vKey As Integer) As Boolean
    Dim state As Long
    state = GetAsyncKeyState(vKey)
    IsKeyPressed = ((state And &H8000&) <> 0)
End Function

Private Function FindLinkSprite(ByVal sheetName As String) As String
    ' Find Link sprite on sheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Sheets(sheetName)
    
    Dim names As Variant
    names = Array("LinkDown1", "LinkDown2", "LinkUp1", "LinkUp2", _
                  "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
    
    Dim i As Integer
    For i = LBound(names) To UBound(names)
        If Not ws.Shapes(names(i)) Is Nothing Then
            FindLinkSprite = names(i)
            Exit Function
        End If
    Next i

    FindLinkSprite = ""
End Function

'###################################################################################
'                              Trigger System
'###################################################################################

Private Sub HandleTriggers()
    ' Check and execute trigger codes from map cells
    ' Format: S[Dir][Action][Pad][EnemyCode][Dir][Cell]
    ' Example: S1XXXXETOC02DR484
    On Error Resume Next
    
    Dim mapSheet As Worksheet
    Dim linkCell As Range
    Dim triggerCell As Range
    Dim code As String

    Set mapSheet = Sheets(m_GameState.CurrentScreen)
    Set linkCell = m_SpriteManager.LinkSprite.TopLeftCell
    If linkCell Is Nothing Then Exit Sub
    Set triggerCell = mapSheet.Range(linkCell.Address).Offset(3, 2)

    code = Trim$(CStr(triggerCell.Value))
    If Len(code) < 8 Then Exit Sub
    
    ' Update state
    m_GameState.LinkCellAddress = linkCell.Address
    m_GameState.CodeCell = code
    Sheets(SHEET_DATA).Range("C18").Value = m_GameState.LinkCellAddress
    
    ' Parse: S1XXXXETOC02DR484
    '        ││    ││      │└─ Cell (R484)
    '        ││    ││      └─── Direction (D)
    '        ││    │└────────── Enemy code (ETOC02) or padding (XXXXXX)
    '        ││    └─────────── Action (ET/RL/SE/FL/PU)
    '        │└──────────────── Scroll direction (1=Right,2=Left,3=Down,4=Up)
    '        └───────────────── Scroll indicator (S=scroll, X=no scroll)
    
    Dim scrollInd As String: scrollInd = VBA.Left$(code, 1)
    Dim scrollDir As String: scrollDir = Mid$(code, 2, 1)
    Dim actionInd As String: actionInd = Mid$(code, 3, 2)
    
    ' Execute scroll
    If scrollInd = "S" Then
        Call myScroll(scrollDir)
        m_ActionManager.Initialize m_GameState.CurrentScreen
        m_SpriteManager.UpdateVisibility
    End If
    
    ' Execute action
    Select Case actionInd
        Case "FL": Call Falling
        Case "JD": Call JumpDown
        Case "PU": ' Push - not implemented yet
        Case "RL": Call Relocate(code): Exit Sub
        Case "ET": Call EnemyTrigger(code)
        Case "SE": Call SpecialEventTrigger(code)
    End Select
End Sub

'###################################################################################
'                              Enemy Management
'###################################################################################

Private Sub HandleEnemies()
    Dim i As Integer
    For i = 1 To 4
        m_EnemyManager.ProcessEnemy i, m_SpriteManager.LinkSprite
    Next i
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

Sub Relocate(ByVal code As String)
    On Error GoTo RelocateError
    
    Dim scrollDir As String
    Dim offsetDir As String
    Dim targetAddress As String
    Dim mapSheet As Worksheet
    Dim targetCell As Range
    Dim setupMacro As String

    If Len(Trim$(code)) > 0 And Left$(Trim$(code), 1) <> "S" Then
        Call RelocateToSimpleLocation(code)
        Exit Sub
    End If
    
    scrollDir = Mid$(code, 2, 1)
    offsetDir = Mid$(code, 13, 1)
    targetAddress = Mid$(code, 14)
    If targetAddress = "" Then targetAddress = Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value
    
    Set mapSheet = Sheets(m_GameState.CurrentScreen)
    Set targetCell = mapSheet.Range(targetAddress)
    If targetCell Is Nothing Then Err.Raise vbObjectError + 101, "Relocate", "Target cell not found: " & targetAddress
    
    Select Case offsetDir
        Case "U": Set targetCell = targetCell.Offset(-1, 0)
        Case "D": Set targetCell = targetCell.Offset(1, 0)
        Case "L": Set targetCell = targetCell.Offset(0, -1)
        Case "R": Set targetCell = targetCell.Offset(0, 2)
    End Select
    
    m_SpriteManager.AlignSprites targetCell.Left, targetCell.Top
    m_SpriteManager.LinkSpriteLeft = targetCell.Left
    m_SpriteManager.LinkSpriteTop = targetCell.Top
    m_GameState.LinkCellAddress = targetCell.Address
    Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = m_GameState.LinkCellAddress
    m_GameState.CodeCell = ""
    
    Call alignScreen
    Call calculateScreenLocation(scrollDir, offsetDir)
    m_ActionManager.Initialize m_GameState.CurrentScreen
    m_SpriteManager.UpdateVisibility
    
    On Error GoTo ScreenSetupError
    setupMacro = m_GameState.CurrentScreen
    If setupMacro <> "" Then Application.Run setupMacro
    Exit Sub
    
ScreenSetupError:
    MsgBox "Screen setup macro not found: " & setupMacro, vbCritical, "Screen Setup Error"
    Exit Sub
    
RelocateError:
    MsgBox "Error in Relocate: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Relocate Error"
End Sub

Private Sub RelocateToSimpleLocation(ByVal location As String)
    On Error GoTo RelocateSimpleError
    location = Trim$(location)
    If location = "" Then Exit Sub

    Dim gs As GameState
    If m_GameState Is Nothing Then
        Set gs = GameStateInstance()
    Else
        Set gs = m_GameState
    End If
    If gs Is Nothing Or gs.CurrentScreen = "" Then Exit Sub

    Dim ws As Worksheet
    Set ws = Sheets(gs.CurrentScreen)

    Dim dataSheet As Worksheet
    Set dataSheet = Sheets(SHEET_DATA)

    Dim targetCell As Range
    On Error Resume Next
    Set targetCell = ws.Range(location)
    On Error GoTo RelocateSimpleError

    If targetCell Is Nothing Then
        Dim cellId As String
        cellId = Right$(location, 4)
        If cellId <> "" Then
            Set targetCell = ws.Cells.Find(What:=cellId, After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
        End If
    End If

    If targetCell Is Nothing Then Exit Sub

    m_SpriteManager.AlignSprites targetCell.Left, targetCell.Top
    m_SpriteManager.LinkSpriteLeft = targetCell.Left
    m_SpriteManager.LinkSpriteTop = targetCell.Top

    gs.LinkCellAddress = targetCell.Address
    dataSheet.Range(RANGE_CURRENT_CELL).Value = gs.LinkCellAddress
    gs.CodeCell = ""

    Call alignScreen
    On Error Resume Next
    Call calculateScreenLocation("", "")
    On Error GoTo RelocateSimpleError

    Exit Sub

RelocateSimpleError:
    Debug.Print "RelocateToSimpleLocation error: " & Err.Description
End Sub

Private Sub DisableExcelNavigation()
    Application.OnKey "{UP}", "AA_GameLoop.HandleGameKey"
    Application.OnKey "{DOWN}", "AA_GameLoop.HandleGameKey"
    Application.OnKey "{LEFT}", "AA_GameLoop.HandleGameKey"
    Application.OnKey "{RIGHT}", "AA_GameLoop.HandleGameKey"
    Application.OnKey "q", "AA_GameLoop.HandleGameKey"
    Application.OnKey "Q", "AA_GameLoop.HandleGameKey"
    Application.OnKey "c", "AA_GameLoop.HandleGameKey"
    Application.OnKey "C", "AA_GameLoop.HandleGameKey"
    Application.OnKey "d", "AA_GameLoop.HandleGameKey"
    Application.OnKey "D", "AA_GameLoop.HandleGameKey"
End Sub

Private Sub RestoreExcelNavigation()
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "q"
    Application.OnKey "Q"
    Application.OnKey "c"
    Application.OnKey "C"
    Application.OnKey "d"
    Application.OnKey "D"
End Sub

Public Sub HandleGameKey()
    ' Swallow default navigation - actual input handled via GetAsyncKeyState
End Sub

Private Sub EnterGameMode()
    If m_InGameMode Then Exit Sub
    m_PreviousScreenUpdating = Application.ScreenUpdating
    m_PreviousDisplayStatusBar = Application.DisplayStatusBar
    m_PreviousCalculation = Application.Calculation
    m_PreviousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    m_InGameMode = True
End Sub

Private Sub ExitGameMode()
    On Error Resume Next
    If Not m_InGameMode Then Exit Sub
    Application.ScreenUpdating = m_PreviousScreenUpdating
    Application.EnableEvents = m_PreviousEnableEvents
    Application.DisplayStatusBar = m_PreviousDisplayStatusBar
    Application.Calculation = m_PreviousCalculation
    m_InGameMode = False
    On Error GoTo 0
End Sub