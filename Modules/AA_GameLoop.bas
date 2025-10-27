'Attribute VB_Name = "AA_GameLoop"
Option Explicit

'###################################################################################
'                              EXCELDA II - MAIN GAME LOOP
'###################################################################################

' Win32 API Declarations
Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long
Private Declare PtrSafe Function GetKeyState Lib "User32.dll" (ByVal nVirtKey As Long) As Integer

' Module-level variables
Private m_GameState As GameState
Private m_SpriteManager As SpriteManager
Private m_ActionManager As ActionManager
Private m_EnemyManager As EnemyManager
Private m_SceneManager As SceneManager
Private m_PreviousScreenUpdating As Boolean
Private m_PreviousEnableEvents As Boolean
Private m_PreviousDisplayStatusBar As Boolean
Private m_PreviousCalculation As XlCalculation
Private m_InGameMode As Boolean
Private m_IsRunning As Boolean
Private m_MoveBlocked As Boolean
Private m_PendingStartCell As String
Private m_CustomGameSheet As String
Private m_StopClearCustomSheetOverride As Variant
Private m_PostStopActivationSheet As String

'###################################################################################
'                              ENTRY POINT
'###################################################################################
' Standard game pattern: Start → Update → Cleanup

Public Sub Start()
    ' Standard game entry point - call this to begin a new game
    On Error GoTo ErrorHandler
    
    ' Setup and start
    Call StartGame
    Call UpdateLoop
    
    Exit Sub
    
ErrorHandler:
    PerformGameStopCleanup
    MsgBox "Game Error: " & Err.Description, vbCritical
End Sub

Public Sub PrepareNewGameStart(Optional ByVal startCell As String = DEFAULT_START_CELL)
    Dim trimmed As String
    trimmed = Trim$(startCell)
    If trimmed = "" Then
        m_PendingStartCell = DEFAULT_START_CELL
    Else
        m_PendingStartCell = trimmed
    End If
End Sub

Public Sub ConfigureGameSheet(ByVal sheetName As String)
    m_CustomGameSheet = Trim$(sheetName)
End Sub

Public Sub ResetGameOnSheet(ByVal sheetName As String, Optional ByVal startCell As String = DEFAULT_START_CELL)
    ConfigureGameSheet sheetName
    ResetGame startCell
End Sub

Public Sub StartNewGameOnSheet(ByVal sheetName As String, Optional ByVal startCell As String = DEFAULT_START_CELL)
    ConfigureGameSheet sheetName
    StartNewGame startCell
End Sub

Private Function ActiveGameSheetName() As String
    If Trim$(m_CustomGameSheet) <> "" Then
        ActiveGameSheetName = Trim$(m_CustomGameSheet)
    Else
        ActiveGameSheetName = SHEET_GAME
    End If
End Function

Private Function ResolveGameWorksheet() As Worksheet
    Dim sheetName As String
    sheetName = ActiveGameSheetName()

    If Not SheetExists(sheetName) Then
        Err.Raise vbObjectError + 201, "AA_GameLoop.ResolveGameWorksheet", _
                  "Game sheet '" & sheetName & "' not found."
    End If

    Set ResolveGameWorksheet = Sheets(sheetName)
End Function

Public Sub StartNewGame(Optional ByVal startCell As String = DEFAULT_START_CELL)
    ' Convenience entry point for menu buttons: reset to a fresh game state, then start the loop
    Call ResetGame(startCell)
    Call Start
End Sub

Public Sub ContinueGame()
    ' Wrapper for menu buttons that should resume without forcing an explicit reset
    Call Start
End Sub

Public Sub ResetGame(Optional ByVal startCell As String = DEFAULT_START_CELL)
    On Error GoTo ResetError

    m_StopClearCustomSheetOverride = Empty
    m_PostStopActivationSheet = ""

    Dim desiredStart As String
    desiredStart = Trim$(startCell)
    If desiredStart = "" Then desiredStart = DEFAULT_START_CELL

    Dim previousUpdating As Boolean
    previousUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim clearCustomSheet As Boolean
    clearCustomSheet = (Trim$(m_CustomGameSheet) = "")

    StopGameLoop clearCustomSheet

    m_IsRunning = False
    m_MoveBlocked = False

    Dim wsGame As Worksheet
    Set wsGame = ResolveGameWorksheet()
    wsGame.Activate

    ResetAllManagers

    Set m_GameState = GameStateInstance()
    Set m_SpriteManager = SpriteManagerInstance()
    Set m_ActionManager = ActionManagerInstance()
    Set m_EnemyManager = EnemyManagerInstance()
    Set m_SceneManager = SceneManagerInstance()

    ApplySpriteDefinitionsForSheet wsGame

    Dim spriteName As String
    spriteName = FindLinkSprite(wsGame.Name)
    If spriteName = "" Then
        Err.Raise vbObjectError + 302, "ResetGame", "Player sprite not found on sheet " & wsGame.Name
    End If

    m_SpriteManager.BindLinkSprite wsGame.Name, spriteName
    m_SpriteManager.UpdateVisibility
    m_ActionManager.Initialize
    m_EnemyManager.Initialize
    m_SceneManager.ActivateSceneBySheet wsGame.Name

    m_GameState.RefreshFromDataSheet
    ApplySheetSpecificTuning wsGame
    m_GameState.CurrentScreen = wsGame.Name
    m_GameState.MoveDir = ""
    m_GameState.IsFalling = False

    m_PendingStartCell = desiredStart
    ApplyPendingStartState
    If Not m_SpriteManager Is Nothing Then m_SpriteManager.ResyncFramePositions

    Dim viewport As ViewportManager
    Set viewport = ViewportManagerInstance()
    viewport.AlignToLink
    viewport.RefreshVisibleDimensions

    Application.ScreenUpdating = previousUpdating
    Exit Sub

ResetError:
    Application.ScreenUpdating = previousUpdating
    MsgBox "Reset Error: " & Err.Description, vbCritical, "Reset Game"
End Sub
'###################################################################################
'                              STARTUP SEQUENCE
'###################################################################################
Private Sub StartGame()
    ' Initialize everything needed for a new game
    On Error GoTo ErrorHandler
    
    m_StopClearCustomSheetOverride = Empty
    m_PostStopActivationSheet = ""

    ' Reset managers
    Call ResetAllManagers
    
    ' Get manager instances
    Set m_GameState = GameStateInstance()
    Set m_SpriteManager = SpriteManagerInstance()
    Set m_ActionManager = ActionManagerInstance()
    Set m_EnemyManager = EnemyManagerInstance()
    Set m_SceneManager = SceneManagerInstance()
    
    ' Setup starting state
    Dim wsGame As Worksheet
    Set wsGame = ResolveGameWorksheet()
    wsGame.Activate

    ApplySpriteDefinitionsForSheet wsGame

    Dim screen As String
    screen = wsGame.Name

    m_SceneManager.ActivateSceneBySheet screen

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
    m_SpriteManager.BindLinkSprite screen, spriteName
    m_SpriteManager.UpdateVisibility
    m_ActionManager.Initialize
    
    ' Set game state
    m_GameState.RefreshFromDataSheet
    ApplySheetSpecificTuning wsGame
    m_GameState.CurrentScreen = screen
    m_GameState.MoveDir = direction
    If Not m_SpriteManager Is Nothing Then m_SpriteManager.ResyncFramePositions
    
    ' Sync state
    m_GameState.LinkCellAddress = m_SpriteManager.LinkSprite.TopLeftCell.Address
    Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = m_GameState.LinkCellAddress

    If m_PendingStartCell <> "" Then
        ApplyPendingStartState
        If Not m_SpriteManager Is Nothing Then m_SpriteManager.ResyncFramePositions
    End If
    
    ' Align view and run screen setup
    Call alignScreen
    On Error Resume Next
    Call calculateScreenLocation("", "")
    Dim initialScreenCode As String
    initialScreenCode = m_GameState.CurrentScreenCode
    If initialScreenCode = "" Then initialScreenCode = m_GameState.CurrentScreen
    If initialScreenCode <> "" Then
        m_SceneManager.ApplyScreen initialScreenCode
    End If
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    m_IsRunning = True
    
    Exit Sub
    
ErrorHandler:
    PerformGameStopCleanup
    MsgBox "Start Error: " & Err.Description, vbCritical
    If SheetExists(SHEET_TITLE) Then Sheets(SHEET_TITLE).Activate
End Sub

'###################################################################################
'                              RUNTIME LOOP
'###################################################################################
Private Sub UpdateLoop()
    On Error GoTo ErrorHandler
    If Not m_IsRunning Then Exit Sub

    Dim targetStep As Double
    targetStep = FIXED_FRAME_SECONDS

    Dim lastTick As Double
    lastTick = Timer
    Dim accumulator As Double
    accumulator = 0#

    Do While m_IsRunning
        Dim now As Double
        now = Timer
        Dim elapsed As Double
        elapsed = now - lastTick
        If elapsed < 0# Then elapsed = elapsed + 86400#
        accumulator = accumulator + elapsed
        If accumulator > targetStep * 5# Then accumulator = targetStep * 5#
        lastTick = now

        Do While accumulator >= targetStep And m_IsRunning
            Dim deltaSeconds As Double
            deltaSeconds = m_GameState.BeginFrame(targetStep)
            m_GameState.RefreshFromDataSheet
            ApplySheetSpecificTuning
            Update deltaSeconds

            accumulator = accumulator - targetStep
            If IsQuitRequested() Then
                m_IsRunning = False
                Exit Do
            End If
        Loop

        If Not m_IsRunning Then Exit Do

        ' Render interpolated visuals between fixed logic updates
        Dim alpha As Double
        alpha = 0#
        If targetStep > 0# Then alpha = accumulator / targetStep
        If alpha < 0# Then alpha = 0#
        If alpha > 1# Then alpha = 1#
        On Error Resume Next
        If Not m_SpriteManager Is Nothing Then m_SpriteManager.RenderInterpolated alpha
        On Error GoTo ErrorHandler

    Dim wasUpdating As Boolean
    wasUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = True
        Application.CutCopyMode = False
        DoEvents
    Application.ScreenUpdating = wasUpdating

        If IsQuitRequested() Then
            m_IsRunning = False
            Exit Do
        End If

    Loop

    PerformGameStopCleanup
    Exit Sub

ErrorHandler:
    PerformGameStopCleanup
    MsgBox "Update Error: " & Err.Description, vbCritical
End Sub

'###################################################################################
'                              PER-FRAME UPDATE
'###################################################################################
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
    Call HandleInput(deltaSeconds)
    Call HandleTriggers
    If Not m_IsRunning Then Exit Sub
    Call HandleEnemies
    Call UpdateSprites(deltaSeconds)
End Sub

'###################################################################################
'                              INPUT HANDLING
'###################################################################################
Private Sub HandleInput(ByVal deltaSeconds As Double)
    ' Process player input
    DoEvents
    Dim newDir As String
    newDir = ""
    Dim currentCell As Range
    On Error Resume Next
    Set currentCell = m_SpriteManager.LinkSprite.TopLeftCell
    On Error GoTo 0

    Static releaseTimer As Double
    Static bufferedDir As String

    If m_SpriteManager.LinkSprite Is Nothing Then
        releaseTimer = 0#
        bufferedDir = ""
    End If
    
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
    
    ' Determine intended movement direction and apply buffering
    Dim attemptedDir As String
    attemptedDir = newDir

    If attemptedDir <> "" Then
        releaseTimer = INPUT_BUFFER_SECONDS
        bufferedDir = attemptedDir
    Else
        If releaseTimer > 0# Then
            releaseTimer = releaseTimer - deltaSeconds
            If releaseTimer > 0# And bufferedDir <> "" Then
                attemptedDir = bufferedDir
            Else
                bufferedDir = ""
                releaseTimer = 0#
            End If
        Else
            bufferedDir = ""
        End If
    End If
    
    ' Evaluate collision state after resolving buffered intent
    Dim blocked As Boolean
    blocked = False
    If attemptedDir <> "" And Not currentCell Is Nothing Then
        blocked = DirectionBlocked(attemptedDir, currentCell)
    End If

    m_MoveBlocked = (attemptedDir <> "" And blocked)
    
    ' Update direction
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = attemptedDir
    m_GameState.MoveDir = attemptedDir

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

    Dim effectiveDir As String
    If m_MoveBlocked Then
        effectiveDir = ""
    Else
        effectiveDir = movementDir
    End If
    Dim moveSpeed As Double
    moveSpeed = m_GameState.MoveSpeed
    m_SpriteManager.UpdateFrame effectiveDir, facingDir, moveSpeed, deltaSeconds
    m_SpriteManager.UpdatePosition
    m_SpriteManager.UpdateVisibility

    Dim viewport As ViewportManager
    Set viewport = ViewportManagerInstance()
    If Not viewport Is Nothing Then
        viewport.MaintainLinkViewport
    End If
    On Error Resume Next
    Dim linkCell As Range
    Set linkCell = m_SpriteManager.LinkSprite.TopLeftCell
    If Not linkCell Is Nothing Then
        m_GameState.LinkCellAddress = linkCell.Address
        Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = m_GameState.LinkCellAddress
    End If
    On Error GoTo 0
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = ""
    m_GameState.MoveDir = ""
    m_MoveBlocked = False
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
    IsQuitRequested = IsKeyPressed(KEY_Q)
End Function

Private Function IsKeyPressed(ByVal vKey As Integer) As Boolean
    Dim asyncState As Long
    Dim isCurrentlyDown As Boolean
    Dim pressedSinceLastCall As Boolean
    Dim syncState As Long

    asyncState = GetAsyncKeyState(vKey)

    isCurrentlyDown = ((asyncState And &H8000&) <> 0)
    pressedSinceLastCall = ((asyncState And 1) <> 0)

    If Not isCurrentlyDown And Not pressedSinceLastCall Then
        syncState = CLng(GetKeyState(CLng(vKey)))
        isCurrentlyDown = ((syncState And &H8000&) <> 0)
    End If

    IsKeyPressed = isCurrentlyDown Or pressedSinceLastCall
End Function

Private Function FindLinkSprite(ByVal sheetName As String) As String
    ' Find active player sprite on sheet using configured frame names
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim spriteManager As SpriteManager
    Set spriteManager = SpriteManagerInstance()

    Dim configuredNames As Variant
    configuredNames = spriteManager.GetConfiguredFrameNames()

    Dim candidate As Variant
    For Each candidate In configuredNames
        Dim frameName As String
        frameName = Trim$(CStr(candidate))
        If frameName <> "" Then
            Dim frameShape As Shape
            Set frameShape = Nothing
            On Error Resume Next
            Set frameShape = ws.Shapes(frameName)
            On Error GoTo 0
            If Not frameShape Is Nothing Then
                FindLinkSprite = frameName
                Exit Function
            End If
        End If
    Next candidate

    ' Fallback to legacy Link naming in case custom configuration was not provided
    Dim legacyNames As Variant
    legacyNames = Array("LinkDown1", "LinkDown2", "LinkUp1", "LinkUp2", _
                        "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")

    For Each candidate In legacyNames
        Dim legacyName As String
        legacyName = CStr(candidate)
        Dim legacyShape As Shape
        Set legacyShape = Nothing
        On Error Resume Next
        Set legacyShape = ws.Shapes(legacyName)
        On Error GoTo 0
        If Not legacyShape Is Nothing Then
            FindLinkSprite = legacyName
            Exit Function
        End If
    Next candidate
End Function

Private Sub ApplySpriteDefinitionsForSheet(ByVal ws As Worksheet)
    Dim sm As SpriteManager
    Set sm = SpriteManagerInstance()

    ' Reset to defaults before applying overrides for the active sheet.
    sm.Initialize

    If ws Is Nothing Then Exit Sub

    If ws Is Sheet9 Then
        Sheet9.ApplyMinotaurSpriteConfig
    End If
End Sub

Private Sub ApplySheetSpecificTuning(Optional ByVal wsOverride As Worksheet)
    Dim targetSheet As Worksheet

    If wsOverride Is Nothing Then
        Dim currentScreenName As String
        currentScreenName = ""
        If Not m_GameState Is Nothing Then currentScreenName = m_GameState.CurrentScreen
        If currentScreenName <> "" Then
            On Error Resume Next
            Set targetSheet = Sheets(currentScreenName)
            On Error GoTo 0
        End If
    Else
        Set targetSheet = wsOverride
    End If

    If targetSheet Is Nothing Then Exit Sub

    If targetSheet Is Sheet9 Then
        Dim gs As GameState
        If m_GameState Is Nothing Then
            Set gs = GameStateInstance()
        Else
            Set gs = m_GameState
        End If

        If Not gs Is Nothing Then
            gs.MoveSpeed = MINOTAUR_LINK_SPEED
        End If
    End If
End Sub

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
    If code = "" Then Exit Sub
    If UCase$(code) = "B" Then Exit Sub
    If StrComp(code, "TRIGGER", vbTextCompare) = 0 Then
        HandleEndScreenTrigger
        Exit Sub
    End If
    
    ' Update state
    m_GameState.LinkCellAddress = linkCell.Address
    m_GameState.CodeCell = code
    Sheets(SHEET_DATA).Range("C18").Value = m_GameState.LinkCellAddress
    
    ' Parse trigger payload according to legacy format
    Dim scrollInd As String
    Dim scrollDir As String
    Dim fallInd As String
    Dim actionInd As String
    Dim enemyType As String
    Dim enemySlot As String
    Dim actionDir As String
    Dim actionCell As String

    ParseTriggerCode code, scrollInd, scrollDir, fallInd, actionInd, enemyType, enemySlot, actionDir, actionCell

    If Not m_GameState Is Nothing Then
        m_GameState.TriggerCellAddress = actionCell
    End If

    ' Execute scroll
    If scrollInd = "S" And scrollDir <> "" Then
        Call myScroll(scrollDir)
        m_ActionManager.Initialize
        m_SpriteManager.UpdateVisibility
    End If

    ' Execute fall/jump indicators
    Select Case fallInd
        Case "FL": Call Falling
        Case "JD": Call JumpDown
    End Select

    ' Execute action
    Select Case actionInd
        Case "PU": ' Push - not implemented yet
        Case "RL": Call Relocate(code): Exit Sub
        Case "ET": Call EnemyTrigger(code)
        Case "SE": Call SpecialEventTrigger(code)
    End Select
End Sub

Private Sub HandleEndScreenTrigger()
    m_StopClearCustomSheetOverride = False
    m_PostStopActivationSheet = "End Screen"
    m_CustomGameSheet = ""
    m_IsRunning = False
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

Private Sub ApplyPendingStartState()
    Dim pending As String
    pending = Trim$(m_PendingStartCell)
    m_PendingStartCell = ""
    If pending = "" Then Exit Sub

    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    On Error GoTo StartStateError
    RelocateToSimpleLocation pending

    Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_CELL).Value = gs.LinkCellAddress
    Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_SCROLL).Value = ""
    Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value = ""
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = ""
    Sheets(SHEET_DATA).Range(RANGE_FALLING).Value = "N"
    Sheets(SHEET_DATA).Range(RANGE_FALL_SEQUENCE).Value = "N"
    Sheets(SHEET_DATA).Range(RANGE_ACTION_C).Value = ""
    Sheets(SHEET_DATA).Range(RANGE_ACTION_D).Value = ""
    Sheets(SHEET_DATA).Range(RANGE_C_ITEM).Value = ""
    Sheets(SHEET_DATA).Range(RANGE_D_ITEM).Value = ""
    Sheets(SHEET_DATA).Range(RANGE_SHIELD_STATE).Value = ""

    gs.TriggerCellAddress = ""
    gs.CodeCell = ""
    gs.MoveDir = ""
    Exit Sub

StartStateError:
    Debug.Print "ApplyPendingStartState error: " & Err.Description
    On Error GoTo 0
End Sub

Sub Relocate(ByVal code As String)
    On Error GoTo RelocateError

    Dim trimmedCode As String
    Dim scrollIndicator As String
    Dim scrollDir As String
    Dim fallIndicator As String
    Dim actionIndicator As String
    Dim enemyType As String
    Dim enemySlot As String
    Dim offsetDir As String
    Dim targetAddress As String
    Dim actionCell As String
    Dim mapSheet As Worksheet
    Dim targetCell As Range
    Dim gs As GameState

    trimmedCode = Trim$(code)
    If trimmedCode <> "" Then
        If Mid$(trimmedCode, 1, 1) <> "S" Then
            Call RelocateToSimpleLocation(trimmedCode)
            Exit Sub
        End If
    End If

    ParseTriggerCode trimmedCode, scrollIndicator, scrollDir, fallIndicator, actionIndicator, enemyType, enemySlot, offsetDir, actionCell

    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    If actionCell = "" Then
        actionCell = gs.TriggerCellAddress
    End If
    If actionCell = "" Then
        actionCell = Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value
    End If

    If scrollIndicator <> "S" And trimmedCode <> "" And actionCell = "" Then
        Call RelocateToSimpleLocation(trimmedCode)
        Exit Sub
    End If

    targetAddress = actionCell
    offsetDir = UpperCaseText(offsetDir)

    If gs.CurrentScreen <> "" Then
        On Error Resume Next
        Set mapSheet = Sheets(gs.CurrentScreen)
        On Error GoTo RelocateError
    End If

    Set targetCell = ResolveTargetCell(targetAddress, mapSheet)
    If targetCell Is Nothing Then Err.Raise vbObjectError + 101, "Relocate", "Target cell not found: " & targetAddress

    Set mapSheet = targetCell.Worksheet
    gs.CurrentScreen = mapSheet.Name

    Select Case offsetDir
        Case "U": Set targetCell = targetCell.Offset(-1, 0)
        Case "D": Set targetCell = targetCell.Offset(1, 0)
        Case "L": Set targetCell = targetCell.Offset(0, -1)
        Case "R": Set targetCell = targetCell.Offset(0, 2)
    End Select

    m_SpriteManager.AlignSprites targetCell.Left, targetCell.Top
    m_SpriteManager.LinkSpriteLeft = targetCell.Left
    m_SpriteManager.LinkSpriteTop = targetCell.Top
    gs.LinkCellAddress = targetCell.Address

    FinalizeRelocation scrollDir, offsetDir
    Exit Sub

RelocateError:
    MsgBox "Error in Relocate: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Relocate Error"
End Sub

Private Sub RelocateToSimpleLocation(ByVal location As String)
    On Error GoTo RelocateSimpleError
    location = Trim$(location)
    If location = "" Then Exit Sub

    Dim gs As GameState
    Dim ws As Worksheet
    Dim targetCell As Range

    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    If gs.CurrentScreen <> "" Then
        On Error Resume Next
        Set ws = Sheets(gs.CurrentScreen)
        On Error GoTo RelocateSimpleError
    End If

    Set targetCell = ResolveTargetCell(location, ws)
    If targetCell Is Nothing Then Exit Sub

    Set ws = targetCell.Worksheet
    gs.CurrentScreen = ws.Name

    m_SpriteManager.AlignSprites targetCell.Left, targetCell.Top
    m_SpriteManager.LinkSpriteLeft = targetCell.Left
    m_SpriteManager.LinkSpriteTop = targetCell.Top

    gs.LinkCellAddress = targetCell.Address

    FinalizeRelocation "", ""
    Exit Sub

RelocateSimpleError:
    Debug.Print "RelocateToSimpleLocation error: " & Err.Description
End Sub

Private Sub FinalizeRelocation(ByVal scrollDir As String, ByVal offsetDir As String)
    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = gs.LinkCellAddress
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = ""
    gs.MoveDir = ""
    gs.CodeCell = ""

    alignScreen

    On Error Resume Next
    calculateScreenLocation scrollDir, offsetDir
    On Error GoTo 0

    If Not m_ActionManager Is Nothing Then m_ActionManager.Initialize
    If Not m_SpriteManager Is Nothing Then
        m_SpriteManager.UpdateVisibility
        m_SpriteManager.ResyncFramePositions
    End If

    Dim setupMacro As String
    setupMacro = gs.CurrentScreenCode
    If setupMacro = "" Then setupMacro = gs.CurrentScreen
    If setupMacro <> "" And Not m_SceneManager Is Nothing Then
        On Error GoTo ScreenSetupError
        m_SceneManager.ApplyScreen setupMacro
        On Error GoTo 0
    End If
    Exit Sub

ScreenSetupError:
    MsgBox "Screen setup macro not found: " & setupMacro, vbCritical, "Screen Setup Error"
    On Error GoTo 0
End Sub

Private Function ResolveTargetCell(ByVal location As String, ByVal defaultSheet As Worksheet) As Range
    Dim trimmed As String
    trimmed = Trim$(location)
    If trimmed = "" Then Exit Function

    Dim sheetPart As String
    Dim addressPart As String
    Dim bangPos As Long
    bangPos = InStr(trimmed, "!")
    If bangPos > 0 Then
    sheetPart = Replace(Mid$(trimmed, 1, bangPos - 1), "'", "")
        addressPart = Mid$(trimmed, bangPos + 1)
    Else
        addressPart = trimmed
    End If

    Dim candidateSheet As Worksheet
    Dim allowDirectRange As Boolean
    Dim legacyId As String
    Dim found As Range
    If sheetPart <> "" Then
        On Error Resume Next
        Set candidateSheet = Sheets(sheetPart)
        On Error GoTo 0
    End If
    If candidateSheet Is Nothing Then Set candidateSheet = defaultSheet

    legacyId = ExtractLegacyCellId(trimmed)
    If legacyId <> "" Then
        Set found = FindCellValueAcrossSheets(legacyId, candidateSheet)
        If Not found Is Nothing Then
            Set ResolveTargetCell = found
            Exit Function
        End If
    End If

    allowDirectRange = ShouldAttemptDirectAddress(addressPart)

    If allowDirectRange And Not candidateSheet Is Nothing Then
        On Error Resume Next
        Set ResolveTargetCell = candidateSheet.Range(addressPart)
        On Error GoTo 0
        If Not ResolveTargetCell Is Nothing Then Exit Function
    End If

    Set found = FindCellValueAcrossSheets(trimmed, candidateSheet)
    If Not found Is Nothing Then
        Set ResolveTargetCell = found
        Exit Function
    End If

    If legacyId = "" Then
        Dim fallbackId As String
        If Len(trimmed) <= 4 Then
            fallbackId = trimmed
        Else
            fallbackId = Mid$(trimmed, Len(trimmed) - 3)
        End If
        If fallbackId <> "" And fallbackId <> trimmed Then
            Set found = FindCellValueAcrossSheets(fallbackId, candidateSheet)
            If Not found Is Nothing Then
                Set ResolveTargetCell = found
                Exit Function
            End If
        End If
    End If

    Set ResolveTargetCell = Nothing
End Function

Private Function ExtractLegacyCellId(ByVal location As String) As String
    Dim trimmed As String
    Dim candidate As String
    Dim firstChar As String

    trimmed = Trim$(location)
    If trimmed = "" Then Exit Function

    If Len(trimmed) <= 4 Then
        candidate = trimmed
    Else
        candidate = Mid$(trimmed, Len(trimmed) - 3, 4)
    End If

    If candidate = "" Then Exit Function

    firstChar = Mid$(candidate, 1, 1)
    If firstChar < "A" Or firstChar > "Z" Then Exit Function

    ExtractLegacyCellId = candidate
End Function

Private Function ShouldAttemptDirectAddress(ByVal addressPart As String) As Boolean
    Dim trimmed As String
    trimmed = Trim$(addressPart)
    If trimmed = "" Then
        ShouldAttemptDirectAddress = False
        Exit Function
    End If

    Dim idx As Long
    Dim firstDigit As Long
    For idx = 1 To Len(trimmed)
        Dim ch As String
        ch = Mid$(trimmed, idx, 1)
        If ch >= "0" And ch <= "9" Then
            firstDigit = idx
            Exit For
        End If
    Next idx

    If firstDigit = 0 Then
        ' No digits found; treat as name that requires lookup
        ShouldAttemptDirectAddress = False
        Exit Function
    End If

    Dim digitPart As String
    digitPart = Mid$(trimmed, firstDigit)

    If Len(digitPart) > 1 And Mid$(digitPart, 1, 1) = "0" Then
        ' Leading zeros indicate code identifiers like X014; skip direct range
        ShouldAttemptDirectAddress = False
    Else
        ShouldAttemptDirectAddress = True
    End If
End Function

Private Function FindCellValueAcrossSheets(ByVal lookupValue As String, ByVal preferredSheet As Worksheet) As Range
    Dim match As Range
    If lookupValue = "" Then Exit Function

    Set match = FindInWorksheet(preferredSheet, lookupValue)
    If Not match Is Nothing Then
        Set FindCellValueAcrossSheets = match
        Exit Function
    End If

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Not ws Is preferredSheet Then
            Set match = FindInWorksheet(ws, lookupValue)
            If Not match Is Nothing Then
                Set FindCellValueAcrossSheets = match
                Exit Function
            End If
        End If
    Next ws
End Function

Private Function FindInWorksheet(ByVal ws As Worksheet, ByVal lookupValue As String) As Range
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set FindInWorksheet = ws.Cells.Find(What:=lookupValue, After:=ws.Cells(1, 1), LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
    On Error GoTo 0
End Function

Private Sub ParseTriggerCode(ByVal rawCode As String, _
                             ByRef scrollIndicator As String, _
                             ByRef scrollDir As String, _
                             ByRef fallIndicator As String, _
                             ByRef actionIndicator As String, _
                             ByRef enemyType As String, _
                             ByRef enemySlot As String, _
                             ByRef actionDirection As String, _
                             ByRef actionCell As String)
    Dim payload As String
    payload = Trim$(CStr(rawCode))

    scrollIndicator = UpperCaseText(SliceRange(payload, 1, 1))
    scrollDir = UpperCaseText(SliceRange(payload, 2, 1))
    fallIndicator = UpperCaseText(SliceRange(payload, 3, 2))
    actionIndicator = UpperCaseText(SliceRange(payload, 7, 2))
    enemyType = UpperCaseText(SliceRange(payload, 9, 2))
    enemySlot = UpperCaseText(SliceRange(payload, 11, 2))
    actionDirection = UpperCaseText(SliceRange(payload, 13, 1))

    Dim rawActionCell As String
    If Len(payload) >= 14 Then
        rawActionCell = SliceRange(payload, 14)
    End If

    If Trim$(rawActionCell) = "" Then
        Dim actionPos As Long
        actionPos = InStr(payload, actionIndicator)
        If actionPos > 0 Then
            Dim fallbackStart As Long
            fallbackStart = actionPos + Len(actionIndicator)
            rawActionCell = SliceRange(payload, fallbackStart)
        End If
    End If

    actionCell = Trim$(rawActionCell)
End Sub

Private Function SliceRange(ByVal source As String, ByVal startPos As Long, Optional ByVal extractLength As Long = 0) As String
    Dim srcLen As Long
    srcLen = Len(source)
    If startPos <= 0 Or startPos > srcLen Then Exit Function

    If extractLength <= 0 Then
        SliceRange = Mid$(source, startPos)
    Else
        If startPos + extractLength - 1 > srcLen Then
            extractLength = srcLen - startPos + 1
        End If
        SliceRange = Mid$(source, startPos, extractLength)
    End If
End Function

Private Function UpperCaseText(ByVal value As String) As String
    UpperCaseText = UCase$(Trim$(value))
End Function

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

Private Sub PerformGameStopCleanup()
    Dim clearCustomSheet As Boolean
    If IsEmpty(m_StopClearCustomSheetOverride) Then
        clearCustomSheet = True
    Else
        clearCustomSheet = CBool(m_StopClearCustomSheetOverride)
    End If

    StopGameLoop clearCustomSheet

    If m_PostStopActivationSheet <> "" Then
        On Error Resume Next
        Sheets(m_PostStopActivationSheet).Activate
        On Error GoTo 0
        m_CustomGameSheet = ""
    End If

    m_StopClearCustomSheetOverride = Empty
    m_PostStopActivationSheet = ""
End Sub

Private Sub StopGameLoop(Optional ByVal clearCustomSheet As Boolean = True)
    ' Centralized stop/cleanup logic used by the loop and error handlers
    On Error Resume Next
    m_IsRunning = False
    ' Tear down managers and restore Excel UI
    DestroyAllManagers
    Call ExitGameMode
    Call RestoreExcelNavigation
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    m_PendingStartCell = ""
    If clearCustomSheet Then
        m_CustomGameSheet = ""
        ' Try to activate title sheet if present
        If SheetExists(SHEET_TITLE) Then
            Sheets(SHEET_TITLE).Activate
        End If
    End If
    On Error GoTo 0
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function