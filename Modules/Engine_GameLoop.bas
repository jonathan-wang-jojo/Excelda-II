Option Explicit

Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long
Private Declare PtrSafe Function GetKeyState Lib "User32.dll" (ByVal nVirtKey As Long) As Integer

Private m_GameState As GameState
Private m_SpriteManager As SpriteManager
Private m_ActionManager As ActionManager
Private m_EnemyManager As EnemyManager
Private m_SceneManager As SceneManager
Private m_IsRunning As Boolean
Private m_MoveBlocked As Boolean
Private m_PendingStartCell As String
Private m_CustomGameSheet As String
Private m_PostStopActivationSheet As String
Private m_FrameCount As Long

' Excel state restoration
Private m_PrevScreenUpdating As Boolean
Private m_PrevEnableEvents As Boolean
Private m_PrevDisplayStatusBar As Boolean
Private m_PrevCalculation As XlCalculation

Public Sub Start()
    On Error GoTo ErrorHandler
    StartGame
    UpdateLoop
    Exit Sub
    
ErrorHandler:
    Cleanup
    MsgBox "Game Error: " & Err.Description, vbCritical
End Sub

Public Sub StartNewGame(Optional ByVal startCell As String = "")
    ResetGame startCell
    Start
End Sub

Public Sub ContinueGame()
    Start
End Sub

Public Sub ConfigureGameSheet(ByVal sheetName As String)
    m_CustomGameSheet = Trim$(sheetName)
End Sub

Public Sub StartNewGameOnSheet(ByVal sheetName As String, Optional ByVal startCell As String = "")
    ConfigureGameSheet sheetName
    StartNewGame startCell
End Sub

Public Sub ContinueGameOnSheet(ByVal sheetName As String)
    ConfigureGameSheet sheetName
    ContinueGame
End Sub

Private Function GetActiveSheetName() As String
    GetActiveSheetName = IIf(m_CustomGameSheet <> "", m_CustomGameSheet, SHEET_GAME)
End Function

Private Function GetGameWorksheet() As Worksheet
    Dim sheetName As String
    sheetName = GetActiveSheetName()

    On Error Resume Next
    Set GetGameWorksheet = Sheets(sheetName)
    On Error GoTo 0

    If GetGameWorksheet Is Nothing Then
        Err.Raise vbObjectError + 301, "GetGameWorksheet", _
            "Game worksheet '" & sheetName & "' was not found."
    End If
End Function

Private Sub ResetGame(Optional ByVal startCell As String = "")
1   On Error GoTo ResetError

2   Dim prevUpdating As Boolean
3   prevUpdating = Application.ScreenUpdating
4   Application.ScreenUpdating = False

5   InitializeManagers

6   Dim registry As GameRegistry
7   Set registry = GameRegistryInstance()

8   Dim wsGame As Worksheet
9   Set wsGame = GetGameWorksheet()
10  wsGame.Activate

    ' Load game config and apply settings
11  Dim gameConfig As IGameConfig
12  Set gameConfig = registry.GetConfigBySheet(wsGame.Name)
    
    ' Initialize DataCache (pass sheet name to avoid ActiveSheet dependency)
13  Dim dataSheet As Worksheet
14  Set dataSheet = registry.GetGameDataSheet(wsGame.Name)
15  If Not dataSheet Is Nothing Then
16      DataCacheInstance().Initialize dataSheet
    End If

18  ApplySpriteDefinitionsForSheet wsGame

19  Dim startAddress As String
20  startAddress = Trim$(startCell)
21  If startAddress = vbNullString And Not gameConfig Is Nothing Then
22      startAddress = Trim$(gameConfig.StartCell)
23  End If
24  If startAddress = vbNullString Then
25      startAddress = wsGame.Cells(1, 1).Address(False, False)
26  End If

27  Dim spriteName As String
28  spriteName = FindPlayerSprite(wsGame.Name)
29  If spriteName = vbNullString Then
30      Err.Raise vbObjectError + 302, "ResetGame", _
            "Player sprite not found on sheet " & wsGame.Name
31  End If

32  m_SpriteManager.BindPlayerSprite wsGame.Name, spriteName
33  m_SpriteManager.UpdateVisibility
34  m_ActionManager.Initialize
35  m_EnemyManager.Initialize
36  m_SceneManager.ActivateSceneBySheet wsGame.Name

37  m_GameState.RefreshFromDataSheet
38  ApplySheetSpecificTuning wsGame
39  m_GameState.CurrentScreen = wsGame.Name
40  m_GameState.MoveDir = vbNullString
41  m_GameState.IsFalling = False

42  m_PendingStartCell = startAddress
43  ApplyPendingStartState
44  m_SpriteManager.ResyncFramePositions

45  If Not gameConfig Is Nothing Then
46      registry.ApplyConfig gameConfig
47  End If

48  AlignViewport

49  Application.ScreenUpdating = prevUpdating
50  Exit Sub

ResetError:
    Application.ScreenUpdating = prevUpdating
    MsgBox "Reset Error: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Source: " & Err.Source & vbCrLf & _
           "Line: " & Erl, vbCritical, "Reset Game"
End Sub

Private Sub StartGame()
    On Error GoTo ErrorHandler
    
    ResetAllManagers
    InitializeManagers
    
    Dim registry As GameRegistry
    Set registry = GameRegistryInstance()
    
    Dim wsGame As Worksheet
    Set wsGame = GetGameWorksheet()
    wsGame.Activate
    
    ' Load and apply game configuration
    Dim gameConfig As IGameConfig
    Set gameConfig = registry.GetConfigBySheet(wsGame.Name)
    
    ' Initialize DataCache from Data sheet (pass sheet name to avoid ActiveSheet dependency)
    Dim dataSheet As Worksheet
    Set dataSheet = registry.GetGameDataSheet(wsGame.Name)
    If Not dataSheet Is Nothing Then
        DataCacheInstance().Initialize dataSheet
    End If

    If Not gameConfig Is Nothing Then
        registry.ApplyConfig gameConfig
    End If

    ApplySpriteDefinitionsForSheet wsGame

    Dim screen As String
    screen = wsGame.Name

    m_SceneManager.ActivateSceneBySheet screen

    EnterGameMode
    DisableExcelNavigation
    Application.ScreenUpdating = True
    
    ' Set initial direction if none exists
    Dim direction As String
    direction = DataCacheInstance().GetValue(RANGE_MOVE_DIR)
    If direction = vbNullString Then
        direction = "D"
        DataCacheInstance().SetValue RANGE_MOVE_DIR, direction
    End If
    
    Dim spriteName As String
    spriteName = FindPlayerSprite(screen)
    If spriteName = vbNullString Then Err.Raise vbObjectError + 1, "StartGame", "Player sprite not found"
    
    m_SpriteManager.BindPlayerSprite screen, spriteName
    m_SpriteManager.UpdateVisibility
    m_ActionManager.Initialize
    
    m_GameState.RefreshFromDataSheet
    ApplySheetSpecificTuning wsGame
    m_GameState.CurrentScreen = screen
    m_GameState.MoveDir = direction
    m_SpriteManager.ResyncFramePositions
    
    ' Apply pending start location BEFORE reading sprite position
    ' This ensures spawn location takes priority over sprite's physical Excel position
    If m_PendingStartCell <> vbNullString Then
        ApplyPendingStartState
        m_SpriteManager.ResyncFramePositions
    Else
        ' Only read sprite position if no pending start (Continue mode)
        m_GameState.PlayerCellAddress = m_SpriteManager.PlayerSprite.TopLeftCell.Address
        DataCacheInstance().SetValue RANGE_CURRENT_CELL, m_GameState.PlayerCellAddress
    End If
    
    AlignViewport
    On Error Resume Next
    CalculateScreenLocation vbNullString, vbNullString
    Dim initialScreenCode As String
    initialScreenCode = m_GameState.CurrentScreenCode
    If initialScreenCode = vbNullString Then initialScreenCode = m_GameState.CurrentScreen
    If initialScreenCode <> vbNullString Then
        m_SceneManager.ApplyScreen initialScreenCode
    End If
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    m_IsRunning = True
    m_PostStopActivationSheet = vbNullString
    
    Exit Sub
    
ErrorHandler:
    Cleanup
    MsgBox "Start Error: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Source: " & Err.Source & vbCrLf & _
           "Line: " & Erl, vbCritical, "Start Game"
    If SheetExists(SHEET_TITLE) Then Sheets(SHEET_TITLE).Activate
End Sub


'###################################################################################
'                              RUNTIME LOOP
'###################################################################################
' Fixed timestep game loop (60Hz) with interpolated rendering
' Decouples physics/logic updates from rendering for smooth gameplay regardless of framerate
' Based on "Fix Your Timestep" - https://gafferongames.com/post/fix_your_timestep/
'###################################################################################
Private Sub UpdateLoop()
    On Error GoTo ErrorHandler
    If Not m_IsRunning Then Exit Sub
    
    Dim lastTick As Double
    lastTick = Timer
    Dim accumulator As Double
    accumulator = 0#

    Do While m_IsRunning
        ' Calculate elapsed time since last frame
        Dim now As Double
        now = Timer
        Dim frameTime As Double
        frameTime = now - lastTick
        
        ' Handle midnight rollover
        If frameTime < 0# Then frameTime = frameTime + 86400#
        
        ' Cap frame time to prevent spiral of death (lag spike protection)
        If frameTime > FIXED_FRAME_SECONDS * MAX_FRAME_SKIP Then
            frameTime = FIXED_FRAME_SECONDS * MAX_FRAME_SKIP
        End If
        
        lastTick = now
        accumulator = accumulator + frameTime

        ' Fixed timestep update loop (FixedUpdate in Unity)
        Do While accumulator >= FIXED_FRAME_SECONDS
            FixedUpdate FIXED_FRAME_SECONDS
            accumulator = accumulator - FIXED_FRAME_SECONDS
            
            If IsQuitRequested() Or Not m_IsRunning Then Exit Do
        Loop

        If Not m_IsRunning Then Exit Do

        ' Interpolated rendering (smooth visuals between physics steps)
        Dim alpha As Double
        alpha = accumulator / FIXED_FRAME_SECONDS
        Render alpha

        ' Yield control to Excel every 3rd frame (20Hz UI refresh, saves 40-80ms/sec)
        m_FrameCount = m_FrameCount + 1
        If m_FrameCount Mod 3 = 0 Then
            Application.ScreenUpdating = True
            DoEvents
            Application.ScreenUpdating = False
        End If

        If IsQuitRequested() Then m_IsRunning = False
    Loop

    Cleanup
    Exit Sub

ErrorHandler:
    Cleanup
    MsgBox "Update Error: " & Err.Description, vbCritical
End Sub

'###################################################################################
'                              FIXED UPDATE (Physics/Logic)
'###################################################################################
' Called at fixed 60Hz rate regardless of render framerate
' All game logic, physics, and state changes happen here
'###################################################################################
Private Sub FixedUpdate(ByVal deltaTime As Double)
    ' Update frame timing
    m_GameState.BeginFrame deltaTime
    
    ' NOTE: RefreshFromDataSheet removed from hot loop (was reading C19/C4 60x/sec)
    ' Speed/timing config now set once at game start via ApplySheetSpecificTuning
    ' If dynamic speed changes needed, use m_GameState.MoveSpeed property directly
    
    ' Decrement timers
    If m_GameState.ScreenSetUpTimer > 0 Then
        m_GameState.ScreenSetUpTimer = m_GameState.ScreenSetUpTimer - 1
    End If
    
    ' Core game systems (order matters!)
    HandleInput deltaTime
    HandleTriggers
    If Not m_IsRunning Then Exit Sub
    HandleEnemies
    UpdateSprites deltaTime
    UpdateFriendlies deltaTime
    
    ' Update new entity system (runs alongside old systems during migration)
    EntityManagerInstance.UpdateAll deltaTime
End Sub

'###################################################################################
'                              RENDER (Visual Update)
'###################################################################################
' Interpolates sprite positions for smooth rendering between fixed steps
' alpha = progress through current physics frame (0.0 to 1.0)
'###################################################################################
Private Sub Render(ByVal alpha As Double)
    On Error Resume Next
    m_SpriteManager.RenderInterpolated alpha
    ' Apply all queued shape updates in single batch
    BatchRendererInstance().ApplyBatch
    On Error GoTo 0
End Sub

'###################################################################################
'                              LEGACY UPDATE WRAPPER
'###################################################################################
' Kept for backward compatibility - routes to FixedUpdate
'###################################################################################
Private Sub Update(ByVal deltaSeconds As Double)
    FixedUpdate deltaSeconds
End Sub

'###################################################################################
'                              INPUT HANDLING
'###################################################################################
Private Sub HandleInput(ByVal deltaSeconds As Double)
    ' Input polling moved here, DoEvents handled in main loop
    
    Static releaseTimer As Double
    Static bufferedDir As String

    Dim currentCell As Range
    On Error Resume Next
    Set currentCell = m_SpriteManager.PlayerSprite.TopLeftCell
    On Error GoTo 0

    If m_SpriteManager.PlayerSprite Is Nothing Then
        releaseTimer = 0#
        bufferedDir = vbNullString
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
    
    Dim newDir As String
    newDir = vbNullString
    If moveUp Then newDir = newDir & "U"
    If moveDown Then newDir = newDir & "D"
    If moveLeft Then newDir = newDir & "L"
    If moveRight Then newDir = newDir & "R"
    
    Dim attemptedDir As String
    attemptedDir = newDir

    If attemptedDir <> vbNullString Then
        releaseTimer = INPUT_BUFFER_SECONDS
        bufferedDir = attemptedDir
    Else
        If releaseTimer > 0# Then
            releaseTimer = releaseTimer - deltaSeconds
            If releaseTimer > 0# And bufferedDir <> vbNullString Then
                attemptedDir = bufferedDir
            Else
                bufferedDir = vbNullString
                releaseTimer = 0#
            End If
        Else
            bufferedDir = vbNullString
        End If
    End If
    
    Dim blocked As Boolean
    blocked = False
    If attemptedDir <> vbNullString And Not currentCell Is Nothing Then
        blocked = DirectionBlocked(attemptedDir, currentCell)
    End If

    m_MoveBlocked = (attemptedDir <> vbNullString And blocked)
    
    DataCacheInstance().SetValue RANGE_MOVE_DIR, attemptedDir
    m_GameState.MoveDir = attemptedDir

    ' Update action key states (generic)
    m_ActionManager.UpdateKeys
    
    ' Dispatch to game-specific action handlers
    Dim gameConfig As IGameConfig
    Set gameConfig = GameRegistryInstance().GetConfigBySheet(m_GameState.CurrentScreen)
    If Not gameConfig Is Nothing Then
        ' Each game config can implement action handling in its lifecycle hooks
        ' For now, Link game uses its own module
        If m_GameState.CurrentScreen = "Game1" Then
            Game_LinkActions.ProcessLinkActions m_ActionManager
        End If
    End If
End Sub

Private Sub UpdateSprites(ByVal deltaSeconds As Double)
    Dim movementDir As String
    movementDir = m_GameState.MoveDir
    
    Dim facingDir As String
    facingDir = IIf(movementDir = vbNullString, m_GameState.LastDir, movementDir)

    Dim effectiveDir As String
    effectiveDir = IIf(m_MoveBlocked, vbNullString, movementDir)
    
    m_SpriteManager.UpdateFrame effectiveDir, facingDir, m_GameState.MoveSpeed, deltaSeconds
    m_SpriteManager.UpdatePosition
    m_SpriteManager.UpdateVisibility

    Dim viewport As ViewportManager
    Set viewport = ViewportManagerInstance()
    viewport.MaintainPlayerViewport
    
    On Error Resume Next
    Dim linkCell As Range
    Set linkCell = m_SpriteManager.PlayerSprite.TopLeftCell
    If Not linkCell Is Nothing Then
        m_GameState.PlayerCellAddress = linkCell.Address
    DataCacheInstance().SetValue RANGE_CURRENT_CELL, m_GameState.PlayerCellAddress
    End If
    On Error GoTo 0
    
    ' NOTE: Do NOT clear MoveDir here - HandleInput manages direction lifecycle
    ' Clearing here causes "blinking" movement (1 frame per key press instead of continuous)
    m_MoveBlocked = False
End Sub

Private Sub UpdateFriendlies(ByVal deltaSeconds As Double)
    Dim manager As FriendlyManager
    Set manager = FriendlyManagerInstance()
    manager.Tick deltaSeconds
End Sub

Private Function DirectionBlocked(ByVal direction As String, ByVal baseCell As Range) As Boolean
    On Error Resume Next
    If baseCell Is Nothing Then Exit Function
    
    Dim hasU As Boolean, hasD As Boolean, hasL As Boolean, hasR As Boolean
    hasU = InStr(direction, "U") > 0
    hasD = InStr(direction, "D") > 0
    hasL = InStr(direction, "L") > 0
    hasR = InStr(direction, "R") > 0
    
    Dim blocked As Boolean
    If hasD Then blocked = blocked Or (baseCell.Offset(4, 3).Value = "B")
    If hasU Then blocked = blocked Or (baseCell.Offset(0, 3).Value = "B")
    If hasL Then blocked = blocked Or (baseCell.Offset(4, 0).Value = "B")
    If hasR Then blocked = blocked Or (baseCell.Offset(1, 2).Value = "B") Or (baseCell.Offset(4, 4).Value = "B")
    
    ' Diagonal collision checks
    If hasR And hasU Then blocked = blocked Or (baseCell.Offset(0, 3).Value = "B")
    If hasL And hasU Then blocked = blocked Or (baseCell.Value = "B")
    If hasR And hasD Then blocked = blocked Or (baseCell.Offset(4, 3).Value = "B")
    If hasL And hasD Then blocked = blocked Or (baseCell.Offset(4, 0).Value = "B")
    
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

Private Function FindPlayerSprite(ByVal sheetName As String) As String
    ' Find active sprite on sheet using configured name or auto-discovery
    ' Priority: 1) GameConfig.PlayerSpriteName, 2) Auto-discovered frames
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 303, "FindPlayerSprite", _
                  "Sheet '" & sheetName & "' not found."
    End If

    ' Check if game config specifies a player sprite name
    Dim gameConfig As IGameConfig
    Set gameConfig = GameRegistryInstance().GetConfigBySheet(sheetName)
    
    If Not gameConfig Is Nothing Then
        Dim configSpriteName As String
        configSpriteName = Trim$(gameConfig.PlayerSpriteName)
        
        If configSpriteName <> "" Then
            ' Try the configured sprite name first
            Dim configShape As Shape
            Set configShape = Nothing
            On Error Resume Next
            Set configShape = ws.Shapes(configSpriteName)
            On Error GoTo 0
            If Not configShape Is Nothing Then
                FindPlayerSprite = configSpriteName
                Exit Function
            End If
        End If
    End If

    ' Fall back to auto-discovery
    Dim spriteManager As SpriteManager
    Set spriteManager = SpriteManagerInstance()

    Dim configuredNames As Variant
    configuredNames = spriteManager.GetConfiguredFrameNames()

    ' Search for configured/discovered frame names on sheet
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
                FindPlayerSprite = frameName
                Exit Function
            End If
        End If
    Next candidate
    
    ' No valid sprite found - provide helpful error message
    Dim discoveredCount As Long
    discoveredCount = spriteManager.GetDiscoveredFrameCount()
    
    Dim errorMsg As String
    If discoveredCount > 0 Then
        errorMsg = "Discovered " & discoveredCount & " sprite frames, but none are currently visible on sheet '" & sheetName & "'." & vbCrLf & _
                   "Expected one of: " & Join(configuredNames, ", ")
    Else
        errorMsg = "No player sprites found on sheet '" & sheetName & "'." & vbCrLf & _
                   "Sprite Discovery supports flexible naming patterns:" & vbCrLf & _
                   "  - Single: Player" & vbCrLf & _
                   "  - Directional: PlayerDown, PlayerUp, PlayerLeft, PlayerRight" & vbCrLf & _
                   "  - Animated: PlayerDown1, PlayerDown2, PlayerUp1, PlayerUp2" & vbCrLf & _
                   "  - State-based: PlayerIdleDown1, PlayerRunLeft2, PlayerAttackUp1"
    End If
    
    Err.Raise vbObjectError + 304, "FindPlayerSprite", errorMsg
End Function

Private Sub InitializeManagers()
    Set m_GameState = GameStateInstance()
    Set m_SpriteManager = SpriteManagerInstance()
    Set m_ActionManager = ActionManagerInstance()
    Set m_EnemyManager = EnemyManagerInstance()
    Set m_SceneManager = SceneManagerInstance()
End Sub

Private Sub AlignViewport()
    Dim viewport As ViewportManager
    Set viewport = ViewportManagerInstance()
    viewport.AlignToPlayer
    viewport.RefreshVisibleDimensions
End Sub

Private Sub CalculateScreenLocation(ByVal scrollDir As String, ByVal offsetDir As String)
    ViewportManagerInstance().UpdateScreenLocation scrollDir, offsetDir
End Sub

Private Sub ApplySpriteDefinitionsForSheet(ByVal ws As Worksheet)
    Dim sm As SpriteManager
    Set sm = SpriteManagerInstance()
    
    ' Use discovery engine to auto-detect sprites on sheet with config-defined base name
    Dim config As IGameConfig
    Dim baseName As String
    baseName = "Player"

    Set config = GameRegistryInstance().GetConfigBySheet(ws.Name)
    If Not config Is Nothing Then
        baseName = Trim$(config.PlayerSpriteBaseName)
        If baseName = vbNullString Then baseName = Trim$(config.PlayerSpriteName)
        If baseName = vbNullString Then baseName = "Player"
    End If

    sm.DiscoverSpritesOnSheet ws, baseName
End Sub

Private Sub ApplySheetSpecificTuning(Optional ByVal wsOverride As Worksheet)
    Dim targetSheet As Worksheet

    If wsOverride Is Nothing Then
        Dim currentScreenName As String
        currentScreenName = m_GameState.CurrentScreen
        If currentScreenName <> "" Then
            On Error Resume Next
            Set targetSheet = Sheets(currentScreenName)
            On Error GoTo 0
        End If
    Else
        Set targetSheet = wsOverride
    End If

    If targetSheet Is Nothing Then Exit Sub

    ' Apply game-specific speed from config
    Dim gameConfig As IGameConfig
    Set gameConfig = GameRegistryInstance().GetConfigBySheet(targetSheet.Name)
    If Not gameConfig Is Nothing Then
        m_GameState.MoveSpeed = gameConfig.PlayerSpeed
    End If
End Sub

'###################################################################################
'                              Trigger System
'###################################################################################

Private Sub HandleTriggers()
    ' Trigger System Documentation:
    ' Triggers are parsed from cell values in the format: S[Dir][Fall][XX][Action][Enemy][Dir][Cell]
    ' Position breakdown:
    '   1: Scroll indicator (S = scroll trigger)
    '   2: Scroll direction code (1=V, 2=H, 3=Down, 4=Up, or direct U/D/L/R)
    '   3-4: Fall indicator (FL=fall, JD=jump, XX=none)
    '   5-6: Padding (XX)
    '   7-8: Action code (RL=relocate, ET=enemy, SE=special event, PU=push)
    '   9-10: Enemy type code
    '   11-12: Enemy slot
    '   13: Action direction
    '   14+: Target cell address or identifier
    ' Special case: "TRIGGER" = end screen trigger, "B" = blocked
    
    On Error Resume Next
    
    Dim mapSheet As Worksheet
    Dim linkCell As Range
    Dim triggerCell As Range
    Dim code As String

    Set mapSheet = Sheets(m_GameState.CurrentScreen)
    Set linkCell = m_SpriteManager.PlayerSprite.TopLeftCell
    If linkCell Is Nothing Then Exit Sub
    Set triggerCell = mapSheet.Range(linkCell.Address).Offset(3, 2)

    code = Trim$(CStr(triggerCell.Value))
    If code = vbNullString Then Exit Sub
    If UCase$(code) = "B" Then Exit Sub
    If StrComp(code, "TRIGGER", vbTextCompare) = 0 Then
        HandleEndScreenTrigger
        Exit Sub
    End If
    
    m_GameState.PlayerCellAddress = linkCell.Address
    m_GameState.CodeCell = code
    DataCacheInstance().SetValue RANGE_CURRENT_CELL, m_GameState.PlayerCellAddress
    
    Dim scrollInd As String, scrollDir As String, fallInd As String, actionInd As String
    Dim enemyType As String, enemySlot As String, actionDir As String, actionCell As String

    ParseTriggerCode code, scrollInd, scrollDir, fallInd, actionInd, enemyType, enemySlot, actionDir, actionCell
    m_GameState.TriggerCellAddress = actionCell

    If scrollInd = "S" And scrollDir <> vbNullString Then
        ViewportManagerInstance().HandleScrollTransition scrollDir
        m_ActionManager.Initialize
        m_SpriteManager.UpdateVisibility
    End If

    Select Case fallInd
        Case "FL": Falling
        Case "JD": JumpDown
    End Select

    Select Case actionInd
        Case "PU": ' Push - not implemented yet
        Case "RL": Relocate code: Exit Sub
        Case "ET": EnemyTrigger code
        Case "SE": SpecialEventTrigger code
    End Select
End Sub

Private Sub HandleEndScreenTrigger()
    m_PostStopActivationSheet = "End Screen"
    m_CustomGameSheet = vbNullString
    m_IsRunning = False
End Sub

'###################################################################################
'                              Enemy Management
'###################################################################################

Private Sub HandleEnemies()
    If m_EnemyManager Is Nothing Or m_SpriteManager Is Nothing Then Exit Sub

    Dim i As Long
    For i = 1 To 4
        On Error Resume Next
        m_EnemyManager.ProcessEnemy i, m_SpriteManager.PlayerSprite
        If Err.Number <> 0 Then
            Debug.Print "HandleEnemies ProcessEnemy error: " & Err.Description
            Err.Clear
            Exit For
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub ApplyPendingStartState()
    Dim pending As String
    pending = Trim$(m_PendingStartCell)
    m_PendingStartCell = vbNullString
    If pending = vbNullString Then Exit Sub

    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    Dim dataSheet As Worksheet
    Set dataSheet = GameRegistryInstance().GetGameDataSheet()

    Dim cache As DataCache
    Set cache = DataCacheInstance()

    On Error GoTo StartStateError

    ' Clear carried-over action assignments and overlays before relocating
    If Not dataSheet Is Nothing Then
        With dataSheet
            .Range(RANGE_ACTION_C).Value = ""
            .Range(RANGE_ACTION_D).Value = ""
            .Range(RANGE_C_ITEM).Value = ""
            .Range(RANGE_D_ITEM).Value = ""
            .Range(RANGE_SHIELD_STATE).Value = ""
        End With
    End If

    If Not cache Is Nothing Then
        cache.SetValue RANGE_ACTION_C, ""
        cache.SetValue RANGE_ACTION_D, ""
        cache.SetValue RANGE_C_ITEM, ""
        cache.SetValue RANGE_D_ITEM, ""
        cache.SetValue RANGE_SHIELD_STATE, ""
    End If

    RelocateToSimpleLocation pending
    
    ' Apply batched sprite position changes immediately
    BatchRendererInstance().ApplyBatch

    ' Batch write to minimize COM calls
    If dataSheet Is Nothing Then
        Set dataSheet = GameRegistryInstance().GetGameDataSheet()
    End If

    If Not dataSheet Is Nothing Then
        With dataSheet
            .Range(RANGE_PREVIOUS_CELL).Value = gs.PlayerCellAddress
            .Range(RANGE_PREVIOUS_SCROLL & ":" & RANGE_SHIELD_STATE).ClearContents
            .Range(RANGE_FALLING).Value = "N"
            .Range(RANGE_FALL_SEQUENCE).Value = "N"
        End With
    End If

    ' Keep cache in sync with initial spawn state to avoid legacy fall/bounce flags
    If Not cache Is Nothing Then
        cache.SetValue RANGE_PREVIOUS_CELL, gs.PlayerCellAddress
        cache.SetValue RANGE_PREVIOUS_SCROLL, ""
        cache.SetValue RANGE_SHIELD_STATE, ""
        cache.SetValue RANGE_FALLING, "N"
        cache.SetValue RANGE_FALL_SEQUENCE, "N"
    End If

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
    
    Dim scrollIndicator As String, scrollDir As String, fallIndicator As String
    Dim actionIndicator As String, enemyType As String, enemySlot As String
    Dim offsetDir As String, actionCell As String, targetAddress As String
    Dim trimmedCode As String
    Dim mapSheet As Worksheet
    Dim targetCell As Range
    Dim gs As GameState

    trimmedCode = Trim$(code)
    
    ' Simple relocation if not a scroll trigger
    If trimmedCode = vbNullString Or Mid$(trimmedCode, 1, 1) <> "S" Then
        RelocateToSimpleLocation trimmedCode
        Exit Sub
    End If

    ParseTriggerCode trimmedCode, scrollIndicator, scrollDir, fallIndicator, actionIndicator, enemyType, enemySlot, offsetDir, actionCell

    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    Dim dataSheet As Worksheet
    Set dataSheet = GameRegistryInstance().GetGameDataSheet()
    
    ' Resolve action cell with fallback chain
    If actionCell = vbNullString Then actionCell = gs.TriggerCellAddress
    If actionCell = vbNullString Then actionCell = DataCacheInstance().GetValue(RANGE_CURRENT_CELL)

    If scrollIndicator <> "S" And actionCell = vbNullString Then
        RelocateToSimpleLocation trimmedCode
        Exit Sub
    End If

    ' Resolve target cell
    If gs.CurrentScreen <> vbNullString Then
        On Error Resume Next
        Set mapSheet = Sheets(gs.CurrentScreen)
        On Error GoTo RelocateError
    End If

    Set targetCell = ResolveTargetCell(actionCell, mapSheet)
    If targetCell Is Nothing Then Err.Raise vbObjectError + 101, "Relocate", "Target cell not found: " & actionCell

    Set mapSheet = targetCell.Worksheet
    gs.CurrentScreen = mapSheet.Name

    ' Apply directional offset (offsetDir already uppercase from ParseTriggerCode)
    Select Case offsetDir
        Case "U": Set targetCell = targetCell.Offset(-1, 0)
        Case "D": Set targetCell = targetCell.Offset(1, 0)
        Case "L": Set targetCell = targetCell.Offset(0, -1)
        Case "R": Set targetCell = targetCell.Offset(0, 2)
    End Select

    PerformRelocation targetCell, gs, scrollDir, offsetDir
    Exit Sub

RelocateError:
    MsgBox "Error in Relocate: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Relocate Error"
End Sub

Private Sub RelocateToSimpleLocation(ByVal location As String)
    On Error GoTo RelocateSimpleError
    location = Trim$(location)
    If location = vbNullString Then Exit Sub

    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    Dim ws As Worksheet
    If gs.CurrentScreen <> vbNullString Then
        On Error Resume Next
        Set ws = Sheets(gs.CurrentScreen)
        On Error GoTo RelocateSimpleError
    End If

    Dim targetCell As Range
    Set targetCell = ResolveTargetCell(location, ws)
    If targetCell Is Nothing Then Exit Sub

    gs.CurrentScreen = targetCell.Worksheet.Name
    PerformRelocation targetCell, gs, "", ""
    Exit Sub

RelocateSimpleError:
    Debug.Print "RelocateToSimpleLocation error: " & Err.Description
End Sub

Private Sub PerformRelocation(ByVal targetCell As Range, ByVal gs As GameState, ByVal scrollDir As String, ByVal offsetDir As String)
    ' Update sprite positions
    m_SpriteManager.AlignSprites targetCell.Left, targetCell.Top
    m_SpriteManager.PlayerSpriteLeft = targetCell.Left
    m_SpriteManager.PlayerSpriteTop = targetCell.Top
    gs.PlayerCellAddress = targetCell.Address

    ' Update data sheet and state
    Dim dataSheet As Worksheet
    Set dataSheet = GameRegistryInstance().GetGameDataSheet()
    DataCacheInstance().SetValue RANGE_CURRENT_CELL, gs.PlayerCellAddress
    DataCacheInstance().SetValue RANGE_MOVE_DIR, ""
    gs.MoveDir = ""
    gs.CodeCell = ""

    AlignViewport

    On Error Resume Next
    CalculateScreenLocation scrollDir, offsetDir
    On Error GoTo 0

    ' Refresh managers
    m_ActionManager.Initialize
    m_SpriteManager.UpdateVisibility
    m_SpriteManager.ResyncFramePositions

    ' Apply screen setup
    Dim setupMacro As String
    setupMacro = gs.CurrentScreenCode
    If setupMacro = vbNullString Then setupMacro = gs.CurrentScreen
    If setupMacro <> vbNullString Then
        On Error GoTo ScreenSetupError
        m_SceneManager.ApplyScreen setupMacro
        On Error GoTo 0
    End If
    Exit Sub

ScreenSetupError:
    Debug.Print "PerformRelocation screen setup error: " & Err.Description
End Sub

Private Function ResolveTargetCell(ByVal location As String, ByVal defaultSheet As Worksheet) As Range
    Dim trimmed As String
    trimmed = Trim$(location)
    If trimmed = vbNullString Then Exit Function

    ' Parse sheet!address format
    Dim sheetPart As String, addressPart As String
    Dim bangPos As Long
    bangPos = InStr(trimmed, "!")
    If bangPos > 0 Then
        sheetPart = Replace(Mid$(trimmed, 1, bangPos - 1), "'", vbNullString)
        addressPart = Mid$(trimmed, bangPos + 1)
    Else
        addressPart = trimmed
    End If

    ' Resolve sheet
    Dim candidateSheet As Worksheet
    If sheetPart <> vbNullString Then
        On Error Resume Next
        Set candidateSheet = Sheets(sheetPart)
        On Error GoTo 0
    End If
    If candidateSheet Is Nothing Then Set candidateSheet = defaultSheet

    ' Try direct range address FIRST (e.g., CO598, A1, B10)
    If ShouldAttemptDirectAddress(addressPart) And Not candidateSheet Is Nothing Then
        On Error Resume Next
        Set ResolveTargetCell = candidateSheet.Range(addressPart)
        On Error GoTo 0
        If Not ResolveTargetCell Is Nothing Then Exit Function
    End If

    ' Fall back to legacy cell ID lookup (search for value in cells - e.g., trigger codes)
    Dim legacyId As String
    legacyId = ExtractLegacyCellId(trimmed)
    If legacyId <> vbNullString Then
        Set ResolveTargetCell = FindCellValueAcrossSheets(legacyId, candidateSheet)
        If Not ResolveTargetCell Is Nothing Then Exit Function
    End If

    ' Try full value lookup
    Set ResolveTargetCell = FindCellValueAcrossSheets(trimmed, candidateSheet)
    If Not ResolveTargetCell Is Nothing Then Exit Function

    ' Final fallback: extract last 4 chars as potential ID
    If legacyId = vbNullString And Len(trimmed) > 4 Then
        Dim fallbackId As String
        fallbackId = Mid$(trimmed, Len(trimmed) - 3)
        Set ResolveTargetCell = FindCellValueAcrossSheets(fallbackId, candidateSheet)
    End If
End Function

Private Function ExtractLegacyCellId(ByVal location As String) As String
    ' Extract last 4 chars if they form a valid cell ID (letter followed by digits)
    Dim candidate As String
    Dim locLen As Long
    
    locLen = Len(location)
    If locLen = 0 Then Exit Function
    
    If locLen <= 4 Then
        candidate = location
    Else
        candidate = Mid$(location, locLen - 3, 4)
    End If

    ' Validate first char is A-Z
    If candidate <> vbNullString Then
        Dim firstChar As String
        firstChar = Mid$(candidate, 1, 1)
        If firstChar >= "A" And firstChar <= "Z" Then
            ExtractLegacyCellId = candidate
        End If
    End If
End Function

Private Function ShouldAttemptDirectAddress(ByVal addressPart As String) As Boolean
    ' Check if address looks like a cell reference (e.g., A1, B10) vs code ID (e.g., X014)
    If addressPart = vbNullString Then Exit Function

    ' Find first digit position
    Dim idx As Long, ch As String, firstDigit As Long
    For idx = 1 To Len(addressPart)
        ch = Mid$(addressPart, idx, 1)
        If ch >= "0" And ch <= "9" Then
            firstDigit = idx
            Exit For
        End If
    Next idx

    ' No digits means it's a name requiring lookup
    If firstDigit = 0 Then Exit Function

    ' Leading zeros indicate code IDs (X014), not cell addresses
    ShouldAttemptDirectAddress = (Mid$(addressPart, firstDigit, 1) <> "0" Or Len(Mid$(addressPart, firstDigit)) = 1)
End Function

Private Function FindCellValueAcrossSheets(ByVal lookupValue As String, ByVal preferredSheet As Worksheet) As Range
    Dim match As Range
    If lookupValue = vbNullString Then Exit Function

    If Not preferredSheet Is Nothing Then
        Set match = FindInWorksheet(preferredSheet, lookupValue)
        If Not match Is Nothing Then
            Set FindCellValueAcrossSheets = match
            Exit Function
        End If
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
    ' Parse trigger format: S[Dir][Fall][XX][Action][Enemy][Dir][Cell]
    Dim payload As String
    payload = UCase$(Trim$(CStr(rawCode)))
    Dim payloadLen As Long
    payloadLen = Len(payload)

    scrollIndicator = Mid$(payload, 1, 1)
    scrollDir = Mid$(payload, 2, 1)
    fallIndicator = Mid$(payload, 3, 2)
    actionIndicator = Mid$(payload, 7, 2)
    enemyType = Mid$(payload, 9, 2)
    enemySlot = Mid$(payload, 11, 2)
    actionDirection = Mid$(payload, 13, 1)

    ' Extract action cell with fallback
    If payloadLen >= 14 Then
        actionCell = Trim$(Mid$(payload, 14))
    End If

    If actionCell = vbNullString Then
        Dim actionPos As Long
        actionPos = InStr(payload, actionIndicator)
        If actionPos > 0 Then
            actionCell = Trim$(Mid$(payload, actionPos + Len(actionIndicator)))
        End If
    End If
End Sub

Private Sub DisableExcelNavigation()
    ' Intercept arrow keys and action keys during gameplay
    Const MODULE_NAME As String = "Engine_GameLoop"
    Const HANDLER As String = MODULE_NAME & ".HandleGameKey"
    
    Dim keys As Variant
    keys = Array("{UP}", "{DOWN}", "{LEFT}", "{RIGHT}", "q", "Q", "c", "C", "d", "D")
    
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Application.OnKey CStr(keys(i)), HANDLER
    Next i
End Sub

Private Sub RestoreExcelNavigation()
    Dim keys As Variant
    keys = Array("{UP}", "{DOWN}", "{LEFT}", "{RIGHT}", "q", "Q", "c", "C", "d", "D")
    
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Application.OnKey CStr(keys(i))
    Next i
End Sub

Public Sub HandleGameKey()
    ' Swallow default navigation - actual input handled via GetAsyncKeyState
End Sub

Private Sub EnterGameMode()
    ' Cache current Excel state and optimize for game loop performance
    With Application
        m_PrevScreenUpdating = .ScreenUpdating
        m_PrevDisplayStatusBar = .DisplayStatusBar
        m_PrevCalculation = .Calculation
        m_PrevEnableEvents = .EnableEvents
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayStatusBar = False
        .Calculation = xlCalculationManual
    End With
End Sub

Private Sub ExitGameMode()
    On Error Resume Next
    With Application
        .ScreenUpdating = m_PrevScreenUpdating
        .EnableEvents = m_PrevEnableEvents
        .DisplayStatusBar = m_PrevDisplayStatusBar
        .Calculation = m_PrevCalculation
    End With
    On Error GoTo 0
End Sub

Private Sub Cleanup()
    On Error Resume Next
    m_IsRunning = False
    
    StopGameLoop True

    If m_PostStopActivationSheet <> "" Then
        Sheets(m_PostStopActivationSheet).Activate
        m_CustomGameSheet = ""
        m_PostStopActivationSheet = ""
    End If
End Sub

Private Sub StopGameLoop(Optional ByVal clearCustomSheet As Boolean = True)
    On Error Resume Next
    m_IsRunning = False
    
    ' Flush DataCache before destroying managers
    If DataCacheInstance().IsDirty Then
        DataCacheInstance().Flush GameRegistryInstance().GetGameDataSheet()
    End If
    
    DestroyAllManagers
    ExitGameMode
    RestoreExcelNavigation
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    m_PendingStartCell = ""
    
    If clearCustomSheet Then
        m_CustomGameSheet = ""
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