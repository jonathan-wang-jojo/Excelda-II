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

    DisableExcelNavigation
    
    Dim direction As String
    direction = Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value
    If direction = "" Then direction = "D"
    
    ' Find Link sprite
    Dim spriteName As String
    spriteName = FindLinkSprite(screen)
    If spriteName = "" Then Err.Raise vbObjectError + 1, "StartGame", "Link sprite not found"
    
    ' Initialize sprite manager
    m_SpriteManager.Initialize screen, spriteName
    m_SpriteManager.UpdateVisibility
    
    ' Set game state
    m_GameState.CurrentScreen = screen
    m_GameState.MoveDir = direction
    m_GameState.GameSpeed = CLng(Val(Sheets(SHEET_DATA).Range(RANGE_GAME_SPEED).Value))
    
    ' Sync legacy globals
    Set LinkSprite = m_SpriteManager.LinkSprite
    CurrentScreen = screen
    m_GameState.LinkCellAddress = LinkSprite.TopLeftCell.Address
    Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = m_GameState.LinkCellAddress
    
    ' Align view and run screen setup
    Call alignScreen
    On Error Resume Next
    Call calculateScreenLocation("", direction)
    If CurrentScreen <> "" Then Application.Run CurrentScreen
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    RestoreExcelNavigation
    MsgBox "Start Error: " & Err.Description, vbCritical
    Sheets(SHEET_TITLE).Activate
End Sub

Private Sub UpdateLoop()
    ' Main game loop - runs every frame
    On Error GoTo ErrorHandler
    Dim frameDelay As Long
    
    Do
        ' Quit check
        If GetAsyncKeyState(KEY_Q) <> 0 Then Exit Do
        
        ' Update game state
        Call Update
        
        ' Sleep for frame timing
        frameDelay = m_GameState.GameSpeed
        If frameDelay <= 0 Then
            frameDelay = CLng(Val(Sheets(SHEET_DATA).Range(RANGE_GAME_SPEED).Value))
            If frameDelay <= 0 Then frameDelay = DEFAULT_GAME_SPEED
            m_GameState.GameSpeed = frameDelay
        End If
    Sleep frameDelay
    DoEvents
    Loop
    
    ' Cleanup
    Call DestroyAllManagers
    RestoreExcelNavigation
    Sheets(SHEET_TITLE).Activate
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Update Error: " & Err.Description, vbCritical
    Call DestroyAllManagers
    RestoreExcelNavigation
    Sheets(SHEET_TITLE).Activate
End Sub

Private Sub Update()
    ' Core game update - called every frame
    
    ' Update timers
    If m_GameState.ScreenSetUpTimer > 0 Then
        m_GameState.ScreenSetUpTimer = m_GameState.ScreenSetUpTimer - 1
    End If
    
    ' Handle collision bounce
    If m_GameState.RNDBounceback <> "" Then
        Call BounceBack(m_SpriteManager.LinkSprite, ActiveSheet.Shapes(m_GameState.CollidedWith))
        m_GameState.RNDBounceback = ""
    End If
    
    ' Check falling state
    m_GameState.IsFalling = (Sheets(SHEET_DATA).Range(RANGE_FALLING).Value = "Y")
    
    ' Handle input and update
    Call HandleInput
    Call HandleTriggers
    Call UpdateSprites
End Sub

Private Sub HandleInput()
    ' Process player input
    Dim newDir As String
    newDir = ""
    
    ' Check movement keys
    If GetAsyncKeyState(KEY_UP) <> 0 Then newDir = newDir & "U"
    If GetAsyncKeyState(KEY_DOWN) <> 0 Then newDir = newDir & "D"
    If GetAsyncKeyState(KEY_LEFT) <> 0 Then newDir = newDir & "L"
    If GetAsyncKeyState(KEY_RIGHT) <> 0 Then newDir = newDir & "R"
    
    ' Update direction
    Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value = newDir
    m_GameState.MoveDir = newDir
    
    ' Process actions
    m_ActionManager.ProcessAction KEY_C, m_ActionManager.CItem, m_ActionManager.CPress
    m_ActionManager.ProcessAction KEY_D, m_ActionManager.DItem, m_ActionManager.DPress
End Sub

Private Sub UpdateSprites()
    ' Update sprite frames and positions
    m_SpriteManager.UpdateFrame m_GameState.MoveDir, m_GameState.MoveSpeed
    m_SpriteManager.UpdatePosition
    
    ' Sync legacy global
    Set LinkSprite = m_SpriteManager.LinkSprite
End Sub

'###################################################################################
'                              Helper Functions
'###################################################################################

Private Function GetCellAddress(ByVal rng As Range) As String
    ' Get cell address from range (handles multi-cell ranges)
    If rng.Cells.Count > 1 Then
        GetCellAddress = rng.Cells(1, 1).Address
    Else
        GetCellAddress = rng.Address
    End If
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
    Set linkCell = LinkSprite.TopLeftCell
    If linkCell Is Nothing Then Exit Sub
    Set triggerCell = mapSheet.Range(linkCell.Address).Offset(3, 2)

    code = Trim$(CStr(triggerCell.Value))
    If Len(code) < 8 Then Exit Sub
    
    ' Update state
    m_GameState.LinkCellAddress = LinkSprite.TopLeftCell.Address
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
    If scrollInd = "S" Then Call myScroll(scrollDir)
    
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
    m_SpriteManager.AlignSprites targetCell.Left, targetCell.Top
    
    ' Update sprite positions
    m_SpriteManager.LinkSpriteLeft = targetCell.Left
    m_SpriteManager.LinkSpriteTop = targetCell.Top
    
    ' Clear trigger state
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

Private Sub DisableExcelNavigation()
    Application.OnKey "{UP}", "HandleGameKey"
    Application.OnKey "{DOWN}", "HandleGameKey"
    Application.OnKey "{LEFT}", "HandleGameKey"
    Application.OnKey "{RIGHT}", "HandleGameKey"
End Sub

Private Sub RestoreExcelNavigation()
    Application.OnKey "{UP}"
    Application.OnKey "{DOWN}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
End Sub

Public Sub HandleGameKey()
    ' Swallow default navigation - actual input handled via GetAsyncKeyState
End Sub