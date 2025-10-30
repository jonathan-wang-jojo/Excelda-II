Option Explicit

Private Type GameData
    ' Screen & Scene
    CurrentScreen As String
    CurrentSceneId As String
    CurrentScreenCode As String
    
    ' Timing & Speed
    GameSpeed As Long
    MoveSpeed As Double
    LastFrameTimestamp As Double
    FrameDeltaSeconds As Double
    
    ' Player State
    MoveDir As String
    LastDir As String
    IsPaused As Boolean
    IsFalling As Boolean
    PlayerCellAddress As String
    
    ' Trigger & Event
    CodeCell As String
    TriggerCellAddress As String
    
    ' Timers
    ScreenSetUpTimer As Long
End Type

Private m_Game As GameData

Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Reset()
    m_Game.GameSpeed = ReadGameSpeed()
    m_Game.MoveSpeed = ReadMoveSpeed()
    m_Game.IsPaused = False
    m_Game.MoveDir = ""
    m_Game.LastDir = "D"
    m_Game.IsFalling = False
    m_Game.PlayerCellAddress = ""
    m_Game.CodeCell = ""
    m_Game.ScreenSetUpTimer = 0
    m_Game.TriggerCellAddress = ""
    m_Game.CurrentScreen = ""
    m_Game.CurrentSceneId = ""
    m_Game.CurrentScreenCode = ""
    m_Game.FrameDeltaSeconds = DEFAULT_FRAME_SECONDS
    m_Game.LastFrameTimestamp = Timer
End Sub

Public Sub Initialize()
    Reset
End Sub

Public Sub Destroy()
    ' Clear all state on destruction
    m_Game.GameSpeed = 0
    m_Game.MoveSpeed = 0
    m_Game.IsPaused = False
    m_Game.MoveDir = ""
    m_Game.LastDir = ""
    m_Game.IsFalling = False
    m_Game.PlayerCellAddress = ""
    m_Game.CodeCell = ""
    m_Game.ScreenSetUpTimer = 0
    m_Game.TriggerCellAddress = ""
    m_Game.CurrentScreen = ""
    m_Game.CurrentSceneId = ""
    m_Game.CurrentScreenCode = ""
    m_Game.FrameDeltaSeconds = 0
    m_Game.LastFrameTimestamp = 0
End Sub

Private Function ClampLong(ByVal value As Long, ByVal minVal As Long, ByVal maxVal As Long, ByVal defaultVal As Long) As Long
    If value < minVal Then
        ClampLong = defaultVal
    ElseIf value > maxVal Then
        ClampLong = maxVal
    Else
        ClampLong = value
    End If
End Function

Private Function ClampDouble(ByVal value As Double, ByVal minVal As Double, ByVal maxVal As Double, ByVal defaultVal As Double) As Double
    If value < minVal Then
        ClampDouble = defaultVal
    ElseIf value > maxVal Then
        ClampDouble = maxVal
    Else
        ClampDouble = value
    End If
End Function

'===================================================================================
'                              SCREEN & SCENE PROPERTIES
'===================================================================================
Public Property Get CurrentScreen() As String
    CurrentScreen = m_Game.CurrentScreen
End Property

Public Property Let CurrentScreen(ByVal newScreen As String)
    m_Game.CurrentScreen = newScreen
End Property

Public Property Get CurrentScreenCode() As String
    CurrentScreenCode = m_Game.CurrentScreenCode
End Property

Public Property Let CurrentScreenCode(ByVal screenCode As String)
    m_Game.CurrentScreenCode = screenCode
End Property

Public Property Get CurrentSceneId() As String
    CurrentSceneId = m_Game.CurrentSceneId
End Property

Public Property Let CurrentSceneId(ByVal newScene As String)
    m_Game.CurrentSceneId = newScene
End Property

'===================================================================================
'                              TIMING & SPEED PROPERTIES
'===================================================================================
Public Property Get GameSpeed() As Long
    GameSpeed = m_Game.GameSpeed
End Property

Public Property Let MoveSpeed(ByVal newSpeed As Double)
    m_Game.MoveSpeed = ClampDouble(newSpeed, MIN_PLAYER_SPEED, MAX_PLAYER_SPEED, DEFAULT_PLAYER_SPEED)
End Property

Public Property Get MoveSpeed() As Double
    MoveSpeed = m_Game.MoveSpeed
End Property

'===================================================================================
'                              PLAYER STATE PROPERTIES
'===================================================================================
Public Property Get MoveDir() As String
    MoveDir = m_Game.MoveDir
End Property

Public Property Let MoveDir(ByVal direction As String)
    m_Game.MoveDir = direction
    If direction <> "" Then m_Game.LastDir = direction
End Property

Public Property Get LastDir() As String
    LastDir = m_Game.LastDir
End Property

Public Property Get IsPaused() As Boolean
    IsPaused = m_Game.IsPaused
End Property

Public Property Let IsPaused(ByVal paused As Boolean)
    m_Game.IsPaused = paused
End Property

Public Property Get IsFalling() As Boolean
    IsFalling = m_Game.IsFalling
End Property

Public Property Let IsFalling(ByVal falling As Boolean)
    m_Game.IsFalling = falling
End Property

Public Property Get PlayerCellAddress() As String
    PlayerCellAddress = m_Game.PlayerCellAddress
End Property

Public Property Let PlayerCellAddress(ByVal cellAddress As String)
    m_Game.PlayerCellAddress = cellAddress
End Property

'===================================================================================
'                              TRIGGER & EVENT PROPERTIES
'===================================================================================
Public Property Get CodeCell() As String
    CodeCell = m_Game.CodeCell
End Property

Public Property Let CodeCell(ByVal cellValue As String)
    m_Game.CodeCell = cellValue
End Property

Public Property Get TriggerCellAddress() As String
    TriggerCellAddress = m_Game.TriggerCellAddress
End Property

Public Property Let TriggerCellAddress(ByVal cellAddress As String)
    m_Game.TriggerCellAddress = cellAddress
End Property

'===================================================================================
'                              TIMER PROPERTIES
'===================================================================================
Public Property Get ScreenSetUpTimer() As Long
    ScreenSetUpTimer = m_Game.ScreenSetUpTimer
End Property

Public Property Let ScreenSetUpTimer(ByVal ticks As Long)
    m_Game.ScreenSetUpTimer = ticks
End Property

'===================================================================================
'                              FRAME TIMING
'===================================================================================
Public Function BeginFrame(ByVal fixedSeconds As Double) As Double
    ' Fixed timestep game loop - always uses the same delta time
    ' This ensures deterministic physics and consistent gameplay
    m_Game.FrameDeltaSeconds = fixedSeconds
    m_Game.LastFrameTimestamp = Timer
    BeginFrame = fixedSeconds
End Function

Public Property Get FrameDeltaSeconds() As Double
    FrameDeltaSeconds = m_Game.FrameDeltaSeconds
End Property

Public Property Get ExpectedFrameSeconds() As Double
    If m_Game.GameSpeed <= 0 Then
        ExpectedFrameSeconds = DEFAULT_FRAME_SECONDS
    Else
        ExpectedFrameSeconds = CDbl(m_Game.GameSpeed) / 1000#
    End If
End Property

'===================================================================================
'                              SHEET SYNCHRONISATION
'===================================================================================
Public Sub RefreshFromDataSheet()
    m_Game.GameSpeed = ReadGameSpeed()
    m_Game.MoveSpeed = ReadMoveSpeed()
End Sub

Private Function ReadNumericValue(ByVal rangeName As String, ByVal minValue As Double, ByVal maxValue As Double, ByVal defaultValue As Double) As Double
    ' Generic helper to read and clamp numeric values from data sheet
    On Error Resume Next
    Dim dataSheet As Worksheet
    Set dataSheet = GameRegistryInstance().GetGameDataSheet()
    Dim rawValue As Double
    rawValue = CDbl(dataSheet.Range(rangeName).Value)
    If Err.Number <> 0 Or rawValue < minValue Then
        rawValue = defaultValue
    ElseIf rawValue > maxValue Then
        rawValue = maxValue
    End If
    ReadNumericValue = rawValue
    On Error GoTo 0
End Function

Private Function ReadGameSpeed() As Long
    ReadGameSpeed = CLng(ReadNumericValue(RANGE_GAME_SPEED, MIN_GAME_SPEED, MAX_GAME_SPEED, DEFAULT_GAME_SPEED))
End Function

Private Function ReadMoveSpeed() As Double
    ReadMoveSpeed = ReadNumericValue(RANGE_PLAYER_MOVE, MIN_PLAYER_SPEED, MAX_PLAYER_SPEED, DEFAULT_PLAYER_SPEED)
End Function