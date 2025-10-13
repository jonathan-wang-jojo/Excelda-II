Option Explicit

'###################################################################################
'                              CORE GAME STATE
'###################################################################################

Private Type GameData
    CurrentScreen As String
    GameSpeed As Long
    IsPaused As Boolean
    MoveDir As String
    LastDir As String
    MoveSpeed As Long
    IsFalling As Boolean
    LinkCellAddress As String
    CodeCell As String
    RNDBounceback As String
    CollidedWith As String
    ScreenSetUpTimer As Long
End Type

Private m_Game As GameData

' Initialization
Private Sub Class_Initialize()
    On Error Resume Next
    With m_Game
        .GameSpeed = ReadGameSpeed()
        .MoveSpeed = Sheets(SHEET_DATA).Range(RANGE_LINK_MOVE).Value
        .IsPaused = False
        .MoveDir = ""
        .LastDir = "D"
        .IsFalling = False
        .LinkCellAddress = ""
        .CodeCell = ""
        .RNDBounceback = ""
        .CollidedWith = ""
        .ScreenSetUpTimer = 0
    End With
    On Error GoTo 0
End Sub

' Core lifecycle
Public Sub Reset()
    m_Game.GameSpeed = ReadGameSpeed()
    m_Game.MoveDir = ""
    m_Game.IsFalling = False
    m_Game.RNDBounceback = ""
    m_Game.CollidedWith = ""
    m_Game.ScreenSetUpTimer = 0
End Sub

Public Sub Destroy()
    Reset
End Sub

' Essential properties
Public Property Get CurrentScreen() As String
    CurrentScreen = m_Game.CurrentScreen
End Property

Public Property Let CurrentScreen(ByVal Value As String)
    m_Game.CurrentScreen = Value
End Property

Public Property Get GameSpeed() As Long
    GameSpeed = m_Game.GameSpeed
End Property

Public Property Let GameSpeed(ByVal Value As Long)
    If Value <= 0 Then
        m_Game.GameSpeed = DEFAULT_GAME_SPEED
    Else
        m_Game.GameSpeed = Value
    End If
End Property

Public Property Get IsPaused() As Boolean
    IsPaused = m_Game.IsPaused
End Property

Public Property Let IsPaused(ByVal Value As Boolean)
    m_Game.IsPaused = Value
End Property

Public Property Get MoveDir() As String
    MoveDir = m_Game.MoveDir
End Property

Public Property Let MoveDir(ByVal Value As String)
    m_Game.MoveDir = Value
    If Value <> "" Then m_Game.LastDir = Value
End Property

Public Property Get LastDir() As String
    LastDir = m_Game.LastDir
End Property

Public Property Get MoveSpeed() As Long
    MoveSpeed = m_Game.MoveSpeed
End Property

Public Property Get IsFalling() As Boolean
    IsFalling = m_Game.IsFalling
End Property

Public Property Let IsFalling(ByVal Value As Boolean)
    m_Game.IsFalling = Value
End Property

Public Property Get LinkCellAddress() As String
    LinkCellAddress = m_Game.LinkCellAddress
End Property

Public Property Let LinkCellAddress(ByVal Value As String)
    m_Game.LinkCellAddress = Value
End Property

Public Property Get CodeCell() As String
    CodeCell = m_Game.CodeCell
End Property

Public Property Let CodeCell(ByVal Value As String)
    m_Game.CodeCell = Value
End Property

Public Property Get RNDBounceback() As String
    RNDBounceback = m_Game.RNDBounceback
End Property

Public Property Let RNDBounceback(ByVal Value As String)
    m_Game.RNDBounceback = Value
End Property

Public Property Get CollidedWith() As String
    CollidedWith = m_Game.CollidedWith
End Property

Public Property Let CollidedWith(ByVal Value As String)
    m_Game.CollidedWith = Value
End Property

Public Property Get ScreenSetUpTimer() As Long
    ScreenSetUpTimer = m_Game.ScreenSetUpTimer
End Property

Public Property Let ScreenSetUpTimer(ByVal Value As Long)
    m_Game.ScreenSetUpTimer = Value
End Property

Private Function ReadGameSpeed() As Long
    On Error Resume Next
    Dim value As Long
    value = CLng(Val(Sheets(SHEET_DATA).Range(RANGE_GAME_SPEED).Value))
    If value <= 0 Then value = DEFAULT_GAME_SPEED
    ReadGameSpeed = value
    On Error GoTo 0
End Function