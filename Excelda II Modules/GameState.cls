VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameState"
Option Explicit

' Shared game state
Private Type State
    CurrentScreen As String
    GameSpeed As Long
    IsPaused As Boolean
End Type

Private m_State As State

' Cell location tracking
Private Type CellState
    LinkCellAddress As String
    CodeCell As String
End Type

Private m_CellState As CellState

' Movement state
Private Type MovementState
    MoveDir As String
    LastDir As String
    MoveSpeed As Long
    IsFalling As Boolean
End Type

Private m_Movement As MovementState

' Singleton pattern
Private m_Instance As GameState
Private m_Initialized As Boolean

Private Sub Class_Initialize()
    If Not m_Initialized Then
        InitializeState
        m_Initialized = True
    End If
End Sub

Private Sub InitializeState()
    With m_State
        .GameSpeed = Sheets(SHEET_DATA).Range(RANGE_GAME_SPEED).Value
        .IsPaused = False
    End With
    
    With m_Movement
        .MoveSpeed = Sheets(SHEET_DATA).Range(RANGE_LINK_MOVE).Value
        .MoveDir = ""
        .LastDir = "D" ' Default facing down
    End With
End Sub

' Properties
Public Property Get CurrentScreen() As String
    CurrentScreen = m_State.CurrentScreen
End Property

Public Property Let CurrentScreen(ByVal Value As String)
    m_State.CurrentScreen = Value
End Property

Public Property Get MoveDir() As String
    MoveDir = m_Movement.MoveDir
End Property

Public Property Let MoveDir(ByVal Value As String)
    m_Movement.MoveDir = Value
    If Value <> "" Then m_Movement.LastDir = Value
End Property

Public Property Get LastDir() As String
    LastDir = m_Movement.LastDir
End Property

Public Property Get GameSpeed() As Long
    GameSpeed = m_State.GameSpeed
End Property

Public Property Get MoveSpeed() As Long
    MoveSpeed = m_Movement.MoveSpeed
End Property

Public Property Get IsFalling() As Boolean
    IsFalling = m_Movement.IsFalling
End Property

Public Property Let IsFalling(ByVal Value As Boolean)
    m_Movement.IsFalling = Value
End Property

Public Property Get LinkCellAddress() As String
    LinkCellAddress = m_CellState.LinkCellAddress
End Property

Public Property Let LinkCellAddress(ByVal Value As String)
    m_CellState.LinkCellAddress = Value
End Property

Public Property Get CodeCell() As String
    CodeCell = m_CellState.CodeCell
End Property

Public Property Let CodeCell(ByVal Value As String)
    m_CellState.CodeCell = Value
End Property

' Singleton accessor
Public Property Get Instance() As GameState
    If m_Instance Is Nothing Then
        Set m_Instance = New GameState
    End If
    Set Instance = m_Instance
End Property