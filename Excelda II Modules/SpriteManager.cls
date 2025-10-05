VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpriteManager"
Option Explicit

' Sprite references
Private Type SpriteState
    LinkSprite As Object
    SpriteFrame As Integer
    Position As Position
End Type

Private Type Position
    Top As Double
    Left As Double
End Type

Private m_Sprites As SpriteState
Private m_Instance As SpriteManager
Private m_Initialized As Boolean

' Singleton pattern
Private Sub Class_Initialize()
    If Not m_Initialized Then
        InitializeSprites
        m_Initialized = True
    End If
End Sub

Private Sub InitializeSprites()
    With m_Sprites
        Set .LinkSprite = ActiveSheet.Shapes("LinkDown2")
        .SpriteFrame = 2
        .Position.Top = .LinkSprite.Top
        .Position.Left = .LinkSprite.Left
    End With
End Sub

' Properties
Public Property Get LinkSprite() As Object
    Set LinkSprite = m_Sprites.LinkSprite
End Property

Public Property Get Top() As Double
    Top = m_Sprites.Position.Top
End Property

Public Property Let Top(ByVal Value As Double)
    m_Sprites.Position.Top = Value
End Property

Public Property Get Left() As Double
    Left = m_Sprites.Position.Left
End Property

Public Property Let Left(ByVal Value As Double)
    m_Sprites.Position.Left = Value
End Property

' Methods
Public Sub UpdateFrame(ByVal direction As String)
    m_Sprites.SpriteFrame = IIf(Sheets(SHEET_DATA).Range(RANGE_FRAME_COUNT).Value >= 5, 1, 2)
    
    Dim spriteName As String
    Select Case direction
        Case "U": spriteName = "LinkUp"
        Case "D": spriteName = "LinkDown"
        Case "L": spriteName = "LinkLeft"
        Case "R": spriteName = "LinkRight"
        Case Else: Exit Sub
    End Select
    
    Set m_Sprites.LinkSprite = ActiveSheet.Shapes(spriteName & m_Sprites.SpriteFrame)
End Sub

Public Sub UpdatePosition()
    With m_Sprites
        .LinkSprite.Top = .Position.Top
        .LinkSprite.Left = .Position.Left
    End With
End Sub

' Singleton accessor
Public Property Get Instance() As SpriteManager
    If m_Instance Is Nothing Then
        Set m_Instance = New SpriteManager
    End If
    Set Instance = m_Instance
End Property