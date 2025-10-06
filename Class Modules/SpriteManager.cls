Option Explicit

'###################################################################################
'                              SPRITE MANAGEMENT
'###################################################################################
' Handles all sprite operations - positions, frames, visibility
'###################################################################################

' Sprite state
Private Type SpriteState
    LinkSprite As Object
    LinkSpriteTop As Double
    LinkSpriteLeft As Double
    LinkSpriteFrame As Integer
End Type

Private m_Sprite As SpriteState

' Singleton pattern
Private m_Instance As SpriteManager
Private m_Initialized As Boolean

Private Sub Class_Initialize()
    If Not m_Initialized Then
        InitializeSprites
        m_Initialized = True
    End If
End Sub

Private Sub InitializeSprites()
    With m_Sprite
        Set .LinkSprite = ActiveSheet.Shapes("LinkDown2")
        .LinkSpriteFrame = 2
        .LinkSpriteTop = .LinkSprite.Top
        .LinkSpriteLeft = .LinkSprite.Left
    End With
End Sub

' Sprite properties
Public Property Get LinkSprite() As Object
    Set LinkSprite = m_Sprite.LinkSprite
End Property

Public Property Set LinkSprite(ByVal Value As Object)
    Set m_Sprite.LinkSprite = Value
End Property

Public Property Get LinkSpriteTop() As Double
    LinkSpriteTop = m_Sprite.LinkSpriteTop
End Property

Public Property Let LinkSpriteTop(ByVal Value As Double)
    m_Sprite.LinkSpriteTop = Value
End Property

Public Property Get LinkSpriteLeft() As Double
    LinkSpriteLeft = m_Sprite.LinkSpriteLeft
End Property

Public Property Let LinkSpriteLeft(ByVal Value As Double)
    m_Sprite.LinkSpriteLeft = Value
End Property

Public Property Get LinkSpriteFrame() As Integer
    LinkSpriteFrame = m_Sprite.LinkSpriteFrame
End Property

Public Property Let LinkSpriteFrame(ByVal Value As Integer)
    m_Sprite.LinkSpriteFrame = Value
End Property

' Sprite methods
Public Sub UpdateFrame(ByVal direction As String, ByVal moveSpeed As Long)
    m_Sprite.LinkSpriteFrame = IIf(Sheets(SHEET_DATA).Range(RANGE_FRAME_COUNT).Value >= 5, 1, 2)
    
    ' Only update position if there's movement
    If direction <> "" Then
        Select Case direction
            Case "U":
                Set m_Sprite.LinkSprite = IIf(m_Sprite.LinkSpriteFrame = 1, ActiveSheet.Shapes("LinkUp1"), ActiveSheet.Shapes("LinkUp2"))
                m_Sprite.LinkSpriteTop = m_Sprite.LinkSpriteTop - moveSpeed
            Case "D":
                Set m_Sprite.LinkSprite = IIf(m_Sprite.LinkSpriteFrame = 1, ActiveSheet.Shapes("LinkDown1"), ActiveSheet.Shapes("LinkDown2"))
                m_Sprite.LinkSpriteTop = m_Sprite.LinkSpriteTop + moveSpeed
            Case "L":
                Set m_Sprite.LinkSprite = IIf(m_Sprite.LinkSpriteFrame = 1, ActiveSheet.Shapes("LinkLeft1"), ActiveSheet.Shapes("LinkLeft2"))
                m_Sprite.LinkSpriteLeft = m_Sprite.LinkSpriteLeft - moveSpeed
            Case "R":
                Set m_Sprite.LinkSprite = IIf(m_Sprite.LinkSpriteFrame = 1, ActiveSheet.Shapes("LinkRight1"), ActiveSheet.Shapes("LinkRight2"))
                m_Sprite.LinkSpriteLeft = m_Sprite.LinkSpriteLeft + moveSpeed
            Case "LU", "UL":
                Set m_Sprite.LinkSprite = IIf(m_Sprite.LinkSpriteFrame = 1, ActiveSheet.Shapes("LinkUp1"), ActiveSheet.Shapes("LinkUp2"))
                m_Sprite.LinkSpriteTop = m_Sprite.LinkSpriteTop - moveSpeed
                m_Sprite.LinkSpriteLeft = m_Sprite.LinkSpriteLeft - moveSpeed
            Case "RU", "UR":
                Set m_Sprite.LinkSprite = IIf(m_Sprite.LinkSpriteFrame = 1, ActiveSheet.Shapes("LinkUp1"), ActiveSheet.Shapes("LinkUp2"))
                m_Sprite.LinkSpriteTop = m_Sprite.LinkSpriteTop - moveSpeed
                m_Sprite.LinkSpriteLeft = m_Sprite.LinkSpriteLeft + moveSpeed
            Case "LD", "DL":
                Set m_Sprite.LinkSprite = IIf(m_Sprite.LinkSpriteFrame = 1, ActiveSheet.Shapes("LinkDown2"), ActiveSheet.Shapes("LinkDown1"))
                m_Sprite.LinkSpriteTop = m_Sprite.LinkSpriteTop + moveSpeed
                m_Sprite.LinkSpriteLeft = m_Sprite.LinkSpriteLeft - moveSpeed
            Case "RD", "DR":
                Set m_Sprite.LinkSprite = IIf(m_Sprite.LinkSpriteFrame = 1, ActiveSheet.Shapes("LinkDown1"), ActiveSheet.Shapes("LinkDown2"))
                m_Sprite.LinkSpriteTop = m_Sprite.LinkSpriteTop + moveSpeed
                m_Sprite.LinkSpriteLeft = m_Sprite.LinkSpriteLeft + moveSpeed
        End Select
    End If
End Sub

Public Sub UpdatePosition()
    With m_Sprite
        .LinkSprite.Top = .LinkSpriteTop
        .LinkSprite.Left = .LinkSpriteLeft
    End With
End Sub

Public Sub UpdateVisibility()
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
    m_Sprite.LinkSprite.Visible = True
End Sub

Public Sub AlignLinkSprites(ByVal leftPos As Double, ByVal topPos As Double)
    Dim s As Variant
    For Each s In Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
        ActiveSheet.Shapes(s).Left = leftPos
        ActiveSheet.Shapes(s).Top = topPos
    Next s
End Sub

' Singleton accessor
Public Property Get Instance() As SpriteManager
    If m_Instance Is Nothing Then
        Set m_Instance = New SpriteManager
    End If
    Set Instance = m_Instance
End Property