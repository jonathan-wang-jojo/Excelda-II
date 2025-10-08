Option Explicit

'###################################################################################
'                              SPRITE MANAGEMENT
'###################################################################################
' Handles all sprite operations - positions, frames, visibility
'###################################################################################

Private Type SpriteState
    LinkSprite As Object
    LinkSpriteTop As Double
    LinkSpriteLeft As Double
    LinkSpriteFrame As Integer
End Type

Private m_Sprite As SpriteState
Private m_Initialized As Boolean

'------------------------------- Initialization -------------------------------
Private Sub Class_Initialize()
    If Not m_Initialized Then
        InitializeSprites
        m_Initialized = True
    End If
End Sub

Private Sub InitializeSprites()
    Dim wsGame As Worksheet
    Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    
    With m_Sprite
        Set .LinkSprite = wsGame.Shapes("LinkDown2")
        .LinkSpriteFrame = 2
        .LinkSpriteTop = .LinkSprite.Top
        .LinkSpriteLeft = .LinkSprite.Left
    End With
End Sub

'------------------------------- Reset / Destroy -------------------------------
Public Sub Reset()
    With m_Sprite
        .LinkSpriteTop = .LinkSprite.Top
        .LinkSpriteLeft = .LinkSprite.Left
        .LinkSpriteFrame = 2
    End With
    Call UpdateVisibility
End Sub

Public Sub Destroy()
    Dim emptySprite As SpriteState
    m_Sprite = emptySprite
End Sub

'------------------------------- Properties -------------------------------
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

'------------------------------- Sprite Logic -------------------------------
Public Sub UpdateFrame(ByVal direction As String, ByVal moveSpeed As Long)
    m_Sprite.LinkSpriteFrame = IIf(Sheets(SHEET_DATA).Range(RANGE_FRAME_COUNT).Value >= 5, 1, 2)
    
    If direction = "" Then Exit Sub
    
    Dim wsGame As Worksheet
    Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    
    Select Case direction
        Case "U":
            Set m_Sprite.LinkSprite = wsGame.Shapes("LinkUp" & m_Sprite.LinkSpriteFrame)
            m_Sprite.LinkSpriteTop = m_Sprite.LinkSpriteTop - moveSpeed
        Case "D":
            Set m_Sprite.LinkSprite = wsGame.Shapes("LinkDown" & m_Sprite.LinkSpriteFrame)
            m_Sprite.LinkSpriteTop = m_Sprite.LinkSpriteTop + moveSpeed
        Case "L":
            Set m_Sprite.LinkSprite = wsGame.Shapes("LinkLeft" & m_Sprite.LinkSpriteFrame)
            m_Sprite.LinkSpriteLeft = m_Sprite.LinkSpriteLeft - moveSpeed
        Case "R":
            Set m_Sprite.LinkSprite = wsGame.Shapes("LinkRight" & m_Sprite.LinkSpriteFrame)
            m_Sprite.LinkSpriteLeft = m_Sprite.LinkSpriteLeft + moveSpeed
    End Select
End Sub

Public Sub UpdatePosition()
    With m_Sprite
        .LinkSprite.Top = .LinkSpriteTop
        .LinkSprite.Left = .LinkSpriteLeft
    End With
End Sub

Public Sub UpdateVisibility()
    Dim wsGame As Worksheet
    Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    
    Dim directions As Variant: directions = Array("Up", "Down", "Left", "Right")
    Dim frames As Variant: frames = Array("1", "2")
    Dim dir As Variant, frame As Variant
    
    For Each dir In directions
        For Each frame In frames
            wsGame.Shapes("Link" & dir & frame).Visible = False
        Next frame
    Next dir
    
    m_Sprite.LinkSprite.Visible = True
End Sub

Public Sub AlignLinkSprites(ByVal leftPos As Double, ByVal topPos As Double)
    Dim wsGame As Worksheet
    Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    
    Dim s As Variant
    For Each s In Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", _
                        "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
        wsGame.Shapes(s).Left = leftPos
        wsGame.Shapes(s).Top = topPos
    Next s
End Sub