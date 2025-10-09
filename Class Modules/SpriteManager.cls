Option Explicit

'###################################################################################
'                              SPRITE MANAGEMENT
'###################################################################################
' Handles all sprite operations - positions, frames, visibility
'###################################################################################

Private Type SpriteData
    LinkSprite As Shape
    Top As Double
    Left As Double
    Frame As Integer
End Type

Private m_Sprite As SpriteData

' Initialization
Private Sub Class_Initialize()
    m_Sprite.Frame = 2
End Sub

Public Sub Initialize(ByVal screenName As String, ByVal spriteName As String)
    On Error GoTo ErrorHandler
    
    ' Find and set the sprite
    Set m_Sprite.LinkSprite = Sheets(screenName).Shapes(spriteName)
    
    If m_Sprite.LinkSprite Is Nothing Then
        Err.Raise vbObjectError + 1, "SpriteManager.Initialize", "Sprite '" & spriteName & "' not found on " & screenName
    End If
    
    ' Set initial position from sprite's actual location
    m_Sprite.Top = m_Sprite.LinkSprite.Top
    m_Sprite.Left = m_Sprite.LinkSprite.Left
    
    ' Determine frame from sprite name (LinkDown1 = frame 1, LinkDown2 = frame 2, etc.)
    If InStr(spriteName, "1") > 0 Then
        m_Sprite.Frame = 1
    ElseIf InStr(spriteName, "2") > 0 Then
        m_Sprite.Frame = 2
    Else
        m_Sprite.Frame = 2  ' Default to frame 2
    End If
    
    Exit Sub

ErrorHandler:
    ' Log error but don't crash - let caller handle it
    Debug.Print "SpriteManager.Initialize Error: " & Err.Description
    On Error GoTo 0
End Sub

' Core lifecycle
Public Sub Reset()
    On Error Resume Next
    
    Dim gs As GameState
    Set gs = GameStateInstance
    
    If Not gs Is Nothing And gs.CurrentScreen <> "" Then
        ' Try to reset to last known sprite or default to LinkDown2
        Set m_Sprite.LinkSprite = Sheets(gs.CurrentScreen).Shapes("LinkDown2")
        
        If Not m_Sprite.LinkSprite Is Nothing Then
            m_Sprite.Top = m_Sprite.LinkSprite.Top
            m_Sprite.Left = m_Sprite.LinkSprite.Left
            m_Sprite.Frame = 2
        End If
    End If
    
    On Error GoTo 0
End Sub

Public Sub Destroy()
    Set m_Sprite.LinkSprite = Nothing
    m_Sprite.Top = 0
    m_Sprite.Left = 0
    m_Sprite.Frame = 0
End Sub


' Essential properties
Public Property Get LinkSprite() As Object
    Set LinkSprite = m_Sprite.LinkSprite
End Property

Public Property Get LinkSpriteTop() As Double
    LinkSpriteTop = m_Sprite.Top
End Property

Public Property Let LinkSpriteTop(ByVal Value As Double)
    m_Sprite.Top = Value
End Property

Public Property Get LinkSpriteLeft() As Double
    LinkSpriteLeft = m_Sprite.Left
End Property

Public Property Let LinkSpriteLeft(ByVal Value As Double)
    m_Sprite.Left = Value
End Property

Public Property Get LinkSpriteFrame() As Integer
    LinkSpriteFrame = m_Sprite.Frame
End Property

Public Property Let LinkSpriteFrame(ByVal Value As Integer)
    m_Sprite.Frame = Value
End Property

' Core sprite operations - delegates to optimized logic
Public Sub UpdateFrame(ByVal direction As String, ByVal moveSpeed As Long)
    On Error Resume Next
    
    m_Sprite.Frame = IIf(Sheets(SHEET_DATA).Range(RANGE_FRAME_COUNT).Value >= 5, 1, 2)
    
    If direction <> "" Then
        Set m_Sprite.LinkSprite = Sheets(GameStateInstance.CurrentScreen).Shapes("Link" & _
            IIf(direction = "U", "Up", IIf(direction = "D", "Down", IIf(direction = "L", "Left", "Right"))) & _
            m_Sprite.Frame)
        
        ' Sync global for legacy code compatibility
        Set LinkSprite = m_Sprite.LinkSprite
        
        Dim newTop As Double, newLeft As Double
        newTop = m_Sprite.Top
        newLeft = m_Sprite.Left

        ' Adjust position based on direction components
        If InStr(direction, "U") > 0 Then newTop = newTop - moveSpeed
        If InStr(direction, "D") > 0 Then newTop = newTop + moveSpeed
        If InStr(direction, "L") > 0 Then newLeft = newLeft - moveSpeed
        If InStr(direction, "R") > 0 Then newLeft = newLeft + moveSpeed

        ' Update sprite position
        m_Sprite.Top = newTop
        m_Sprite.Left = newLeft
    End If
    
    On Error GoTo 0
End Sub

Public Sub UpdatePosition()
    On Error Resume Next
    m_Sprite.LinkSprite.Top = m_Sprite.Top
    m_Sprite.LinkSprite.Left = m_Sprite.Left
    On Error GoTo 0
End Sub

Public Sub UpdateVisibility()
    On Error Resume Next
    
    Dim wsGame As Worksheet: Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    Dim s As Variant
    
    For Each s In Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", _
                        "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
        wsGame.Shapes(s).Visible = False
    Next s
    
    m_Sprite.LinkSprite.Visible = True
    
    On Error GoTo 0
End Sub

Public Sub AlignSprites(ByVal leftPos As Double, ByVal topPos As Double)
    On Error Resume Next
    
    m_Sprite.Left = leftPos
    m_Sprite.Top = topPos
    
    Dim wsGame As Worksheet: Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    Dim s As Variant
    
    For Each s In Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", _
                        "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
        wsGame.Shapes(s).Left = leftPos
        wsGame.Shapes(s).Top = topPos
    Next s
    
    On Error GoTo 0
End Sub