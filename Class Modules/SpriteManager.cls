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
    FrameCounter As Long
End Type

Private m_Sprite As SpriteData

' Initialization
Private Sub Class_Initialize()
    InitializeState
End Sub

Private Sub InitializeState()
    Set m_Sprite.LinkSprite = Nothing
    m_Sprite.Top = 0
    m_Sprite.Left = 0
    m_Sprite.Frame = 1
    m_Sprite.FrameCounter = 0
End Sub

Public Sub Initialize(Optional ByVal screenName As String = "", Optional ByVal spriteName As String = "")
    InitializeState
    If screenName = "" Or spriteName = "" Then Exit Sub

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
    InitializeState
    
    Dim gs As GameState
    Set gs = GameStateInstance
    
    If Not gs Is Nothing And gs.CurrentScreen <> "" Then
        ' Try to reset to last known sprite or default to LinkDown2
        Set m_Sprite.LinkSprite = Sheets(gs.CurrentScreen).Shapes("LinkDown2")
        
        If Not m_Sprite.LinkSprite Is Nothing Then
            m_Sprite.Top = m_Sprite.LinkSprite.Top
            m_Sprite.Left = m_Sprite.LinkSprite.Left
            m_Sprite.Frame = 1
            m_Sprite.FrameCounter = 0
        End If
    End If
    
    On Error GoTo 0
End Sub

Public Sub Destroy()
    Set m_Sprite.LinkSprite = Nothing
    m_Sprite.Top = 0
    m_Sprite.Left = 0
    m_Sprite.Frame = 0
    m_Sprite.FrameCounter = 0
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
Public Sub UpdateFrame(ByVal movementDir As String, ByVal facingDir As String, ByVal moveSpeed As Long, ByVal deltaSeconds As Double)
    On Error Resume Next
    Dim gs As GameState
    Set gs = GameStateInstance

    m_Sprite.FrameCounter = (m_Sprite.FrameCounter + 1) Mod 6
    m_Sprite.Frame = IIf(m_Sprite.FrameCounter < 3, 1, 2)
    
    If facingDir = "" Then
        If Not gs Is Nothing Then
            facingDir = gs.LastDir
        Else
            facingDir = ""
        End If
    End If
    Dim baseDir As String
    baseDir = ResolveDirection(facingDir)
    
    Dim ws As Worksheet
    If Not gs Is Nothing Then
        If gs.CurrentScreen <> "" Then
            Set ws = Sheets(gs.CurrentScreen)
        End If
    End If
    If ws Is Nothing Then
        If Not m_Sprite.LinkSprite Is Nothing Then
            Set ws = m_Sprite.LinkSprite.Parent
        End If
    End If
    If ws Is Nothing Then Exit Sub
    Dim shapeName As String
    shapeName = "Link" & baseDir & m_Sprite.Frame
    Set m_Sprite.LinkSprite = ws.Shapes(shapeName)
    If m_Sprite.LinkSprite Is Nothing Then Exit Sub
    
    ' Sync global for legacy code compatibility
    Set LinkSprite = m_Sprite.LinkSprite
    
    Dim expectedFrameSeconds As Double
    If Not gs Is Nothing Then
        expectedFrameSeconds = gs.ExpectedFrameSeconds
    End If
    If expectedFrameSeconds <= 0# Then expectedFrameSeconds = DEFAULT_FRAME_SECONDS
    If deltaSeconds <= 0# Then deltaSeconds = expectedFrameSeconds
    Dim frameScale As Double
    frameScale = deltaSeconds / expectedFrameSeconds
    If frameScale < 0.1 Then frameScale = 0.1
    If frameScale > 3# Then frameScale = 3#
    
    Dim distance As Double
    distance = CDbl(moveSpeed) * frameScale
    
    Dim moveUp As Boolean, moveDown As Boolean, moveLeft As Boolean, moveRight As Boolean
    moveUp = InStr(movementDir, "U") > 0
    moveDown = InStr(movementDir, "D") > 0
    If moveUp And moveDown Then
        moveUp = False
        moveDown = False
    End If
    moveLeft = InStr(movementDir, "L") > 0
    moveRight = InStr(movementDir, "R") > 0
    If moveLeft And moveRight Then
        moveLeft = False
        moveRight = False
    End If
    
    Dim hasVertical As Boolean
    Dim hasHorizontal As Boolean
    hasVertical = moveUp Or moveDown
    hasHorizontal = moveLeft Or moveRight
    
    Dim verticalStep As Double
    Dim horizontalStep As Double
    If hasVertical And hasHorizontal Then
        verticalStep = distance / Sqr(2#)
        horizontalStep = distance / Sqr(2#)
    Else
        verticalStep = distance
        horizontalStep = distance
    End If

    Dim newTop As Double, newLeft As Double
    newTop = m_Sprite.Top
    newLeft = m_Sprite.Left

    If moveUp Then newTop = newTop - verticalStep
    If moveDown Then newTop = newTop + verticalStep
    If moveLeft Then newLeft = newLeft - horizontalStep
    If moveRight Then newLeft = newLeft + horizontalStep

    m_Sprite.Top = newTop
    m_Sprite.Left = newLeft
    
    On Error GoTo 0
End Sub

Public Sub UpdatePosition()
    On Error Resume Next
    m_Sprite.LinkSprite.Top = m_Sprite.Top
    m_Sprite.LinkSprite.Left = m_Sprite.Left
    SyncAllLinkFrames
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

Public Sub ApplyLinkBounce(ByVal bounceSpeed As Long)
    On Error Resume Next
    If bounceSpeed <= 0 Then Exit Sub
    
    Dim linkShape As Shape
    Set linkShape = m_Sprite.LinkSprite
    If linkShape Is Nothing Then Exit Sub
    
    Dim originCell As Range
    Set originCell = linkShape.TopLeftCell
    If originCell Is Nothing Then Exit Sub
    
    Dim blockingValue As Variant
    
    Select Case linkShape.Name
        Case "LinkDown1", "LinkDown2"
            blockingValue = originCell.Offset(-1, 2).Value
            If blockingValue = "" Then
                linkShape.Top = linkShape.Top - bounceSpeed
            End If
        Case "LinkUp1", "LinkUp2"
            blockingValue = originCell.Offset(4, 2).Value
            If blockingValue = "" Then
                linkShape.Top = linkShape.Top + bounceSpeed
            End If
        Case "LinkLeft1", "LinkLeft2"
            blockingValue = originCell.Offset(2, 4).Value
            If blockingValue = "" Then
                linkShape.Left = linkShape.Left + bounceSpeed
            End If
        Case "LinkRight1", "LinkRight2"
            blockingValue = originCell.Offset(2, -1).Value
            If blockingValue = "" Then
                linkShape.Left = linkShape.Left - bounceSpeed
            End If
    End Select
    
    m_Sprite.Top = linkShape.Top
    m_Sprite.Left = linkShape.Left
    SyncAllLinkFrames
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

Private Function ResolveDirection(ByVal direction As String) As String
    If InStr(direction, "U") > 0 Then
        ResolveDirection = "Up"
    ElseIf InStr(direction, "D") > 0 Then
        ResolveDirection = "Down"
    ElseIf InStr(direction, "L") > 0 Then
        ResolveDirection = "Left"
    ElseIf InStr(direction, "R") > 0 Then
        ResolveDirection = "Right"
    Else
        ResolveDirection = "Down"
    End If
End Function

Private Sub SyncAllLinkFrames()
    Dim wsGame As Worksheet
    On Error Resume Next
    Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    If wsGame Is Nothing Then Exit Sub
    Dim s As Variant
    For Each s In Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", _
                        "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
        wsGame.Shapes(s).Top = m_Sprite.Top
        wsGame.Shapes(s).Left = m_Sprite.Left
    Next s
    On Error GoTo 0
End Sub