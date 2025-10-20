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
    StepAccumulator As Double
    LastFrameShape As String
    PendingFrameName As String
    PendingHideName As String
End Type

Private m_Sprite As SpriteData

' Initialization
Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Private Sub ResetState()
    Set m_Sprite.LinkSprite = Nothing
    m_Sprite.Top = 0
    m_Sprite.Left = 0
    m_Sprite.Frame = 1
    m_Sprite.StepAccumulator = 0
    m_Sprite.LastFrameShape = ""
    m_Sprite.PendingFrameName = ""
    m_Sprite.PendingHideName = ""
End Sub

Public Sub Initialize(Optional ByVal screenName As String = "", Optional ByVal spriteName As String = "")
    ResetState
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

    m_Sprite.LastFrameShape = m_Sprite.LinkSprite.Name
    m_Sprite.PendingFrameName = m_Sprite.LastFrameShape
    ApplyLinkVisibility m_Sprite.LinkSprite.Parent, m_Sprite.PendingFrameName, "", True
    m_Sprite.PendingFrameName = ""
    m_Sprite.PendingHideName = ""

    Exit Sub

ErrorHandler:
    ' Log error but don't crash - let caller handle it
    Debug.Print "SpriteManager.Initialize Error: " & Err.Description
    On Error GoTo 0
End Sub

' Core lifecycle
Public Sub Reset()
    On Error Resume Next
    ResetState
    
    Dim gs As GameState
    Set gs = GameStateInstance
    
    If Not gs Is Nothing And gs.CurrentScreen <> "" Then
        ' Try to reset to last known sprite or default to LinkDown2
        Set m_Sprite.LinkSprite = Sheets(gs.CurrentScreen).Shapes("LinkDown2")
        
        If Not m_Sprite.LinkSprite Is Nothing Then
            m_Sprite.Top = m_Sprite.LinkSprite.Top
            m_Sprite.Left = m_Sprite.LinkSprite.Left
            m_Sprite.Frame = 1
            m_Sprite.StepAccumulator = 0
            m_Sprite.LastFrameShape = m_Sprite.LinkSprite.Name
            m_Sprite.PendingFrameName = m_Sprite.LastFrameShape
            ApplyLinkVisibility m_Sprite.LinkSprite.Parent, m_Sprite.PendingFrameName, "", True
            m_Sprite.PendingFrameName = ""
            m_Sprite.PendingHideName = ""
        End If
    End If
    
    On Error GoTo 0
End Sub

Public Sub Destroy()
    ResetState
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
Public Sub UpdateFrame(ByVal movementDir As String, ByVal facingDir As String, ByVal moveSpeed As Double, ByVal deltaSeconds As Double)
    On Error Resume Next
    Dim gs As GameState
    Set gs = GameStateInstance

    Dim expectedFrameSeconds As Double
    If Not gs Is Nothing Then
        expectedFrameSeconds = gs.ExpectedFrameSeconds
    End If
    If expectedFrameSeconds <= 0# Then expectedFrameSeconds = DEFAULT_FRAME_SECONDS

    Dim effectiveDelta As Double
    effectiveDelta = deltaSeconds
    If effectiveDelta <= 0# Then effectiveDelta = expectedFrameSeconds

    Dim frameScale As Double
    frameScale = effectiveDelta / expectedFrameSeconds
    If frameScale < 0.25 Then frameScale = 0.25
    If frameScale > 2# Then frameScale = 2#

    Dim distance As Double
    distance = moveSpeed * frameScale

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

    Dim oldTop As Double
    Dim oldLeft As Double
    oldTop = m_Sprite.Top
    oldLeft = m_Sprite.Left

    Dim newTop As Double
    Dim newLeft As Double
    newTop = oldTop
    newLeft = oldLeft

    If moveUp Then newTop = newTop - verticalStep
    If moveDown Then newTop = newTop + verticalStep
    If moveLeft Then newLeft = newLeft - horizontalStep
    If moveRight Then newLeft = newLeft + horizontalStep

    Dim moveDelta As Double
    moveDelta = Sqr((newLeft - oldLeft) ^ 2 + (newTop - oldTop) ^ 2)

    If Len(movementDir) = 0 Or moveDelta <= 0.0001 Then
        If m_Sprite.Frame <> 1 Then m_Sprite.Frame = 1
        m_Sprite.StepAccumulator = 0#
    Else
        Dim thresholdDistance As Double
        If LINK_STEP_FRAME_DISTANCE > 0# Then
            thresholdDistance = LINK_STEP_FRAME_DISTANCE
        Else
            thresholdDistance = moveSpeed * LINK_ANIM_FRAME_SECONDS
            If thresholdDistance <= 0# Then thresholdDistance = 6#
        End If
        m_Sprite.StepAccumulator = m_Sprite.StepAccumulator + moveDelta
        If m_Sprite.StepAccumulator >= thresholdDistance Then
            m_Sprite.StepAccumulator = m_Sprite.StepAccumulator - thresholdDistance
            If m_Sprite.StepAccumulator > thresholdDistance Then
                m_Sprite.StepAccumulator = thresholdDistance
            End If
            If m_Sprite.Frame = 1 Then
                m_Sprite.Frame = 2
            Else
                m_Sprite.Frame = 1
            End If
        End If
    End If

    m_Sprite.Top = newTop
    m_Sprite.Left = newLeft

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

    Dim activeShape As Shape
    Set activeShape = ShapeIfExists(ws, shapeName)
    If activeShape Is Nothing Then Exit Sub

    Set m_Sprite.LinkSprite = activeShape
    m_Sprite.PendingFrameName = activeShape.Name

    ' Sync global for legacy code compatibility
    Set LinkSprite = m_Sprite.LinkSprite

    On Error GoTo 0
End Sub

Public Sub UpdatePosition()
    On Error Resume Next
    If m_Sprite.LinkSprite Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    m_Sprite.LinkSprite.Top = m_Sprite.Top
    m_Sprite.LinkSprite.Left = m_Sprite.Left
    SyncAllLinkFrames
    On Error GoTo 0
End Sub

Public Sub UpdateVisibility()
    On Error Resume Next
    Dim gs As GameState
    Set gs = GameStateInstance
    Dim wsGame As Worksheet
    If Not gs Is Nothing Then
        If gs.CurrentScreen <> "" Then
            Set wsGame = Sheets(gs.CurrentScreen)
        End If
    End If
    If wsGame Is Nothing And Not m_Sprite.LinkSprite Is Nothing Then
        Set wsGame = m_Sprite.LinkSprite.Parent
    End If
    If wsGame Is Nothing Then Exit Sub

    Dim hideShape As Shape
    If m_Sprite.PendingHideName <> "" Then
        Set hideShape = ShapeIfExists(wsGame, m_Sprite.PendingHideName)
        If Not hideShape Is Nothing Then
            If hideShape.Visible <> msoFalse Then hideShape.Visible = msoFalse
        End If
        m_Sprite.PendingHideName = ""
    End If

    Dim targetName As String
    Dim previousName As String
    Dim needsFullRefresh As Boolean

    targetName = m_Sprite.PendingFrameName

    If targetName = "" Then
        If m_Sprite.LastFrameShape = "" Then
            If m_Sprite.LinkSprite Is Nothing Then Exit Sub
            targetName = m_Sprite.LinkSprite.Name
            needsFullRefresh = True
        Else
            Exit Sub
        End If
    Else
        If m_Sprite.LastFrameShape = "" Then
            needsFullRefresh = True
        ElseIf m_Sprite.LastFrameShape <> targetName Then
            previousName = m_Sprite.LastFrameShape
        Else
            m_Sprite.PendingFrameName = ""
            Exit Sub
        End If
    End If

    ApplyLinkVisibility wsGame, targetName, "", needsFullRefresh

    m_Sprite.LastFrameShape = targetName
    m_Sprite.PendingFrameName = ""
    If needsFullRefresh Then
        m_Sprite.PendingHideName = ""
    ElseIf previousName <> "" Then
        m_Sprite.PendingHideName = previousName
    End If

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
    
    Dim gs As GameState
    Set gs = GameStateInstance
    Dim wsGame As Worksheet
    If Not gs Is Nothing Then
        If gs.CurrentScreen <> "" Then
            Set wsGame = Sheets(gs.CurrentScreen)
        End If
    End If
    If wsGame Is Nothing And Not m_Sprite.LinkSprite Is Nothing Then
        Set wsGame = m_Sprite.LinkSprite.Parent
    End If
    If wsGame Is Nothing Then Exit Sub
    
    Dim frameName As Variant
    For Each frameName In LinkFrameNames()
        Dim frameShape As Shape
        Set frameShape = ShapeIfExists(wsGame, CStr(frameName))
        If Not frameShape Is Nothing Then
            frameShape.Left = leftPos
            frameShape.Top = topPos
        End If
    Next frameName
    
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
    Dim gs As GameState
    Set gs = GameStateInstance
    If Not gs Is Nothing Then
        If gs.CurrentScreen <> "" Then
            Set wsGame = Sheets(gs.CurrentScreen)
        End If
    End If
    If wsGame Is Nothing And Not m_Sprite.LinkSprite Is Nothing Then
        Set wsGame = m_Sprite.LinkSprite.Parent
    End If
    If wsGame Is Nothing Then Exit Sub
    Dim frameName As Variant
    For Each frameName In LinkFrameNames()
        Dim frameShape As Shape
        Set frameShape = ShapeIfExists(wsGame, CStr(frameName))
        If Not frameShape Is Nothing Then
            frameShape.Top = m_Sprite.Top
            frameShape.Left = m_Sprite.Left
        End If
    Next frameName
    On Error GoTo 0
End Sub

Private Function LinkFrameNames() As Variant
    LinkFrameNames = Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", _
                           "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
End Function

Private Function ShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    On Error Resume Next
    Set ShapeIfExists = ws.Shapes(shapeName)
    If Err.Number <> 0 Then
        Err.Clear
        Set ShapeIfExists = Nothing
    End If
    On Error GoTo 0
End Function

Private Sub ApplyLinkVisibility(ByVal ws As Worksheet, ByVal activeName As String, Optional ByVal previousName As String = "", Optional ByVal fullRefresh As Boolean = False)
    If activeName = "" Then Exit Sub

    Dim activeShape As Shape
    Set activeShape = ShapeIfExists(ws, activeName)
    If Not activeShape Is Nothing Then
        If activeShape.Visible <> msoTrue Then activeShape.Visible = msoTrue
    End If

    Dim frameShape As Shape
    If previousName <> "" And previousName <> activeName Then
        Set frameShape = ShapeIfExists(ws, previousName)
        If Not frameShape Is Nothing Then
            If frameShape.Visible <> msoFalse Then frameShape.Visible = msoFalse
        End If
    End If

    If fullRefresh Then
        Dim frameName As Variant
        For Each frameName In LinkFrameNames()
            If CStr(frameName) <> activeName And CStr(frameName) <> previousName Then
                Set frameShape = ShapeIfExists(ws, CStr(frameName))
                If Not frameShape Is Nothing Then
                    If frameShape.Visible <> msoFalse Then frameShape.Visible = msoFalse
                End If
            End If
        Next frameName
    End If
End Sub