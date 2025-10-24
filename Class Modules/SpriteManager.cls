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
    LastFrameShape As String
    PendingFrameName As String
    PendingHideName As String
    PrevTop As Double
    PrevLeft As Double
    AnimTickCounter As Long
End Type

Private m_Sprite As SpriteData

'===================================================================================
'                              INITIALIZATION
'===================================================================================
Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

'===================================================================================
'                              INTERNAL STATE
'===================================================================================
Private Sub ResetState()
    Set m_Sprite.LinkSprite = Nothing
    m_Sprite.Top = 0
    m_Sprite.Left = 0
    m_Sprite.Frame = 1
    m_Sprite.LastFrameShape = ""
    m_Sprite.PendingFrameName = ""
    m_Sprite.PendingHideName = ""
    m_Sprite.PrevTop = 0
    m_Sprite.PrevLeft = 0
    m_Sprite.AnimTickCounter = 0
End Sub

'===================================================================================
'                              SETUP & LIFECYCLE
'===================================================================================
Public Sub Initialize()
    ResetState
End Sub

Public Sub BindLinkSprite(ByVal screenName As String, ByVal spriteName As String)
    ResetState
    If screenName = "" Or spriteName = "" Then Exit Sub

    On Error GoTo ErrorHandler

    Set m_Sprite.LinkSprite = Sheets(screenName).Shapes(spriteName)

    If m_Sprite.LinkSprite Is Nothing Then
        Err.Raise vbObjectError + 1, "SpriteManager.BindLinkSprite", "Sprite '" & spriteName & "' not found on " & screenName
    End If

    m_Sprite.Top = m_Sprite.LinkSprite.Top
    m_Sprite.Left = m_Sprite.LinkSprite.Left
    m_Sprite.PrevTop = m_Sprite.Top
    m_Sprite.PrevLeft = m_Sprite.Left

    If InStr(spriteName, "1") > 0 Then
        m_Sprite.Frame = 1
    ElseIf InStr(spriteName, "2") > 0 Then
        m_Sprite.Frame = 2
    Else
        m_Sprite.Frame = 2
    End If

    m_Sprite.LastFrameShape = m_Sprite.LinkSprite.Name
    m_Sprite.PendingFrameName = m_Sprite.LastFrameShape
    ApplyLinkVisibility m_Sprite.LinkSprite.Parent, m_Sprite.PendingFrameName, "", True
    m_Sprite.PendingFrameName = ""
    m_Sprite.PendingHideName = ""
    m_Sprite.AnimTickCounter = 0

    Exit Sub

ErrorHandler:
    Debug.Print "SpriteManager.BindLinkSprite Error: " & Err.Description
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
            m_Sprite.LastFrameShape = m_Sprite.LinkSprite.Name
            m_Sprite.PendingFrameName = m_Sprite.LastFrameShape
            ApplyLinkVisibility m_Sprite.LinkSprite.Parent, m_Sprite.PendingFrameName, "", True
            m_Sprite.PendingFrameName = ""
            m_Sprite.PendingHideName = ""
            m_Sprite.PrevTop = m_Sprite.Top
            m_Sprite.PrevLeft = m_Sprite.Left
            m_Sprite.AnimTickCounter = 0
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

Public Property Let LinkSpriteTop(ByVal positionTop As Double)
    m_Sprite.Top = positionTop
End Property

Public Property Get LinkSpriteLeft() As Double
    LinkSpriteLeft = m_Sprite.Left
End Property

Public Property Let LinkSpriteLeft(ByVal positionLeft As Double)
    m_Sprite.Left = positionLeft
End Property

Public Property Get LinkSpriteFrame() As Integer
    LinkSpriteFrame = m_Sprite.Frame
End Property

Public Property Let LinkSpriteFrame(ByVal frameIndex As Integer)
    m_Sprite.Frame = frameIndex
End Property

'===================================================================================
'                              CORE SPRITE OPERATIONS
'===================================================================================
' movementDir/facingDir capture intent; deltaSeconds maintained for API compatibility.
Public Sub UpdateFrame(ByVal movementDir As String, ByVal facingDir As String, ByVal moveSpeed As Double, ByVal deltaSeconds As Double)
    On Error Resume Next
    Dim gs As GameState
    Set gs = GameStateInstance

    Dim pixelsPerTick As Double
    pixelsPerTick = moveSpeed * LINK_SPEED_MULTIPLIER
    If pixelsPerTick <= 0# Then
        pixelsPerTick = 0#
    ElseIf pixelsPerTick < MIN_PIXELS_PER_TICK Then
        pixelsPerTick = MIN_PIXELS_PER_TICK
    End If

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
        If pixelsPerTick > 0# Then
            verticalStep = pixelsPerTick / Sqr(2#)
            horizontalStep = verticalStep
        Else
            verticalStep = 0#
            horizontalStep = 0#
        End If
    Else
        verticalStep = pixelsPerTick
        horizontalStep = pixelsPerTick
    End If

    Dim oldTop As Double
    Dim oldLeft As Double
    oldTop = m_Sprite.Top
    oldLeft = m_Sprite.Left
    m_Sprite.PrevTop = oldTop
    m_Sprite.PrevLeft = oldLeft

    Dim newTop As Double
    Dim newLeft As Double
    newTop = oldTop
    newLeft = oldLeft

    If moveUp Then newTop = newTop - verticalStep
    If moveDown Then newTop = newTop + verticalStep
    If moveLeft Then newLeft = newLeft - horizontalStep
    If moveRight Then newLeft = newLeft + horizontalStep

    Dim movedThisTick As Boolean
    movedThisTick = (newLeft <> oldLeft) Or (newTop <> oldTop)

    If Not movedThisTick Then
        If m_Sprite.Frame <> 1 Then m_Sprite.Frame = 1
        m_Sprite.AnimTickCounter = 0
    ElseIf ANIMATION_TICKS_PER_FRAME > 0 Then
        m_Sprite.AnimTickCounter = m_Sprite.AnimTickCounter + 1
        If m_Sprite.AnimTickCounter >= ANIMATION_TICKS_PER_FRAME Then
            m_Sprite.AnimTickCounter = 0
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

    Set LinkSprite = m_Sprite.LinkSprite

    On Error GoTo 0
End Sub

'===================================================================================
'                              RENDERING HELPERS
'===================================================================================
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

Public Sub RenderInterpolated(ByVal alpha As Double)
    ' Render sprite at interpolated position between Prev and current
    On Error Resume Next
    If m_Sprite.LinkSprite Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    If alpha < 0# Then alpha = 0#
    If alpha > 1# Then alpha = 1#

    Dim interpTop As Double
    Dim interpLeft As Double
    interpTop = m_Sprite.PrevTop + (m_Sprite.Top - m_Sprite.PrevTop) * alpha
    interpLeft = m_Sprite.PrevLeft + (m_Sprite.Left - m_Sprite.PrevLeft) * alpha

    m_Sprite.LinkSprite.Top = interpTop
    m_Sprite.LinkSprite.Left = interpLeft
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
    If wsGame Is Nothing Then GoTo Cleanup

    Dim previousScreenUpdating As Boolean
    Dim restoreScreenUpdating As Boolean
    previousScreenUpdating = Application.ScreenUpdating
    If previousScreenUpdating Then
        Application.ScreenUpdating = False
        restoreScreenUpdating = True
    End If

    Dim targetName As String
    Dim previousName As String
    Dim needsFullRefresh As Boolean
    Dim pendingHide As String
    Dim hasNewFrame As Boolean
    Dim delayHidePrev As Boolean

    pendingHide = m_Sprite.PendingHideName

    targetName = m_Sprite.PendingFrameName
    hasNewFrame = (targetName <> "")

    If targetName = "" Then
        targetName = m_Sprite.LastFrameShape
    End If

    If targetName = "" Then
        If pendingHide <> "" Then
            HideFrame wsGame, pendingHide
            pendingHide = ""
            m_Sprite.PendingHideName = ""
        End If
        GoTo Cleanup
    End If

    If m_Sprite.LastFrameShape = "" Then
        needsFullRefresh = True
    ElseIf hasNewFrame Then
        If m_Sprite.LastFrameShape <> targetName Then
            previousName = m_Sprite.LastFrameShape
        Else
            m_Sprite.PendingFrameName = ""
            If pendingHide <> "" Then
                If pendingHide <> targetName Then
                    HideFrame wsGame, pendingHide
                    pendingHide = ""
                    m_Sprite.PendingHideName = ""
                End If
            End If
            GoTo Cleanup
        End If
    End If

    delayHidePrev = (hasNewFrame And previousName <> "")

    ApplyLinkVisibility wsGame, targetName, previousName, needsFullRefresh, delayHidePrev

    If pendingHide <> "" Then
        If pendingHide <> targetName And pendingHide <> previousName Then
            HideFrame wsGame, pendingHide
            pendingHide = ""
        End If
    End If

    m_Sprite.LastFrameShape = targetName
    m_Sprite.PendingFrameName = ""
    If previousName <> "" Then
        m_Sprite.PendingHideName = previousName
    ElseIf pendingHide <> "" And pendingHide <> targetName Then
        m_Sprite.PendingHideName = pendingHide
    Else
        m_Sprite.PendingHideName = ""
    End If

Cleanup:
    If restoreScreenUpdating Then Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
End Sub

'===================================================================================
'                              LEGACY SUPPORT UTILITIES
'===================================================================================
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

Private Sub HideFrame(ByVal ws As Worksheet, ByVal frameName As String)
    If ws Is Nothing Then Exit Sub
    If frameName = "" Then Exit Sub

    Dim frameShape As Shape
    Set frameShape = ShapeIfExists(ws, frameName)
    If Not frameShape Is Nothing Then
        If frameShape.Visible <> msoFalse Then frameShape.Visible = msoFalse
    End If
End Sub

Private Sub ApplyLinkVisibility(ByVal ws As Worksheet, ByVal activeName As String, Optional ByVal previousName As String = "", Optional ByVal fullRefresh As Boolean = False, Optional ByVal delayHidePrevious As Boolean = False)
    If activeName = "" Then Exit Sub

    Dim activeShape As Shape
    Set activeShape = ShapeIfExists(ws, activeName)
    If Not activeShape Is Nothing Then
        If activeShape.Visible <> msoTrue Then activeShape.Visible = msoTrue
        On Error Resume Next
        activeShape.ZOrder msoBringToFront
        On Error GoTo 0
    End If

    If Not delayHidePrevious Then
        If previousName <> "" And previousName <> activeName Then
            HideFrame ws, previousName
        End If
    End If

    If fullRefresh Then
        Dim frameName As Variant
        For Each frameName In LinkFrameNames()
            If CStr(frameName) <> activeName Then
                If Not (delayHidePrevious And previousName <> "" And CStr(frameName) = previousName) Then
                    HideFrame ws, CStr(frameName)
                End If
            End If
        Next frameName
    End If
End Sub