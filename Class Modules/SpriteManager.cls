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
    Frame As Integer '0 denotes idle, positive values point to run frames
    LastFrameShape As String
    PendingFrameName As String
    PendingHideName As String
    PrevTop As Double
    PrevLeft As Double
    AnimTickCounter As Long
End Type

Private m_Sprite As SpriteData
 
Private Enum SpriteDirectionIndex
    SpriteDirUnknown = -1
    SpriteDirUp = 0
    SpriteDirRight = 1
    SpriteDirDown = 2
    SpriteDirLeft = 3
End Enum

Private Type SpriteFrameDefinition
    Idle As String
    RunFrames As Variant
End Type

Private Const DEFAULT_IDLE_DIRECTION As Long = SpriteDirDown

Private m_PlayerFrameDefs(SpriteDirUp To SpriteDirLeft) As SpriteFrameDefinition

'===================================================================================
'                              SPRITE DEFINITIONS
'===================================================================================
Private Sub ConfigureDefaultPlayerFrames()
    ClearPlayerFrameDefinitions

    SetPlayerFrameDefinition SpriteDirDown, "LinkDown1", Array("LinkDown1", "LinkDown2")
    SetPlayerFrameDefinition SpriteDirUp, "LinkUp1", Array("LinkUp1", "LinkUp2")
    SetPlayerFrameDefinition SpriteDirLeft, "LinkLeft1", Array("LinkLeft1", "LinkLeft2")
    SetPlayerFrameDefinition SpriteDirRight, "LinkRight1", Array("LinkRight1", "LinkRight2")
End Sub

Private Sub ClearPlayerFrameDefinitions()
    Dim dir As SpriteDirectionIndex
    For dir = SpriteDirUp To SpriteDirLeft
        m_PlayerFrameDefs(dir).Idle = ""
        m_PlayerFrameDefs(dir).RunFrames = Empty
    Next dir
End Sub

Private Sub SetPlayerFrameDefinition(ByVal directionIndex As SpriteDirectionIndex, ByVal idleFrame As String, ByVal runFrames As Variant)
    If Not DirectionIndexValid(directionIndex) Then Exit Sub

    m_PlayerFrameDefs(directionIndex).Idle = Trim$(idleFrame)
    m_PlayerFrameDefs(directionIndex).RunFrames = runFrames
End Sub

Public Sub ResetPlayerFrameDefinitions()
    ClearPlayerFrameDefinitions
End Sub

Public Sub ConfigurePlayerFrames(ByVal directionToken As String, ByVal idleFrame As String, Optional ByVal runFrames As Variant)
    Dim directionIndex As SpriteDirectionIndex
    directionIndex = ResolveDirectionIndexFromToken(directionToken)
    If directionIndex = SpriteDirUnknown Then Exit Sub

    Dim processedFrames As Variant
    processedFrames = WrapRunFrameVariant(runFrames, Not IsMissing(runFrames))

    SetPlayerFrameDefinition directionIndex, idleFrame, processedFrames
End Sub

Public Function GetConfiguredFrameNames() As Variant
    GetConfiguredFrameNames = PlayerFrameNames()
End Function

Private Function ResolveDirectionIndexFromToken(ByVal token As String) As SpriteDirectionIndex
    Dim normalized As String
    normalized = UCase$(Trim$(token))

    Select Case normalized
        Case "U", "UP", "N", "NORTH"
            ResolveDirectionIndexFromToken = SpriteDirUp
        Case "R", "RIGHT", "E", "EAST"
            ResolveDirectionIndexFromToken = SpriteDirRight
        Case "D", "DOWN", "S", "SOUTH"
            ResolveDirectionIndexFromToken = SpriteDirDown
        Case "L", "LEFT", "W", "WEST"
            ResolveDirectionIndexFromToken = SpriteDirLeft
        Case Else
            ResolveDirectionIndexFromToken = SpriteDirUnknown
    End Select
End Function

Private Function WrapRunFrameVariant(ByVal runFrames As Variant, ByVal hasValue As Boolean) As Variant
    If Not hasValue Then
        WrapRunFrameVariant = Empty
        Exit Function
    End If

    If IsArray(runFrames) Then
        WrapRunFrameVariant = runFrames
        Exit Function
    End If

    If IsNull(runFrames) Or IsEmpty(runFrames) Then
        WrapRunFrameVariant = Empty
        Exit Function
    End If

    Dim valueType As VbVarType
    valueType = VarType(runFrames)

    Select Case valueType
        Case vbString
            If Trim$(CStr(runFrames)) = "" Then
                WrapRunFrameVariant = Empty
            Else
                Dim frames(0 To 0) As Variant
                frames(0) = CStr(runFrames)
                WrapRunFrameVariant = frames
            End If
        Case vbVariant
            Dim stringValue As String
            stringValue = Trim$(CStr(runFrames))
            If stringValue = "" Then
                WrapRunFrameVariant = Empty
            Else
                Dim singleVariant(0 To 0) As Variant
                singleVariant(0) = stringValue
                WrapRunFrameVariant = singleVariant
            End If
        Case Else
            WrapRunFrameVariant = Empty
    End Select
End Function

Private Function ResolveDirectionIndex(ByVal direction As String) As SpriteDirectionIndex
    If InStr(direction, "U") > 0 Then
        ResolveDirectionIndex = SpriteDirUp
    ElseIf InStr(direction, "D") > 0 Then
        ResolveDirectionIndex = SpriteDirDown
    ElseIf InStr(direction, "L") > 0 Then
        ResolveDirectionIndex = SpriteDirLeft
    ElseIf InStr(direction, "R") > 0 Then
        ResolveDirectionIndex = SpriteDirRight
    Else
        ResolveDirectionIndex = SpriteDirUnknown
    End If
End Function

Private Function DirectionSuffix(ByVal directionIndex As SpriteDirectionIndex) As String
    Select Case directionIndex
        Case SpriteDirUp
            DirectionSuffix = "Up"
        Case SpriteDirRight
            DirectionSuffix = "Right"
        Case SpriteDirLeft
            DirectionSuffix = "Left"
        Case Else
            DirectionSuffix = "Down"
    End Select
End Function

Private Function DirectionIndexValid(ByVal directionIndex As SpriteDirectionIndex) As Boolean
    DirectionIndexValid = (directionIndex >= SpriteDirUp And directionIndex <= SpriteDirLeft)
End Function

Private Function PlayerRunFrameCount(ByVal directionIndex As SpriteDirectionIndex) As Long
    If Not DirectionIndexValid(directionIndex) Then Exit Function

    Dim frames As Variant
    frames = m_PlayerFrameDefs(directionIndex).RunFrames
    If IsArray(frames) Then
        PlayerRunFrameCount = UBound(frames) - LBound(frames) + 1
    Else
        PlayerRunFrameCount = 0
    End If
End Function

Private Function PlayerRunFrameName(ByVal directionIndex As SpriteDirectionIndex, ByVal frameIndex As Long) As String
    If Not DirectionIndexValid(directionIndex) Then Exit Function

    Dim frames As Variant
    frames = m_PlayerFrameDefs(directionIndex).RunFrames
    If Not IsArray(frames) Then Exit Function

    Dim lowerBound As Long
    Dim upperBound As Long
    lowerBound = LBound(frames)
    upperBound = UBound(frames)

    Dim resolvedIndex As Long
    resolvedIndex = lowerBound + frameIndex - 1
    If resolvedIndex < lowerBound Or resolvedIndex > upperBound Then Exit Function

    PlayerRunFrameName = CStr(frames(resolvedIndex))
End Function

Private Function AdvanceRunFrame(ByVal currentFrame As Long, ByVal frameCount As Long) As Long
    If frameCount <= 0 Then
        AdvanceRunFrame = 0
        Exit Function
    End If

    If currentFrame < 1 Or currentFrame >= frameCount Then
        AdvanceRunFrame = 1
    Else
        AdvanceRunFrame = currentFrame + 1
    End If
End Function

Private Function ResolveActiveFrameName(ByVal directionIndex As SpriteDirectionIndex, ByVal frameIndex As Long, ByVal isMoving As Boolean) As String
    If Not DirectionIndexValid(directionIndex) Then
        directionIndex = DEFAULT_IDLE_DIRECTION
    End If

    Dim idleName As String
    idleName = m_PlayerFrameDefs(directionIndex).Idle

    If Not isMoving Or frameIndex <= 0 Then
        ResolveActiveFrameName = idleName
        Exit Function
    End If

    Dim runName As String
    runName = PlayerRunFrameName(directionIndex, frameIndex)
    If runName <> "" Then
        ResolveActiveFrameName = runName
    Else
        ResolveActiveFrameName = idleName
    End If
End Function

Private Function IdentifyDirectionFromFrameName(ByVal frameName As String, ByRef isIdle As Boolean) As SpriteDirectionIndex
    isIdle = False
    If frameName = "" Then
        IdentifyDirectionFromFrameName = SpriteDirUnknown
        Exit Function
    End If

    Dim dir As SpriteDirectionIndex
    For dir = SpriteDirUp To SpriteDirLeft
        If StrComp(m_PlayerFrameDefs(dir).Idle, frameName, vbTextCompare) = 0 Then
            isIdle = True
            IdentifyDirectionFromFrameName = dir
            Exit Function
        End If

        Dim frames As Variant
        frames = m_PlayerFrameDefs(dir).RunFrames
        If IsArray(frames) Then
            Dim i As Long
            For i = LBound(frames) To UBound(frames)
                If StrComp(CStr(frames(i)), frameName, vbTextCompare) = 0 Then
                    IdentifyDirectionFromFrameName = dir
                    Exit Function
                End If
            Next i
        End If
    Next dir

    IdentifyDirectionFromFrameName = SpriteDirUnknown
End Function

Private Function ResolveRunFrameIndex(ByVal directionIndex As SpriteDirectionIndex, ByVal frameName As String) As Long
    If Not DirectionIndexValid(directionIndex) Then Exit Function
    If frameName = "" Then Exit Function

    Dim frames As Variant
    frames = m_PlayerFrameDefs(directionIndex).RunFrames
    If Not IsArray(frames) Then Exit Function

    Dim lowerBound As Long
    lowerBound = LBound(frames)

    Dim i As Long
    For i = lowerBound To UBound(frames)
        If StrComp(CStr(frames(i)), frameName, vbTextCompare) = 0 Then
            ResolveRunFrameIndex = (i - lowerBound) + 1
            Exit Function
        End If
    Next i

    ResolveRunFrameIndex = 0
End Function

Private Function AllConfiguredFrameNames() As Variant
    Dim frameCollection As Collection
    Set frameCollection = New Collection

    Dim dir As SpriteDirectionIndex
    For dir = SpriteDirUp To SpriteDirLeft
        AppendFrameName frameCollection, m_PlayerFrameDefs(dir).Idle

        Dim frames As Variant
        frames = m_PlayerFrameDefs(dir).RunFrames
        If IsArray(frames) Then
            Dim i As Long
            For i = LBound(frames) To UBound(frames)
                AppendFrameName frameCollection, CStr(frames(i))
            Next i
        End If
    Next dir

    If frameCollection.Count = 0 Then
        AllConfiguredFrameNames = Array()
    Else
        Dim result() As String
        ReDim result(0 To frameCollection.Count - 1)
        Dim index As Long
        For index = 1 To frameCollection.Count
            result(index - 1) = CStr(frameCollection(index))
        Next index
        AllConfiguredFrameNames = result
    End If
End Function

Private Sub AppendFrameName(ByVal frames As Collection, ByVal frameName As String)
    If frameName = "" Then Exit Sub
    On Error Resume Next
    frames.Add frameName, frameName
    On Error GoTo 0
End Sub

Private Function ResolveInitialSpriteShape(ByVal ws As Worksheet) As Shape
    If ws Is Nothing Then Exit Function

    Dim preferredName As String
    preferredName = m_PlayerFrameDefs(DEFAULT_IDLE_DIRECTION).Idle
    If preferredName <> "" Then
        Set ResolveInitialSpriteShape = ShapeIfExists(ws, preferredName)
        If Not ResolveInitialSpriteShape Is Nothing Then Exit Function
    End If

    Dim frameName As Variant
    For Each frameName In PlayerFrameNames()
        Set ResolveInitialSpriteShape = ShapeIfExists(ws, CStr(frameName))
        If Not ResolveInitialSpriteShape Is Nothing Then Exit Function
    Next frameName

    Set ResolveInitialSpriteShape = Nothing
End Function

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
    m_Sprite.Frame = 0
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
    ConfigureDefaultPlayerFrames
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

    Dim bindDir As SpriteDirectionIndex
    Dim isIdleFrame As Boolean
    bindDir = IdentifyDirectionFromFrameName(spriteName, isIdleFrame)
    If bindDir = SpriteDirUnknown Then
        bindDir = DEFAULT_IDLE_DIRECTION
    End If

    If isIdleFrame Then
        m_Sprite.Frame = 0
    Else
        m_Sprite.Frame = ResolveRunFrameIndex(bindDir, spriteName)
        If m_Sprite.Frame = 0 Then
            m_Sprite.Frame = 1
        End If
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
        Dim wsCurrent As Worksheet
        Set wsCurrent = Sheets(gs.CurrentScreen)
        If Not wsCurrent Is Nothing Then
            Set m_Sprite.LinkSprite = ResolveInitialSpriteShape(wsCurrent)
        End If

        If Not m_Sprite.LinkSprite Is Nothing Then
            m_Sprite.Top = m_Sprite.LinkSprite.Top
            m_Sprite.Left = m_Sprite.LinkSprite.Left

            Dim resetDir As SpriteDirectionIndex
            Dim resetIdle As Boolean
            resetDir = IdentifyDirectionFromFrameName(m_Sprite.LinkSprite.Name, resetIdle)
            If resetIdle Then
                m_Sprite.Frame = 0
            Else
                m_Sprite.Frame = ResolveRunFrameIndex(resetDir, m_Sprite.LinkSprite.Name)
            End If

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

Public Sub ResyncFramePositions()
    On Error Resume Next
    SyncAllLinkFrames
    On Error GoTo 0
End Sub

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

    If facingDir = "" Then
        If Not gs Is Nothing Then
            facingDir = gs.LastDir
        Else
            facingDir = ""
        End If
    End If

    Dim directionIndex As SpriteDirectionIndex
    directionIndex = ResolveDirectionIndex(facingDir)
    If directionIndex = SpriteDirUnknown Then
        directionIndex = DEFAULT_IDLE_DIRECTION
    End If

    Dim availableFrames As Long
    availableFrames = PlayerRunFrameCount(directionIndex)

    If Not movedThisTick Then
        If m_Sprite.Frame <> 0 Then m_Sprite.Frame = 0
        m_Sprite.AnimTickCounter = 0
    Else
        If availableFrames <= 0 Then
            m_Sprite.Frame = 0
            m_Sprite.AnimTickCounter = 0
        Else
            If m_Sprite.Frame <= 0 Or m_Sprite.Frame > availableFrames Then
                m_Sprite.Frame = 1
                m_Sprite.AnimTickCounter = 0
            End If

            If ANIMATION_TICKS_PER_FRAME > 0 Then
                m_Sprite.AnimTickCounter = m_Sprite.AnimTickCounter + 1
                If m_Sprite.AnimTickCounter >= ANIMATION_TICKS_PER_FRAME Then
                    m_Sprite.AnimTickCounter = 0
                    m_Sprite.Frame = AdvanceRunFrame(m_Sprite.Frame, availableFrames)
                End If
            End If
        End If
    End If

    m_Sprite.Top = newTop
    m_Sprite.Left = newLeft

    Dim activeFrameName As String
    activeFrameName = ResolveActiveFrameName(directionIndex, m_Sprite.Frame, movedThisTick)

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

    If activeFrameName = "" Then Exit Sub

    Dim activeShape As Shape
    Set activeShape = ShapeIfExists(ws, activeFrameName)
    If activeShape Is Nothing Then Exit Sub

    Set m_Sprite.LinkSprite = activeShape
    m_Sprite.PendingFrameName = activeShape.Name

    If Not m_Sprite.LinkSprite Is Nothing Then
        m_Sprite.LinkSprite.Top = m_Sprite.Top
        m_Sprite.LinkSprite.Left = m_Sprite.Left
    End If

    SyncAllLinkFrames

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

    Dim previousCalculation As XlCalculation
    Dim restoreCalculation As Boolean
    If Application.Calculation <> xlCalculationManual Then
        previousCalculation = Application.Calculation
        Application.Calculation = xlCalculationManual
        restoreCalculation = True
    End If

    Dim targetName As String
    Dim previousName As String
    Dim previousBase As String
    Dim targetBase As String
    Dim needsFullRefresh As Boolean

    targetName = m_Sprite.PendingFrameName
    If targetName = "" Then
        targetName = m_Sprite.LastFrameShape
    End If

    If targetName = "" Then GoTo Cleanup

    If m_Sprite.LastFrameShape <> "" Then
        If m_Sprite.LastFrameShape <> targetName Then
            previousName = m_Sprite.LastFrameShape
            previousBase = BaseFrameName(previousName)
            targetBase = BaseFrameName(targetName)
            If previousBase <> targetBase Then
                needsFullRefresh = True
            End If
        End If
    Else
        needsFullRefresh = True
    End If

    ApplyLinkVisibility wsGame, targetName, previousName, needsFullRefresh

    Dim frameName As Variant
    For Each frameName In PlayerFrameNames()
        Dim candidate As String
        candidate = CStr(frameName)
        If candidate <> targetName Then
            HideFrame wsGame, candidate
        End If
    Next frameName

    m_Sprite.LastFrameShape = targetName
    m_Sprite.PendingFrameName = ""
    m_Sprite.PendingHideName = ""

Cleanup:
    m_Sprite.PendingFrameName = ""
    If restoreCalculation Then Application.Calculation = previousCalculation
    If restoreScreenUpdating Then Application.ScreenUpdating = previousScreenUpdating
    On Error GoTo 0
End Sub

Private Function BaseFrameName(ByVal frameName As String) As String
    Dim nameLength As Long
    nameLength = Len(frameName)
    If nameLength <= 1 Then
        BaseFrameName = frameName
        Exit Function
    End If

    Dim index As Long
    index = nameLength
    Do While index >= 1
        Dim currentChar As String
        currentChar = Mid$(frameName, index, 1)
        If currentChar Like "#" Then
            index = index - 1
        Else
            Exit Do
        End If
    Loop

    If index < nameLength Then
        BaseFrameName = VBA.Left(frameName, index)
    Else
        BaseFrameName = frameName
    End If
End Function

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
    
    Dim bounceDir As SpriteDirectionIndex
    Dim bounceIdle As Boolean
    bounceDir = IdentifyDirectionFromFrameName(linkShape.Name, bounceIdle)
    If bounceDir = SpriteDirUnknown Then Exit Sub

    Select Case bounceDir
        Case SpriteDirDown
            blockingValue = originCell.Offset(-1, 2).Value
            If blockingValue = "" Then
                linkShape.Top = linkShape.Top - bounceSpeed
            End If
        Case SpriteDirUp
            blockingValue = originCell.Offset(4, 2).Value
            If blockingValue = "" Then
                linkShape.Top = linkShape.Top + bounceSpeed
            End If
        Case SpriteDirLeft
            blockingValue = originCell.Offset(2, 4).Value
            If blockingValue = "" Then
                linkShape.Left = linkShape.Left + bounceSpeed
            End If
        Case SpriteDirRight
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
    For Each frameName In PlayerFrameNames()
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
    Dim dirIndex As SpriteDirectionIndex
    dirIndex = ResolveDirectionIndex(direction)
    If dirIndex = SpriteDirUnknown Then
        dirIndex = DEFAULT_IDLE_DIRECTION
    End If

    ResolveDirection = DirectionSuffix(dirIndex)
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
    For Each frameName In PlayerFrameNames()
        Dim frameShape As Shape
        Set frameShape = ShapeIfExists(wsGame, CStr(frameName))
        If Not frameShape Is Nothing Then
            frameShape.Top = m_Sprite.Top
            frameShape.Left = m_Sprite.Left
        End If
    Next frameName
    On Error GoTo 0
End Sub

Private Function PlayerFrameNames() As Variant
    PlayerFrameNames = AllConfiguredFrameNames()
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

Private Sub ApplyLinkVisibility(ByVal ws As Worksheet, ByVal activeName As String, Optional ByVal previousName As String = "", Optional ByVal fullRefresh As Boolean = False)
    If activeName = "" Then Exit Sub

    Dim activeShape As Shape
    Set activeShape = ShapeIfExists(ws, activeName)
    If Not activeShape Is Nothing Then
        If activeShape.Visible <> msoTrue Then activeShape.Visible = msoTrue
        On Error Resume Next
        activeShape.ZOrder msoBringToFront
        On Error GoTo 0
    End If

    If previousName <> "" And previousName <> activeName Then
        HideFrame ws, previousName
    End If

    If fullRefresh Then
        Dim frameName As Variant
        For Each frameName In PlayerFrameNames()
            If CStr(frameName) <> activeName And CStr(frameName) <> previousName Then
                HideFrame ws, CStr(frameName)
            End If
        Next frameName
    End If
End Sub

