'@Folder("Engine.Managers")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' SPRITE MANAGER
'═══════════════════════════════════════════════════════════════════════════════
' Handles sprite operations - positions, frames, visibility, animation discovery
'═══════════════════════════════════════════════════════════════════════════════

Private Type SpriteData
    PlayerSprite As Shape
    Top As Double
    Left As Double
    Frame As Integer '0 denotes idle, positive values point to run frames
    LastFrameShape As String
    PendingFrameName As String
    PendingHideName As String
    PrevTop As Double
    PrevLeft As Double
    AnimTickCounter As Long
    LightShape As Shape
    LightShapeName As String
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

' Discovered sprite frame information
Private Type DiscoveredFrame
    ShapeName As String
    Direction As String      ' U, D, L, R, or empty for omnidirectional
    State As String          ' Idle, Run, Attack, etc. (empty = default)
    FrameNumber As Long      ' 0 = no number, 1+ = animation frame
    IsValid As Boolean
End Type

Private Const DEFAULT_IDLE_DIRECTION As Long = SpriteDirDown
Private Const LIGHT_OVERLAY_NAME As String = "Light"
Private Const DEFAULT_SPRITE_BASE_NAME As String = "Player"

Private m_PlayerFrameDefs(SpriteDirUp To SpriteDirLeft) As SpriteFrameDefinition
Private m_DiscoveredFrames() As DiscoveredFrame
Private m_DiscoveryBaseName As String
Private m_UseDiscovery As Boolean

'===================================================================================
'                              SPRITE DEFINITIONS
'===================================================================================
Private Sub ConfigureDefaultPlayerFrames()
    ClearPlayerFrameDefinitions

    SetPlayerFrameDefinition SpriteDirDown, "PlayerDown1", Array("PlayerDown1", "PlayerDown2")
    SetPlayerFrameDefinition SpriteDirUp, "PlayerUp1", Array("PlayerUp1", "PlayerUp2")
    SetPlayerFrameDefinition SpriteDirLeft, "PlayerLeft1", Array("PlayerLeft1", "PlayerLeft2")
    SetPlayerFrameDefinition SpriteDirRight, "PlayerRight1", Array("PlayerRight1", "PlayerRight2")
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
    GetConfiguredFrameNames = AllConfiguredFrameNames()
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
    ' Convert single value or array into standardized array format
    If Not hasValue Then Exit Function
    If IsNull(runFrames) Or IsEmpty(runFrames) Then Exit Function
    If IsArray(runFrames) Then
        WrapRunFrameVariant = runFrames
        Exit Function
    End If

    ' Convert single string to single-element array
    Dim strValue As String
    strValue = Trim$(CStr(runFrames))
    If strValue = "" Then Exit Function

    Dim result(0 To 0) As Variant
    result(0) = strValue
    WrapRunFrameVariant = result
End Function

'===================================================================================
'                              DIRECTION RESOLUTION
'===================================================================================
Private Function ResolveDirectionIndex(ByVal direction As String) As SpriteDirectionIndex
    ' Resolve to primary direction for sprite facing
    ' For diagonals, prioritize vertical over horizontal for visual consistency
    Dim hasU As Boolean, hasD As Boolean, hasL As Boolean, hasR As Boolean
    hasU = InStr(direction, "U") > 0
    hasD = InStr(direction, "D") > 0
    hasL = InStr(direction, "L") > 0
    hasR = InStr(direction, "R") > 0

    ' Cancel opposite directions
    If hasU And hasD Then
        hasU = False
        hasD = False
    End If
    If hasL And hasR Then
        hasL = False
        hasR = False
    End If

    ' Prioritize vertical for diagonal movement
    If hasU Then
        ResolveDirectionIndex = SpriteDirUp
    ElseIf hasD Then
        ResolveDirectionIndex = SpriteDirDown
    ElseIf hasL Then
        ResolveDirectionIndex = SpriteDirLeft
    ElseIf hasR Then
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

'===================================================================================
'                              FRAME MANAGEMENT
'===================================================================================
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
    For Each frameName In AllConfiguredFrameNames()
        Set ResolveInitialSpriteShape = ShapeIfExists(ws, CStr(frameName))
        If Not ResolveInitialSpriteShape Is Nothing Then Exit Function
    Next frameName

    Set ResolveInitialSpriteShape = Nothing
End Function

'===================================================================================
'                              SPRITE DISCOVERY ENGINE
'===================================================================================
' Auto-discovers sprites on a sheet based on naming patterns
' Supports flexible configurations: single sprite, directional, animated, state-based
'===================================================================================

Public Sub DiscoverSpritesOnSheet(ByVal ws As Worksheet, Optional ByVal baseName As String = DEFAULT_SPRITE_BASE_NAME)
    ' Scan worksheet for all shapes matching base name pattern and build sprite registry
    m_DiscoveryBaseName = Trim$(baseName)
    If m_DiscoveryBaseName = "" Then m_DiscoveryBaseName = DEFAULT_SPRITE_BASE_NAME
    m_UseDiscovery = True
    
    ReDim m_DiscoveredFrames(0 To 0)
    Dim frameCount As Long
    frameCount = 0
    
    If ws Is Nothing Then Exit Sub
    
    Dim shp As Shape
    For Each shp In ws.Shapes
        Dim frameInfo As DiscoveredFrame
        frameInfo = ParseSpriteFrameName(shp.Name, m_DiscoveryBaseName)
        
        If frameInfo.IsValid Then
            ' Expand array and add discovered frame
            If frameCount > 0 Then ReDim Preserve m_DiscoveredFrames(0 To frameCount)
            m_DiscoveredFrames(frameCount) = frameInfo
            frameCount = frameCount + 1
        End If
    Next shp
    
    ' If no sprites discovered, fall back to manual configuration
    If frameCount = 0 Then
        m_UseDiscovery = False
        ConfigureDefaultPlayerFrames
    Else
        ' Build frame definitions from discovered sprites
        BuildFrameDefinitionsFromDiscovery
    End If
End Sub

Private Function ParseSpriteFrameName(ByVal shapeName As String, ByVal baseName As String) As DiscoveredFrame
    ' Parse sprite name to extract: Direction, State, Frame number
    ' Supported patterns:
    '   - Player (single omnidirectional sprite)
    '   - PlayerIdle (state-based, no direction)
    '   - PlayerDown (directional, no animation)
    '   - PlayerDown1, PlayerDown2 (directional with animation frames)
    '   - PlayerIdleDown1 (state + direction + frame)
    '   - PlayerRunDown1, PlayerAttackLeft2 (action states)
    
    Dim result As DiscoveredFrame
    result.IsValid = False
    result.ShapeName = shapeName
    result.Direction = ""
    result.State = ""
    result.FrameNumber = 0
    
    ' Must start with base name
    If Mid$(shapeName, 1, Len(baseName)) <> baseName Then
        ParseSpriteFrameName = result
        Exit Function
    End If
    
    ' Exclude overlay sprites from discovery (shields, jumps, falls, win, shadow, heart, sword)
    ' These are managed separately by game logic, not the animation system
    Dim suffix As String
    suffix = Mid$(shapeName, Len(baseName) + 1)
    If InStr(1, suffix, "Shield", vbTextCompare) > 0 Then Exit Function
    If InStr(1, suffix, "Jump", vbTextCompare) > 0 Then Exit Function
    If InStr(1, suffix, "Fall", vbTextCompare) > 0 Then Exit Function
    If InStr(1, suffix, "Win", vbTextCompare) > 0 Then Exit Function
    If InStr(1, suffix, "Shadow", vbTextCompare) > 0 Then Exit Function
    If InStr(1, suffix, "Heart", vbTextCompare) > 0 Then Exit Function
    If InStr(1, suffix, "Sword", vbTextCompare) > 0 Then Exit Function
    
    ' If empty suffix, it's a single omnidirectional sprite
    If suffix = "" Then
        result.IsValid = True
        ParseSpriteFrameName = result
        Exit Function
    End If
    
    ' Try to extract frame number from end (e.g., "Down1" -> direction="Down", frame=1)
    Dim frameNum As Long
    frameNum = ExtractTrailingNumber(suffix)
    If frameNum > 0 Then
    suffix = Mid$(suffix, 1, Len(suffix) - Len(CStr(frameNum)))
        result.FrameNumber = frameNum
    End If
    
    ' Try to extract direction from end (Up, Down, Left, Right)
    Dim direction As String
    direction = ExtractLeadingDirection(suffix)
    If direction <> "" Then
        result.Direction = direction
        suffix = Mid$(suffix, Len(direction) + 1)
    End If
    
    ' Remaining suffix is the state (Idle, Run, Attack, etc.)
    If suffix <> "" Then
        result.State = suffix
    End If
    
    result.IsValid = True
    ParseSpriteFrameName = result
End Function

Private Function ExtractTrailingNumber(ByVal text As String) As Long
    ' Extract number from end of string (e.g., "Down2" -> 2)
    Dim i As Long
    Dim ch As String
    Dim numStr As String
    
    For i = Len(text) To 1 Step -1
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = ch & numStr
        Else
            Exit For
        End If
    Next i
    
    If numStr <> "" Then
        On Error Resume Next
        ExtractTrailingNumber = CLng(numStr)
        On Error GoTo 0
    End If
End Function

Private Function ExtractLeadingDirection(ByVal text As String) As String
    If text = "" Then Exit Function

    Dim directions As Variant
    directions = Array("Down", "Up", "Left", "Right")

    Dim dir As Variant
    For Each dir In directions
        Dim dirStr As String
        dirStr = CStr(dir)
        If Len(text) >= Len(dirStr) Then
            If StrComp(Mid$(text, 1, Len(dirStr)), dirStr, vbTextCompare) = 0 Then
                ExtractLeadingDirection = dirStr
                Exit Function
            End If
        End If
    Next dir
End Function

Private Sub BuildFrameDefinitionsFromDiscovery()
    ' Convert discovered frames into traditional frame definitions for backward compatibility
    ClearPlayerFrameDefinitions
    
    Dim i As Long
    For i = LBound(m_DiscoveredFrames) To UBound(m_DiscoveredFrames)
        Dim frame As DiscoveredFrame
        frame = m_DiscoveredFrames(i)
        
        If Not frame.IsValid Then Exit For
        
        ' Map direction string to index
        Dim dirIndex As SpriteDirectionIndex
        dirIndex = ParseDirectionToIndex(frame.Direction)
        
        ' Skip if direction not recognized (omnidirectional sprites handled separately)
        If dirIndex = SpriteDirUnknown And frame.Direction <> "" Then GoTo NextFrame
        
        ' If omnidirectional (no direction), apply to all directions
        If frame.Direction = "" Then
            Dim dir As SpriteDirectionIndex
            For dir = SpriteDirUp To SpriteDirLeft
                If m_PlayerFrameDefs(dir).Idle = "" Then
                    m_PlayerFrameDefs(dir).Idle = frame.ShapeName
                    m_PlayerFrameDefs(dir).RunFrames = Array(frame.ShapeName)
                End If
            Next dir
            GoTo NextFrame
        End If
        
        ' Set idle frame (frame 0 or 1, or if no frames exist yet)
        If frame.FrameNumber <= 1 And m_PlayerFrameDefs(dirIndex).Idle = "" Then
            m_PlayerFrameDefs(dirIndex).Idle = frame.ShapeName
        End If
        
        ' Add to run frames array
        If frame.FrameNumber > 0 Then
            AddFrameToRunArray dirIndex, frame.ShapeName
        End If
        
NextFrame:
    Next i
End Sub

Private Function ParseDirectionToIndex(ByVal direction As String) As SpriteDirectionIndex
    Select Case UCase$(Trim$(direction))
        Case "UP", "U"
            ParseDirectionToIndex = SpriteDirUp
        Case "DOWN", "D"
            ParseDirectionToIndex = SpriteDirDown
        Case "LEFT", "L"
            ParseDirectionToIndex = SpriteDirLeft
        Case "RIGHT", "R"
            ParseDirectionToIndex = SpriteDirRight
        Case Else
            ParseDirectionToIndex = SpriteDirUnknown
    End Select
End Function

Private Sub AddFrameToRunArray(ByVal dirIndex As SpriteDirectionIndex, ByVal frameName As String)
    ' Add frame to run frames array, creating or expanding as needed
    Dim currentFrames As Variant
    currentFrames = m_PlayerFrameDefs(dirIndex).RunFrames
    
    If IsEmpty(currentFrames) Then
        m_PlayerFrameDefs(dirIndex).RunFrames = Array(frameName)
    ElseIf IsArray(currentFrames) Then
        ' Expand array
        Dim newSize As Long
        newSize = UBound(currentFrames) + 1
        ReDim Preserve currentFrames(LBound(currentFrames) To newSize)
        currentFrames(newSize) = frameName
        m_PlayerFrameDefs(dirIndex).RunFrames = currentFrames
    End If
End Sub

Public Function GetDiscoveredFrameCount() As Long
    ' Returns count of discovered frames (0 if using manual configuration)
    If Not m_UseDiscovery Then Exit Function
    If Not HasDiscoveredFrames() Then Exit Function

    Dim i As Long
    For i = LBound(m_DiscoveredFrames) To UBound(m_DiscoveredFrames)
        If Not m_DiscoveredFrames(i).IsValid Then Exit Function
        GetDiscoveredFrameCount = GetDiscoveredFrameCount + 1
    Next i
End Function

Private Function HasDiscoveredFrames() As Boolean
    On Error Resume Next
    Dim lowerBound As Long
    Dim upperBound As Long

    lowerBound = LBound(m_DiscoveredFrames)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    upperBound = UBound(m_DiscoveredFrames)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    HasDiscoveredFrames = (upperBound >= lowerBound)
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
    Set m_Sprite.PlayerSprite = Nothing
    m_Sprite.Top = 0
    m_Sprite.Left = 0
    m_Sprite.Frame = 0
    m_Sprite.LastFrameShape = ""
    m_Sprite.PendingFrameName = ""
    m_Sprite.PendingHideName = ""
    m_Sprite.PrevTop = 0
    m_Sprite.PrevLeft = 0
    m_Sprite.AnimTickCounter = 0
    Set m_Sprite.LightShape = Nothing
    m_Sprite.LightShapeName = ""
End Sub

'===================================================================================
'                              SETUP & LIFECYCLE
'===================================================================================
Public Sub Initialize()
    ConfigureDefaultPlayerFrames
    ResetState
End Sub

Public Sub BindPlayerSprite(ByVal screenName As String, ByVal spriteName As String)
    ResetState
    If screenName = "" Or spriteName = "" Then Exit Sub

    On Error GoTo ErrorHandler

    Set m_Sprite.PlayerSprite = Sheets(screenName).Shapes(spriteName)

    If m_Sprite.PlayerSprite Is Nothing Then
        Err.Raise vbObjectError + 1, "SpriteManager.BindPlayerSprite", "Sprite '" & spriteName & "' not found on " & screenName
    End If

    InitializeSpriteState m_Sprite.PlayerSprite, spriteName
    Exit Sub

ErrorHandler:
    Debug.Print "SpriteManager.BindPlayerSprite Error: " & Err.Description
    On Error GoTo 0
End Sub

Public Sub Reset()
    On Error Resume Next
    HideAllConfiguredFrames
    ResetState
    
    Dim gs As GameState
    Set gs = GameStateInstance
    If gs Is Nothing Or gs.CurrentScreen = "" Then Exit Sub
    
    Dim wsCurrent As Worksheet
    Set wsCurrent = Sheets(gs.CurrentScreen)
    If wsCurrent Is Nothing Then Exit Sub

    Set m_Sprite.PlayerSprite = ResolveInitialSpriteShape(wsCurrent)
    If m_Sprite.PlayerSprite Is Nothing Then Exit Sub

    InitializeSpriteState m_Sprite.PlayerSprite, m_Sprite.PlayerSprite.Name
    On Error GoTo 0
End Sub

Private Sub InitializeSpriteState(ByVal sprite As Shape, ByVal spriteName As String)
    ' Common initialization logic for binding and resetting sprites
    m_Sprite.Top = sprite.Top
    m_Sprite.Left = sprite.Left
    m_Sprite.PrevTop = m_Sprite.Top
    m_Sprite.PrevLeft = m_Sprite.Left

    BindLightOverlay sprite.Parent

    Dim isIdleFrame As Boolean
    Dim dirIndex As SpriteDirectionIndex
    dirIndex = IdentifyDirectionFromFrameName(spriteName, isIdleFrame)
    If dirIndex = SpriteDirUnknown Then dirIndex = DEFAULT_IDLE_DIRECTION

    If isIdleFrame Then
        m_Sprite.Frame = 0
    Else
        m_Sprite.Frame = ResolveRunFrameIndex(dirIndex, spriteName)
        If m_Sprite.Frame = 0 Then m_Sprite.Frame = 1
    End If

    m_Sprite.LastFrameShape = sprite.Name
    m_Sprite.PendingFrameName = m_Sprite.LastFrameShape
    ApplyLinkVisibility sprite.Parent, m_Sprite.PendingFrameName, "", True
    m_Sprite.PendingFrameName = ""
    m_Sprite.PendingHideName = ""
    m_Sprite.AnimTickCounter = 0

    UpdateLightOverlay sprite
End Sub

Public Sub Destroy()
    HideAllConfiguredFrames
    ResetState
End Sub

Private Function GetActiveWorksheet() As Worksheet
    ' Resolve current game worksheet from GameState or sprite parent
    Dim gs As GameState
    Set gs = GameStateInstance
    If Not gs Is Nothing And gs.CurrentScreen <> "" Then
        On Error Resume Next
        Set GetActiveWorksheet = Sheets(gs.CurrentScreen)
        On Error GoTo 0
        If Not GetActiveWorksheet Is Nothing Then Exit Function
    End If

    If Not m_Sprite.PlayerSprite Is Nothing Then
        On Error Resume Next
        Set GetActiveWorksheet = m_Sprite.PlayerSprite.Parent
        On Error GoTo 0
    End If
End Function

Private Sub HideAllConfiguredFrames()
    Dim ws As Worksheet
    Set ws = GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub

    Dim frames As Variant
    frames = AllConfiguredFrameNames()
    If IsEmpty(frames) Then Exit Sub
    If IsArray(frames) Then
        If UBound(frames) < LBound(frames) Then Exit Sub
    Else
        Exit Sub
    End If

    Dim renderer As BatchRenderer
    On Error Resume Next
    Set renderer = BatchRendererInstance()
    On Error GoTo 0

    Dim useRenderer As Boolean
    useRenderer = Not renderer Is Nothing

    Dim frameName As Variant
    For Each frameName In frames
        Dim shape As Shape
        Set shape = ShapeIfExists(ws, CStr(frameName))
        If Not shape Is Nothing Then
            If useRenderer Then
                renderer.QueueVisibility ws, shape.Name, False
            Else
                On Error Resume Next
                shape.Visible = msoFalse
                On Error GoTo 0
            End If
        End If
    Next frameName

    If useRenderer Then
        On Error Resume Next
        renderer.ApplyBatch
        On Error GoTo 0
    End If
End Sub


' Essential properties
Public Property Get PlayerSprite() As Object
    Set PlayerSprite = m_Sprite.PlayerSprite
End Property

Public Property Get PlayerSpriteTop() As Double
    PlayerSpriteTop = m_Sprite.Top
End Property

Public Property Let PlayerSpriteTop(ByVal positionTop As Double)
    m_Sprite.Top = positionTop
End Property

Public Property Get PlayerSpriteLeft() As Double
    PlayerSpriteLeft = m_Sprite.Left
End Property

Public Property Let PlayerSpriteLeft(ByVal positionLeft As Double)
    m_Sprite.Left = positionLeft
End Property

Public Property Get PlayerSpriteFrame() As Integer
    PlayerSpriteFrame = m_Sprite.Frame
End Property

Public Property Let PlayerSpriteFrame(ByVal frameIndex As Integer)
    m_Sprite.Frame = frameIndex
End Property

Public Sub ResyncFramePositions()
    On Error Resume Next
    SyncAllPlayerFrames
    UpdateLightOverlay m_Sprite.PlayerSprite
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
    pixelsPerTick = moveSpeed * SPEED_MULTIPLIER
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

    ' Resolve facing direction with fallback chain
    If facingDir = "" And Not gs Is Nothing Then
        facingDir = gs.LastDir
    End If
    If facingDir = "" Then
        facingDir = "D" ' Ultimate fallback to down
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

    ' Update position state before resolving frame
    m_Sprite.Top = newTop
    m_Sprite.Left = newLeft

    ' Determine which frame to display
    Dim activeFrameName As String
    activeFrameName = ResolveActiveFrameName(directionIndex, m_Sprite.Frame, movedThisTick)

    Dim ws As Worksheet
    Set ws = GetActiveWorksheet()
    If ws Is Nothing Then Exit Sub
    If activeFrameName = "" Then Exit Sub

    Dim activeShape As Shape
    Set activeShape = ShapeIfExists(ws, activeFrameName)
    If activeShape Is Nothing Then Exit Sub

    ' Update sprite reference and pending frame
    Set m_Sprite.PlayerSprite = activeShape
    m_Sprite.PendingFrameName = activeShape.Name

    ' Sync all frame positions to new location
    SyncAllPlayerFrames

    On Error GoTo 0
End Sub

'===================================================================================
'                              RENDERING HELPERS
'===================================================================================
Public Sub UpdatePosition()
    On Error Resume Next
    If m_Sprite.PlayerSprite Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = m_Sprite.PlayerSprite.Parent
    BatchRendererInstance().QueuePosition ws, m_Sprite.PlayerSprite.Name, m_Sprite.Top, m_Sprite.Left
    SyncAllPlayerFrames
    On Error GoTo 0
End Sub

Public Sub RenderInterpolated(ByVal alpha As Double)
    ' Render sprite at interpolated position between Prev and current
    On Error Resume Next
    If m_Sprite.PlayerSprite Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    If alpha < 0# Then alpha = 0#
    If alpha > 1# Then alpha = 1#

    Dim interpTop As Double
    Dim interpLeft As Double
    interpTop = m_Sprite.PrevTop + (m_Sprite.Top - m_Sprite.PrevTop) * alpha
    interpLeft = m_Sprite.PrevLeft + (m_Sprite.Left - m_Sprite.PrevLeft) * alpha

    Dim ws As Worksheet
    Set ws = m_Sprite.PlayerSprite.Parent
    BatchRendererInstance().QueuePosition ws, m_Sprite.PlayerSprite.Name, interpTop, interpLeft
    SyncAllPlayerFrames
    On Error GoTo 0
End Sub

Public Sub UpdateVisibility()
    On Error Resume Next
    
    Dim wsGame As Worksheet
    Set wsGame = GetActiveWorksheet()
    If wsGame Is Nothing Then Exit Sub

    ' Determine target frame
    Dim targetName As String
    targetName = m_Sprite.PendingFrameName
    If targetName = "" Then targetName = m_Sprite.LastFrameShape
    If targetName = "" Then Exit Sub

    ' Check if full refresh needed (direction change)
    Dim needsFullRefresh As Boolean
    needsFullRefresh = (m_Sprite.LastFrameShape = "")
    
    If Not needsFullRefresh And m_Sprite.LastFrameShape <> targetName Then
        Dim previousBase As String, targetBase As String
        previousBase = BaseFrameName(m_Sprite.LastFrameShape)
        targetBase = BaseFrameName(targetName)
        needsFullRefresh = (previousBase <> targetBase)
    End If

    ' Apply visibility changes
    ApplyLinkVisibility wsGame, targetName, m_Sprite.LastFrameShape, needsFullRefresh

    ' Hide non-active frames (only if full refresh needed for performance)
    If needsFullRefresh Then
        Dim frameName As Variant
        For Each frameName In AllConfiguredFrameNames()
            Dim candidate As String
            candidate = CStr(frameName)
            If candidate <> targetName And candidate <> m_Sprite.LastFrameShape Then
                HideFrame wsGame, candidate
            End If
        Next frameName
    End If

    ' Update state
    m_Sprite.LastFrameShape = targetName
    m_Sprite.PendingFrameName = ""
    m_Sprite.PendingHideName = ""

    On Error GoTo 0
End Sub

Private Function BaseFrameName(ByVal frameName As String) As String
    ' Strip trailing digits from frame name (e.g., "LinkDown2" -> "LinkDown")
    Dim nameLength As Long
    nameLength = Len(frameName)
    If nameLength <= 1 Then
        BaseFrameName = frameName
        Exit Function
    End If

    ' Find last non-digit character
    Dim i As Long
    For i = nameLength To 1 Step -1
        If Not Mid$(frameName, i, 1) Like "#" Then
            BaseFrameName = Mid$(frameName, 1, i)
            Exit Function
        End If
    Next i

    ' All digits (unlikely but handle it)
    BaseFrameName = frameName
End Function

'===================================================================================
'                              LEGACY SUPPORT UTILITIES
'===================================================================================
Public Sub ApplyPlayerBounce(ByVal bounceSpeed As Long)
    On Error Resume Next
    If bounceSpeed <= 0 Then Exit Sub
    
    Dim playerShape As Shape
    Set playerShape = m_Sprite.PlayerSprite
    If playerShape Is Nothing Then Exit Sub
    
    Dim originCell As Range
    Set originCell = playerShape.TopLeftCell
    If originCell Is Nothing Then Exit Sub
    
    Dim blockingValue As Variant
    
    Dim bounceDir As SpriteDirectionIndex
    Dim bounceIdle As Boolean
    bounceDir = IdentifyDirectionFromFrameName(playerShape.Name, bounceIdle)
    If bounceDir = SpriteDirUnknown Then Exit Sub

    Dim newTop As Double
    Dim newLeft As Double
    newTop = playerShape.Top
    newLeft = playerShape.Left
    
    Select Case bounceDir
        Case SpriteDirDown
            blockingValue = originCell.Offset(-1, 2).Value
            If blockingValue = "" Then
                newTop = newTop - bounceSpeed
            End If
        Case SpriteDirUp
            blockingValue = originCell.Offset(4, 2).Value
            If blockingValue = "" Then
                newTop = newTop + bounceSpeed
            End If
        Case SpriteDirLeft
            blockingValue = originCell.Offset(2, 4).Value
            If blockingValue = "" Then
                newLeft = newLeft + bounceSpeed
            End If
        Case SpriteDirRight
            blockingValue = originCell.Offset(2, -1).Value
            If blockingValue = "" Then
                newLeft = newLeft - bounceSpeed
            End If
    End Select
    
    m_Sprite.Top = newTop
    m_Sprite.Left = newLeft
    BatchRendererInstance().QueuePosition playerShape.Parent, playerShape.Name, newTop, newLeft
    SyncAllPlayerFrames
    On Error GoTo 0
End Sub

Public Sub AlignSprites(ByVal leftPos As Double, ByVal topPos As Double)
    On Error Resume Next
    
    m_Sprite.Left = leftPos
    m_Sprite.Top = topPos
    
    Dim wsGame As Worksheet
    Set wsGame = GetActiveWorksheet()
    If wsGame Is Nothing Then Exit Sub
    
    Dim frameName As Variant
    For Each frameName In AllConfiguredFrameNames()
        Dim frameShape As Shape
        Set frameShape = ShapeIfExists(wsGame, CStr(frameName))
        If Not frameShape Is Nothing Then
            BatchRendererInstance().QueuePosition wsGame, CStr(frameName), topPos, leftPos
        End If
    Next frameName

    UpdateLightOverlay m_Sprite.PlayerSprite
    
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

Private Sub SyncAllPlayerFrames()
    On Error Resume Next
    
    Dim wsGame As Worksheet
    Set wsGame = GetActiveWorksheet()
    If wsGame Is Nothing Then Exit Sub
    
    Dim frameName As Variant
    For Each frameName In AllConfiguredFrameNames()
        Dim frameShape As Shape
        Set frameShape = ShapeIfExists(wsGame, CStr(frameName))
        If Not frameShape Is Nothing Then
            BatchRendererInstance().QueuePosition wsGame, CStr(frameName), m_Sprite.Top, m_Sprite.Left
        End If
    Next frameName
    UpdateLightOverlay m_Sprite.PlayerSprite
    On Error GoTo 0
End Sub

Private Sub BindLightOverlay(ByVal ws As Worksheet)
    m_Sprite.LightShapeName = LIGHT_OVERLAY_NAME
    Set m_Sprite.LightShape = Nothing
    If ws Is Nothing Then Exit Sub

    Dim light As Shape
    Set light = ResolveLightShape(ws, LIGHT_OVERLAY_NAME)
    If light Is Nothing Then Exit Sub

    Set m_Sprite.LightShape = light
    m_Sprite.LightShapeName = light.Name
    light.Visible = True
    light.Rotation = 0
    UpdateLightOverlay m_Sprite.PlayerSprite
End Sub

Private Function EnsureLightOverlay() As Shape
    Dim light As Shape
    Set light = m_Sprite.LightShape

    On Error Resume Next
    If Not light Is Nothing Then
        Dim dummy As String
        dummy = light.Name
        If Err.Number <> 0 Then
            Err.Clear
            Set light = Nothing
        End If
    End If
    On Error GoTo 0

    If light Is Nothing Then
        Dim ws As Worksheet
        If Not m_Sprite.PlayerSprite Is Nothing Then
            On Error Resume Next
            Set ws = m_Sprite.PlayerSprite.Parent
            On Error GoTo 0
        End If

        If ws Is Nothing Then
            Dim gs As GameState
            Set gs = GameStateInstance
            If Not gs Is Nothing Then
                If gs.CurrentScreen <> "" Then
                    On Error Resume Next
                    Set ws = Sheets(gs.CurrentScreen)
                    On Error GoTo 0
                End If
            End If
        End If

        If Not ws Is Nothing Then
            Dim targetName As String
            targetName = m_Sprite.LightShapeName
            If Trim$(targetName) = "" Then targetName = LIGHT_OVERLAY_NAME

            Set light = ResolveLightShape(ws, targetName)
            If Not light Is Nothing Then
                Set m_Sprite.LightShape = light
                m_Sprite.LightShapeName = light.Name
            End If
        End If
    End If

    Set EnsureLightOverlay = light
End Function

Private Function ResolveLightShape(ByVal ws As Worksheet, ByVal nameHint As String) As Shape
    If ws Is Nothing Then Exit Function

    Dim targetName As String
    targetName = Trim$(nameHint)
    If targetName = "" Then targetName = LIGHT_OVERLAY_NAME

    Dim light As Shape
    Set light = ShapeIfExists(ws, targetName)
    If Not light Is Nothing Then
        Set ResolveLightShape = light
        Exit Function
    End If

    Dim candidate As Shape
    For Each candidate In ws.Shapes
        If StrComp(candidate.Name, targetName, vbTextCompare) = 0 Then
            Set ResolveLightShape = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Sub UpdateLightOverlay(Optional ByVal referenceShape As Shape)
    On Error Resume Next
    Dim spriteShape As Shape
    If referenceShape Is Nothing Then
        Set spriteShape = m_Sprite.PlayerSprite
    Else
        Set spriteShape = referenceShape
    End If
    If spriteShape Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If

    Dim light As Shape
    Set light = EnsureLightOverlay()
    If light Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If

    Dim centerTop As Double
    Dim centerLeft As Double
    centerTop = spriteShape.Top + (spriteShape.Height / 2#)
    centerLeft = spriteShape.Left + (spriteShape.Width / 2#)

    Dim lightTop As Double
    Dim lightLeft As Double
    lightTop = centerTop - (light.Height / 2#)
    lightLeft = centerLeft - (light.Width / 2#)
    
    BatchRendererInstance().QueuePosition light.Parent, light.Name, lightTop, lightLeft
    BatchRendererInstance().QueueVisibility light.Parent, light.Name, True
    light.Rotation = 0
    light.ZOrder msoBringToFront

    On Error GoTo 0
End Sub

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
        For Each frameName In AllConfiguredFrameNames()
            If CStr(frameName) <> activeName And CStr(frameName) <> previousName Then
                HideFrame ws, CStr(frameName)
            End If
        Next frameName
    End If
End Sub

