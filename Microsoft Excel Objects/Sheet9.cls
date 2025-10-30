Option Explicit

' Code for Minotaur Game Sheet
' Sheet9 (Maze)

Public Sub btnStartGame_Click()
    ' Reset all managers and start new game
    Engine_GameLoop.StartNewGameOnSheet Me.Name, ""
End Sub

Public Sub btnContinue_Click()
    ' Continue game from current state
    Engine_GameLoop.ContinueGameOnSheet Me.Name
End Sub

'###################################################################################
'                       DEPRECATED - SPRITE CONFIGURATION LOGIC
'###################################################################################
' NOTE: The functions below (ApplyMinotaurSpriteConfig, ConfigureDirection, etc.)
'       are legacy sprite management code that should be moved to:
'       - GameConfigMinotaur.cls (game-specific config)
'       - SpriteManager.cls (sprite system logic)
'
' Keeping them here temporarily for backward compatibility, but they should NOT
' be called from buttons. Use Engine_GameLoop.ConfigureGameSheet() instead.
'###################################################################################

Public Sub ApplyMinotaurSpriteConfig()
    Dim sm As spriteManager
    Set sm = SpriteManagerInstance()

    Dim ws As Worksheet
    Set ws = Me

    Dim downIdle As String
    Dim downRun As Variant
    downIdle = ResolveDirectionalIdle(ws, Array("Front", "Down"))
    downRun = ResolveDirectionalRunFrames(ws, Array("Front", "Down"), downIdle)

    Dim upIdle As String
    Dim upRun As Variant
    upIdle = ResolveDirectionalIdle(ws, Array("Back", "Up"))
    upRun = ResolveDirectionalRunFrames(ws, Array("Back", "Up"), upIdle)

    Dim leftIdle As String
    Dim leftRun As Variant
    leftIdle = ResolveDirectionalIdle(ws, Array("Left"))
    leftRun = ResolveDirectionalRunFrames(ws, Array("Left"), leftIdle)

    Dim rightIdle As String
    Dim rightRun As Variant
    rightIdle = ResolveDirectionalIdle(ws, Array("Right"))
    rightRun = ResolveDirectionalRunFrames(ws, Array("Right"), rightIdle)

    Dim fallbackIdle As String
    fallbackIdle = FirstNonEmpty(Array(downIdle, upIdle, leftIdle, rightIdle))

    If fallbackIdle = "" Then
        Debug.Print "Sheet9: Unable to resolve minotaur idle frames; retaining existing configuration."
        Exit Sub
    End If

    If downIdle = "" Then downIdle = fallbackIdle
    If upIdle = "" Then upIdle = fallbackIdle
    If leftIdle = "" Then leftIdle = fallbackIdle
    If rightIdle = "" Then rightIdle = fallbackIdle

    Dim fallbackRun As Variant
    fallbackRun = FirstNonEmptyVariant(Array(downRun, upRun, leftRun, rightRun))

    If Not HasFrameEntries(downRun) Then downRun = fallbackRun
    If Not HasFrameEntries(upRun) Then upRun = fallbackRun
    If Not HasFrameEntries(leftRun) Then leftRun = fallbackRun
    If Not HasFrameEntries(rightRun) Then rightRun = fallbackRun

    sm.ResetPlayerFrameDefinitions

    ConfigureDirection sm, "Down", downIdle, downRun
    ConfigureDirection sm, "Up", upIdle, upRun
    ConfigureDirection sm, "Left", leftIdle, leftRun
    ConfigureDirection sm, "Right", rightIdle, rightRun

    Dim anchor As Range
    Set anchor = ws.Range(MINOTAUR_START_CELL)
    AlignFramesToAnchor sm, ws, anchor
End Sub

Private Sub ConfigureDirection(ByVal sm As spriteManager, ByVal Direction As String, ByVal idleName As String, ByVal runFrames As Variant)
    If idleName = "" Then Exit Sub
    If HasFrameEntries(runFrames) Then
        sm.ConfigurePlayerFrames Direction, idleName, runFrames
    Else
        sm.ConfigurePlayerFrames Direction, idleName
    End If
End Sub

Private Function ResolveDirectionalIdle(ByVal ws As Worksheet, ByVal prefixOptions As Variant) As String
    Dim prefix As Variant
    For Each prefix In prefixOptions
        Dim baseName As String
        baseName = "Minotaur" & Trim$(CStr(prefix))
        Dim resolved As String
        resolved = ResolveShapeName(ws, baseName & "Idle")
        If resolved <> "" Then
            ResolveDirectionalIdle = resolved
            Exit Function
        End If
    Next prefix

    Dim numbered As Variant
    numbered = CollectDirectionalNumberedFrames(ws, prefixOptions)
    If HasFrameEntries(numbered) Then
        ResolveDirectionalIdle = CStr(numbered(LBound(numbered)))
    End If
End Function

Private Function ResolveDirectionalRunFrames(ByVal ws As Worksheet, ByVal prefixOptions As Variant, ByVal idleName As String) As Variant
    Dim rawFrames As Variant
    rawFrames = CollectDirectionalNumberedFrames(ws, prefixOptions)
    If Not HasFrameEntries(rawFrames) Then
        ResolveDirectionalRunFrames = Empty
        Exit Function
    End If

    Dim filtered As Collection
    Set filtered = New Collection

    Dim frameName As Variant
    For Each frameName In rawFrames
        Dim candidate As String
        candidate = Trim$(CStr(frameName))
        If candidate <> "" Then
            If idleName = "" Or StrComp(candidate, idleName, vbTextCompare) <> 0 Then
                AppendUniqueFrame filtered, candidate
            End If
        End If
    Next frameName

    If filtered.Count = 0 Then
        ResolveDirectionalRunFrames = Empty
    Else
        Dim result() As String
        ReDim result(0 To filtered.Count - 1)
        Dim idx As Long
        For idx = 1 To filtered.Count
            result(idx - 1) = CStr(filtered(idx))
        Next idx
        ResolveDirectionalRunFrames = result
    End If
End Function

Private Function CollectDirectionalNumberedFrames(ByVal ws As Worksheet, ByVal prefixOptions As Variant) As Variant
    Const MAX_FRAMES As Long = 12
    Dim collected As Collection
    Set collected = New Collection

    Dim prefix As Variant
    For Each prefix In prefixOptions
        Dim baseName As String
        baseName = "Minotaur" & Trim$(CStr(prefix))
        Dim frameIndex As Long
        For frameIndex = 1 To MAX_FRAMES
            Dim resolved As String
            resolved = ResolveShapeName(ws, baseName & CStr(frameIndex))
            If resolved <> "" Then
                AppendUniqueFrame collected, resolved
            End If
        Next frameIndex
        If collected.Count > 0 Then Exit For
    Next prefix

    If collected.Count = 0 Then
        CollectDirectionalNumberedFrames = Empty
        Exit Function
    End If

    Dim result() As String
    ReDim result(0 To collected.Count - 1)
    Dim idx As Long
    For idx = 1 To collected.Count
        result(idx - 1) = CStr(collected(idx))
    Next idx

    CollectDirectionalNumberedFrames = result
End Function

Private Sub AppendUniqueFrame(ByVal frames As Collection, ByVal frameName As String)
    If frameName = "" Then Exit Sub
    Dim key As String
    key = UCase$(frameName)
    On Error Resume Next
    frames.Add frameName, key
    On Error GoTo 0
End Sub

Private Function FirstNonEmpty(ByVal candidates As Variant) As String
    Dim candidate As Variant
    For Each candidate In candidates
        Dim value As String
        value = Trim$(CStr(candidate))
        If value <> "" Then
            FirstNonEmpty = value
            Exit Function
        End If
    Next candidate
    FirstNonEmpty = ""
End Function

Private Function FirstNonEmptyVariant(ByVal candidates As Variant) As Variant
    Dim candidate As Variant
    For Each candidate In candidates
        If HasFrameEntries(candidate) Then
            FirstNonEmptyVariant = candidate
            Exit Function
        End If
    Next candidate
    FirstNonEmptyVariant = Empty
End Function

Private Function HasFrameEntries(ByVal frames As Variant) As Boolean
    If IsEmpty(frames) Then Exit Function
    If IsArray(frames) Then
        On Error Resume Next
        HasFrameEntries = (UBound(frames) >= LBound(frames))
        On Error GoTo 0
    ElseIf Trim$(CStr(frames)) <> "" Then
        HasFrameEntries = True
    End If
End Function

Private Function ShapeExists(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    ShapeExists = (ResolveShapeName(ws, shapeName) <> "")
End Function

Private Function ResolveShapeName(ByVal ws As Worksheet, ByVal nameCandidate As String) As String
    On Error Resume Next
    Dim candidate As Shape
    Set candidate = ws.Shapes(nameCandidate)
    On Error GoTo 0
    If Not candidate Is Nothing Then
        ResolveShapeName = candidate.Name
        Exit Function
    End If

    Dim shp As Shape
    Dim target As String
    target = UCase$(Trim$(nameCandidate))
    For Each shp In ws.Shapes
        If UCase$(Trim$(shp.Name)) = target Then
            ResolveShapeName = shp.Name
            Exit Function
        End If
    Next shp
End Function

Private Sub AlignFramesToAnchor(ByVal sm As spriteManager, ByVal ws As Worksheet, ByVal anchor As Range)
    If sm Is Nothing Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If anchor Is Nothing Then Exit Sub

    Dim frames As Variant
    frames = sm.GetConfiguredFrameNames()
    If IsEmpty(frames) Then Exit Sub

    Dim anchorTop As Double
    Dim anchorLeft As Double
    anchorTop = anchor.Top
    anchorLeft = anchor.Left

    Dim primaryResolved As String
    primaryResolved = ""

    Dim frameName As Variant
    For Each frameName In frames
        Dim shapeName As String
        shapeName = Trim$(CStr(frameName))
        If shapeName <> "" Then
            Dim resolvedName As String
            resolvedName = ResolveShapeName(ws, shapeName)
            If resolvedName <> "" Then
                Dim playerShape As Shape
                Set playerShape = ws.Shapes(resolvedName)
                playerShape.Top = anchorTop
                playerShape.Left = anchorLeft
                If primaryResolved = "" Then primaryResolved = resolvedName
                playerShape.visible = (StrComp(resolvedName, primaryResolved, vbTextCompare) = 0)
                If playerShape.visible Then
                    On Error Resume Next
                    playerShape.ZOrder msoBringToFront
                    On Error GoTo 0
                End If
            End If
        End If
    Next frameName
End Sub
