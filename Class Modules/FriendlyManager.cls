Option Explicit

'===================================================================================
'                              FRIENDLY MANAGER CLASS
'===================================================================================

Private Const MAX_FRIENDLY_SLOTS As Long = 4
Private Const DEFAULT_LOOK_INTERVAL_MIN As Double = 1.5
Private Const DEFAULT_LOOK_INTERVAL_MAX As Double = 3.5

Private Type FriendlySlot
    BaseName As String
    DataRow As Long
    Direction As String
    Active As Boolean
    ScreenName As String
    DirectionTimer As Double
    DirectionPool As String
End Type

Private m_Friendlies(1 To MAX_FRIENDLY_SLOTS) As FriendlySlot
Private m_RandomSeedInitialized As Boolean

'===================================================================================
'                              INITIALIZATION & LIFECYCLE
'===================================================================================

Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Initialize()
    EnsureRandomSeed
    ResetState
End Sub

Public Sub Reset()
    Dim slotIndex As Long
    For slotIndex = 1 To MAX_FRIENDLY_SLOTS
        If m_Friendlies(slotIndex).Active Then
            DespawnFriendly slotIndex, m_Friendlies(slotIndex).DataRow
        Else
            ClearFriendly slotIndex
        End If
    Next slotIndex
    ResetState
End Sub

Public Sub Destroy()
    Reset
End Sub

Private Sub ResetState()
    Dim slotIndex As Long
    For slotIndex = 1 To MAX_FRIENDLY_SLOTS
        ClearFriendly slotIndex
    Next slotIndex
End Sub

'===================================================================================
'                              SPAWNING & CLEANUP
'===================================================================================

Public Sub SpawnFriendly(ByVal slotIndex As Long, ByVal baseName As String, ByVal dataRow As Long, ByVal anchorCell As Range, Optional ByVal direction As String = "D")
    If Not ValidSlot(slotIndex) Then Exit Sub
    If baseName = "" Then Exit Sub
    If anchorCell Is Nothing Then Exit Sub

    Dim normalizedDir As String
    normalizedDir = NormalizeDirection(direction)

    Dim screenName As String
    screenName = anchorCell.Worksheet.Name

    RegisterFriendly slotIndex, baseName, dataRow, normalizedDir, screenName
    AlignFriendlyShapes baseName, anchorCell, normalizedDir
    Sheets(SHEET_DATA).Range("C" & dataRow).Value = "Y"
    ForceFriendlyRefresh
End Sub

Public Sub DespawnFriendly(ByVal slotIndex As Long, ByVal dataRow As Long)
    If Not ValidSlot(slotIndex) Then Exit Sub

    Dim baseName As String
    Dim screenName As String
    baseName = m_Friendlies(slotIndex).BaseName
    screenName = m_Friendlies(slotIndex).ScreenName

    If baseName <> "" Then
        HideFriendlyShapes baseName, screenName
    End If

    If dataRow > 0 Then
        Sheets(SHEET_DATA).Range("C" & dataRow).Value = "N"
    End If

    ClearFriendly slotIndex
    ForceFriendlyRefresh
End Sub

'===================================================================================
'                              ACCESSORS & MUTATORS
'===================================================================================

Public Function FriendlyBaseName(ByVal slotIndex As Long) As String
    If ValidSlot(slotIndex) Then
        FriendlyBaseName = m_Friendlies(slotIndex).BaseName
    End If
End Function

Public Function FriendlyDirection(ByVal slotIndex As Long) As String
    If ValidSlot(slotIndex) Then
        FriendlyDirection = m_Friendlies(slotIndex).Direction
    End If
End Function

Public Function FriendlyDataRow(ByVal slotIndex As Long) As Long
    If ValidSlot(slotIndex) Then
        FriendlyDataRow = m_Friendlies(slotIndex).DataRow
    End If
End Function

Public Function FriendlyIsActive(ByVal slotIndex As Long) As Boolean
    If ValidSlot(slotIndex) Then
        FriendlyIsActive = m_Friendlies(slotIndex).Active
    End If
End Function

Public Sub SetFriendlyDirection(ByVal slotIndex As Long, ByVal direction As String)
    If Not ValidSlot(slotIndex) Then Exit Sub
    If Not m_Friendlies(slotIndex).Active Then Exit Sub

    Dim normalizedDir As String
    normalizedDir = NormalizeDirection(direction)

    If normalizedDir = m_Friendlies(slotIndex).Direction Then Exit Sub

    m_Friendlies(slotIndex).Direction = normalizedDir
    UpdateFriendlyVisibility slotIndex
    ForceFriendlyRefresh
End Sub

Public Sub Tick(ByVal deltaSeconds As Double)
    Dim slotIndex As Long
    For slotIndex = 1 To MAX_FRIENDLY_SLOTS
        UpdateFriendlyTimer slotIndex, deltaSeconds
    Next slotIndex
End Sub

Public Function GetFriendlyShape(ByVal slotIndex As Long) As Shape
    If Not ValidSlot(slotIndex) Then Exit Function
    If Not m_Friendlies(slotIndex).Active Then Exit Function

    Dim ws As Worksheet
    Set ws = SheetIfExists(m_Friendlies(slotIndex).ScreenName)
    If ws Is Nothing Then Exit Function

    Dim shapeName As String
    shapeName = m_Friendlies(slotIndex).BaseName & m_Friendlies(slotIndex).Direction
    Set GetFriendlyShape = ShapeIfExists(ws, shapeName)
End Function

Public Sub ForceFriendlyRefresh()
    Dim previousState As Boolean
    previousState = Application.ScreenUpdating
    Application.ScreenUpdating = True
    Application.ScreenUpdating = previousState
End Sub

'===================================================================================
'                              INTERNAL HELPERS
'===================================================================================

Private Sub RegisterFriendly(ByVal slotIndex As Long, ByVal baseName As String, ByVal dataRow As Long, ByVal direction As String, ByVal screenName As String)
    If Not ValidSlot(slotIndex) Then Exit Sub

    With m_Friendlies(slotIndex)
        .BaseName = baseName
        .DataRow = dataRow
        .Direction = direction
        .ScreenName = screenName
        .Active = True
        .DirectionPool = ResolveDirectionPool(baseName)
        If .DirectionPool <> "" Then
            .DirectionTimer = NextDirectionInterval(baseName, .DirectionPool)
        Else
            .DirectionTimer = 0#
        End If
    End With
End Sub

Private Sub ClearFriendly(ByVal slotIndex As Long)
    If Not ValidSlot(slotIndex) Then Exit Sub

    With m_Friendlies(slotIndex)
        .BaseName = ""
        .DataRow = 0
        .Direction = "D"
        .ScreenName = ""
        .Active = False
        .DirectionTimer = 0#
        .DirectionPool = ""
    End With
End Sub

Private Function NormalizeDirection(ByVal rawDirection As String) As String
    Dim dir As String
    dir = UCase$(Trim$(rawDirection))
    Select Case dir
        Case "U", "D", "L", "R"
            NormalizeDirection = dir
        Case Else
            NormalizeDirection = "D"
    End Select
End Function

Private Sub AlignFriendlyShapes(ByVal baseName As String, ByVal anchorCell As Range, ByVal visibleDir As String)
    Dim ws As Worksheet
    Set ws = anchorCell.Worksheet

    Dim dirCode As Variant
    For Each dirCode In FriendlyDirectionCodes()
        Dim suffixes As Variant
        suffixes = FriendlyDirectionSuffixes(CStr(dirCode))
        Dim suffix As Variant
        For Each suffix In suffixes
            Dim shapeName As String
            shapeName = baseName & CStr(suffix)
            Dim shp As Shape
            Set shp = ShapeIfExists(ws, shapeName)
            If Not shp Is Nothing Then
                shp.Top = anchorCell.Top
                shp.Left = anchorCell.Left
                shp.Rotation = 0
            End If
        Next suffix
    Next dirCode

    ApplyFriendlyVisibility ws, baseName, visibleDir
End Sub

Private Sub UpdateFriendlyVisibility(ByVal slotIndex As Long)
    If Not ValidSlot(slotIndex) Then Exit Sub
    If Not m_Friendlies(slotIndex).Active Then Exit Sub

    Dim ws As Worksheet
    Set ws = SheetIfExists(m_Friendlies(slotIndex).ScreenName)
    If ws Is Nothing Then Exit Sub

    Dim baseName As String
    baseName = m_Friendlies(slotIndex).BaseName

    ApplyFriendlyVisibility ws, baseName, m_Friendlies(slotIndex).Direction
End Sub

Private Sub HideFriendlyShapes(ByVal baseName As String, ByVal screenName As String)
    Dim ws As Worksheet
    Set ws = SheetIfExists(screenName)
    If ws Is Nothing Then Exit Sub

    ApplyFriendlyVisibility ws, baseName, ""
End Sub

Private Function FriendlyDirectionCodes() As Variant
    FriendlyDirectionCodes = Array("D", "U", "L", "R")
End Function

Private Function FriendlyDirectionSuffixes(ByVal direction As String) As Variant
    Dim normalized As String
    normalized = UCase$(Trim$(direction))

    Select Case normalized
        Case "U"
            FriendlyDirectionSuffixes = Array("U", "Up", "Back")
        Case "D"
            FriendlyDirectionSuffixes = Array("D", "Down", "Front")
        Case "L"
            FriendlyDirectionSuffixes = Array("L", "Left")
        Case "R"
            FriendlyDirectionSuffixes = Array("R", "Right")
        Case Else
            FriendlyDirectionSuffixes = Array(normalized)
    End Select
End Function

Private Sub ApplyFriendlyVisibility(ByVal ws As Worksheet, ByVal baseName As String, ByVal visibleDirection As String)
    Dim activeShape As Shape
    If Trim$(visibleDirection) <> "" Then
        Set activeShape = ResolveFriendlyShape(ws, baseName, visibleDirection)
        If activeShape Is Nothing Then
            Set activeShape = ResolveFriendlyShape(ws, baseName, "D")
        End If
    End If

    Dim dirCode As Variant
    For Each dirCode In FriendlyDirectionCodes()
        Dim suffixes As Variant
        suffixes = FriendlyDirectionSuffixes(CStr(dirCode))
        Dim suffix As Variant
        For Each suffix In suffixes
            Dim shapeName As String
            shapeName = baseName & CStr(suffix)
            Dim shp As Shape
            Set shp = ShapeIfExists(ws, shapeName)
            If Not shp Is Nothing Then
                Dim shouldShow As Boolean
                shouldShow = (Not activeShape Is Nothing And StrComp(shp.Name, activeShape.Name, vbTextCompare) = 0)
                shp.Visible = shouldShow
                shp.Rotation = 0
                If shouldShow Then
                    On Error Resume Next
                    shp.ZOrder msoBringToFront
                    On Error GoTo 0
                End If
            End If
        Next suffix
    Next dirCode
End Sub

Private Function ResolveFriendlyShape(ByVal ws As Worksheet, ByVal baseName As String, ByVal direction As String) As Shape
    Dim suffixes As Variant
    suffixes = FriendlyDirectionSuffixes(direction)

    Dim suffix As Variant
    For Each suffix In suffixes
        Dim shapeName As String
        shapeName = baseName & CStr(suffix)
        Dim shp As Shape
        Set shp = ShapeIfExists(ws, shapeName)
        If Not shp Is Nothing Then
            Set ResolveFriendlyShape = shp
            Exit Function
        End If
    Next suffix
End Function

Private Sub UpdateFriendlyTimer(ByVal slotIndex As Long, ByVal deltaSeconds As Double)
    If Not ValidSlot(slotIndex) Then Exit Sub
    If deltaSeconds <= 0# Then Exit Sub

    With m_Friendlies(slotIndex)
        If Not .Active Then Exit Sub
        If Trim$(.DirectionPool) = "" Then Exit Sub

        .DirectionTimer = .DirectionTimer - deltaSeconds
        If .DirectionTimer <= 0# Then
            Dim nextDir As String
            nextDir = ChooseNextDirection(.DirectionPool, .Direction)
            If nextDir <> "" Then
                SetFriendlyDirection slotIndex, nextDir
            End If
            .DirectionTimer = NextDirectionInterval(.BaseName, .DirectionPool)
            If .DirectionTimer <= 0# Then
                .DirectionTimer = DEFAULT_LOOK_INTERVAL_MAX
            End If
        End If
    End With
End Sub

Private Function ResolveDirectionPool(ByVal baseName As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(baseName))

    Select Case normalized
        Case "GUY"
            ResolveDirectionPool = "D,R"
        Case Else
            ResolveDirectionPool = ""
    End Select
End Function

Private Function NextDirectionInterval(ByVal baseName As String, ByVal directionPool As String) As Double
    If Trim$(directionPool) = "" Then Exit Function

    Dim minValue As Double
    Dim maxValue As Double

    Select Case UCase$(Trim$(baseName))
        Case "GUY"
            minValue = DEFAULT_LOOK_INTERVAL_MIN
            maxValue = DEFAULT_LOOK_INTERVAL_MAX
        Case Else
            minValue = DEFAULT_LOOK_INTERVAL_MIN
            maxValue = DEFAULT_LOOK_INTERVAL_MAX
    End Select

    NextDirectionInterval = RandomInterval(minValue, maxValue)
End Function

Private Function RandomInterval(ByVal minValue As Double, ByVal maxValue As Double) As Double
    If maxValue <= minValue Then
        RandomInterval = minValue
    Else
        RandomInterval = minValue + (maxValue - minValue) * Rnd()
    End If
End Function

Private Function ChooseNextDirection(ByVal pool As String, ByVal currentDirection As String) As String
    Dim candidates As Variant
    candidates = ParseDirectionPool(pool)
    If IsEmpty(candidates) Then Exit Function

    Dim lower As Long
    Dim upper As Long
    lower = LBound(candidates)
    upper = UBound(candidates)
    If upper < lower Then Exit Function

    Dim count As Long
    count = upper - lower + 1
    Dim startIndex As Long
    startIndex = lower + CLng(Int(count * Rnd()))
    Dim index As Long
    index = startIndex

    Dim attempts As Long
    For attempts = 1 To count
        Dim candidate As String
        candidate = CStr(candidates(index))
        If candidate <> "" Then
            If count = 1 Then
                ChooseNextDirection = candidate
                Exit Function
            ElseIf StrComp(candidate, currentDirection, vbTextCompare) <> 0 Then
                ChooseNextDirection = candidate
                Exit Function
            End If
        End If
        index = index + 1
        If index > upper Then index = lower
    Next attempts

    ChooseNextDirection = CStr(candidates(startIndex))
End Function

Private Function ParseDirectionPool(ByVal pool As String) As Variant
    Dim cleaned As String
    cleaned = Replace(pool, ";", ",")

    Dim tokens As Variant
    tokens = Split(cleaned, ",")

    Dim collected As Collection
    Set collected = New Collection

    Dim token As Variant
    For Each token In tokens
        Dim normalized As String
        normalized = NormalizeDirectionToken(CStr(token))
        If normalized <> "" Then
            On Error Resume Next
            collected.Add normalized, normalized
            On Error GoTo 0
        End If
    Next token

    If collected.Count = 0 Then
        ParseDirectionPool = Empty
    Else
        Dim result() As String
        ReDim result(0 To collected.Count - 1)
        Dim idx As Long
        For idx = 1 To collected.Count
            result(idx - 1) = collected(idx)
        Next idx
        ParseDirectionPool = result
    End If
End Function

Private Function NormalizeDirectionToken(ByVal token As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(token))
    Select Case normalized
        Case "U", "D", "L", "R"
            NormalizeDirectionToken = normalized
    End Select
End Function

Private Sub EnsureRandomSeed()
    If m_RandomSeedInitialized Then Exit Sub
    Randomize
    m_RandomSeedInitialized = True
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

Private Function SheetIfExists(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set SheetIfExists = Sheets(sheetName)
    If Err.Number <> 0 Then
        Err.Clear
        Set SheetIfExists = Nothing
    End If
    On Error GoTo 0
End Function

Private Function ValidSlot(ByVal slotIndex As Long) As Boolean
    ValidSlot = (slotIndex >= 1 And slotIndex <= MAX_FRIENDLY_SLOTS)
End Function