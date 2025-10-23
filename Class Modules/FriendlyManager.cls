Option Explicit

'===================================================================================
'                              FRIENDLY MANAGER CLASS
'===================================================================================

Private Const MAX_FRIENDLY_SLOTS As Long = 4

Private Type FriendlySlot
    BaseName As String
    DataRow As Long
    Direction As String
    Active As Boolean
    ScreenName As String
End Type

Private m_Friendlies(1 To MAX_FRIENDLY_SLOTS) As FriendlySlot

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

    Dim dir As Variant
    For Each dir In Array("D", "U", "L", "R")
        Dim shapeName As String
        shapeName = baseName & CStr(dir)
        Dim shp As Shape
        Set shp = ShapeIfExists(ws, shapeName)
        If Not shp Is Nothing Then
            shp.Top = anchorCell.Top
            shp.Left = anchorCell.Left
            shp.Rotation = 0
            shp.Visible = (CStr(dir) = visibleDir)
        End If
    Next dir
End Sub

Private Sub UpdateFriendlyVisibility(ByVal slotIndex As Long)
    If Not ValidSlot(slotIndex) Then Exit Sub
    If Not m_Friendlies(slotIndex).Active Then Exit Sub

    Dim ws As Worksheet
    Set ws = SheetIfExists(m_Friendlies(slotIndex).ScreenName)
    If ws Is Nothing Then Exit Sub

    Dim baseName As String
    baseName = m_Friendlies(slotIndex).BaseName

    Dim dir As Variant
    For Each dir In Array("D", "U", "L", "R")
        Dim shapeName As String
        shapeName = baseName & CStr(dir)
        Dim shp As Shape
        Set shp = ShapeIfExists(ws, shapeName)
        If Not shp Is Nothing Then
            shp.Visible = (CStr(dir) = m_Friendlies(slotIndex).Direction)
            shp.Rotation = 0
        End If
    Next dir
End Sub

Private Sub HideFriendlyShapes(ByVal baseName As String, ByVal screenName As String)
    Dim ws As Worksheet
    Set ws = SheetIfExists(screenName)
    If ws Is Nothing Then Exit Sub

    Dim dir As Variant
    For Each dir In Array("D", "U", "L", "R")
        Dim shapeName As String
        shapeName = baseName & CStr(dir)
        Dim shp As Shape
        Set shp = ShapeIfExists(ws, shapeName)
        If Not shp Is Nothing Then
            shp.Visible = False
            shp.Rotation = 0
        End If
    Next dir
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