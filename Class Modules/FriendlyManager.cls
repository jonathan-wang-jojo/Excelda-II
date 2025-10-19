Option Explicit

Private Type FriendlySlot
    BaseName As String
    DataRow As Long
    Direction As String
    Active As Boolean
    ScreenName As String
End Type

Private m_Friendlies(1 To 4) As FriendlySlot

Private Sub Class_Initialize()
    InitializeState
End Sub

Public Sub Initialize()
    InitializeState
End Sub

Public Sub Reset()
    Dim i As Long
    For i = 1 To 4
        If m_Friendlies(i).Active Then
            DespawnFriendly i, m_Friendlies(i).DataRow
        Else
            ClearFriendly i
        End If
    Next i
End Sub

Public Sub Destroy()
    Reset
    InitializeState
End Sub

Private Sub InitializeState()
    Dim i As Long
    For i = 1 To 4
        ClearFriendly i
    Next i
End Sub

Public Sub SpawnFriendly(ByVal slot As Long, ByVal baseName As String, ByVal dataRow As Long, ByVal anchorCell As Range, Optional ByVal direction As String = "D")
    If Not ValidSlot(slot) Then Exit Sub
    If baseName = "" Then Exit Sub
    If anchorCell Is Nothing Then Exit Sub

    Dim normalizedDir As String
    normalizedDir = NormalizeDirection(direction)

    Dim screenName As String
    screenName = anchorCell.Worksheet.Name

    RegisterFriendly slot, baseName, dataRow, normalizedDir, screenName
    AlignFriendlyShapes baseName, anchorCell, normalizedDir
    Sheets(SHEET_DATA).Range("C" & dataRow).Value = "Y"
    ForceFriendlyRefresh
End Sub

Public Sub DespawnFriendly(ByVal slot As Long, ByVal dataRow As Long)
    If Not ValidSlot(slot) Then Exit Sub

    Dim baseName As String
    Dim screenName As String
    baseName = m_Friendlies(slot).BaseName
    screenName = m_Friendlies(slot).ScreenName

    If baseName <> "" Then
        HideFriendlyShapes baseName, screenName
    End If

    If dataRow > 0 Then
        Sheets(SHEET_DATA).Range("C" & dataRow).Value = "N"
    End If

    ClearFriendly slot
    ForceFriendlyRefresh
End Sub

Public Function FriendlyBaseName(ByVal slot As Long) As String
    If ValidSlot(slot) Then
        FriendlyBaseName = m_Friendlies(slot).BaseName
    End If
End Function

Public Function FriendlyDirection(ByVal slot As Long) As String
    If ValidSlot(slot) Then
        FriendlyDirection = m_Friendlies(slot).Direction
    End If
End Function

Public Function FriendlyDataRow(ByVal slot As Long) As Long
    If ValidSlot(slot) Then
        FriendlyDataRow = m_Friendlies(slot).DataRow
    End If
End Function

Public Function FriendlyIsActive(ByVal slot As Long) As Boolean
    If ValidSlot(slot) Then
        FriendlyIsActive = m_Friendlies(slot).Active
    End If
End Function

Public Sub SetFriendlyDirection(ByVal slot As Long, ByVal direction As String)
    If Not ValidSlot(slot) Then Exit Sub
    If Not m_Friendlies(slot).Active Then Exit Sub

    Dim normalizedDir As String
    normalizedDir = NormalizeDirection(direction)

    If normalizedDir = m_Friendlies(slot).Direction Then Exit Sub

    m_Friendlies(slot).Direction = normalizedDir
    UpdateFriendlyVisibility slot
    ForceFriendlyRefresh
End Sub

Public Function GetFriendlyShape(ByVal slot As Long) As Shape
    If Not ValidSlot(slot) Then Exit Function
    If Not m_Friendlies(slot).Active Then Exit Function

    Dim ws As Worksheet
    Set ws = SheetIfExists(m_Friendlies(slot).ScreenName)
    If ws Is Nothing Then Exit Function

    Dim shapeName As String
    shapeName = m_Friendlies(slot).BaseName & m_Friendlies(slot).Direction
    Set GetFriendlyShape = ShapeIfExists(ws, shapeName)
End Function

Public Sub ForceFriendlyRefresh()
    Dim previousState As Boolean
    previousState = Application.ScreenUpdating
    Application.ScreenUpdating = True
    Application.ScreenUpdating = previousState
End Sub

Private Sub RegisterFriendly(ByVal slot As Long, ByVal baseName As String, ByVal dataRow As Long, ByVal direction As String, ByVal screenName As String)
    If Not ValidSlot(slot) Then Exit Sub

    With m_Friendlies(slot)
        .BaseName = baseName
        .DataRow = dataRow
        .Direction = direction
        .ScreenName = screenName
        .Active = True
    End With
End Sub

Private Sub ClearFriendly(ByVal slot As Long)
    If Not ValidSlot(slot) Then Exit Sub

    With m_Friendlies(slot)
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

Private Sub UpdateFriendlyVisibility(ByVal slot As Long)
    If Not ValidSlot(slot) Then Exit Sub
    If Not m_Friendlies(slot).Active Then Exit Sub

    Dim ws As Worksheet
    Set ws = SheetIfExists(m_Friendlies(slot).ScreenName)
    If ws Is Nothing Then Exit Sub

    Dim baseName As String
    baseName = m_Friendlies(slot).BaseName

    Dim dir As Variant
    For Each dir In Array("D", "U", "L", "R")
        Dim shapeName As String
        shapeName = baseName & CStr(dir)
        Dim shp As Shape
        Set shp = ShapeIfExists(ws, shapeName)
        If Not shp Is Nothing Then
            shp.Visible = (CStr(dir) = m_Friendlies(slot).Direction)
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

Private Function ValidSlot(ByVal slot As Long) As Boolean
    ValidSlot = (slot >= 1 And slot <= 4)
End Function