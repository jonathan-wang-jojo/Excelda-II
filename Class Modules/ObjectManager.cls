Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const MAX_OBJECT_SLOTS As Long = 30

'===================================================================================
'                              OBJECT MANAGER CLASS
'===================================================================================

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
    ' Object manager currently has no stateful initialization requirements.
End Sub

Public Sub Reset()
    ' Object manager currently has no persistent state to reset.
End Sub

Public Sub Destroy()
    ' Object manager currently has no resources to release.
End Sub

'===================================================================================
'                              OBJECT INTERACTIONS
'===================================================================================

Public Sub HandleSwordHit(ByVal swordShape As Shape, Optional ByVal objectPrefix As String = "Bush", Optional ByVal objectCount As Long = MAX_OBJECT_SLOTS)
    Dim ws As Worksheet
    Set ws = CurrentScreenSheet()
    If ws Is Nothing Then Exit Sub
    If swordShape Is Nothing Then Exit Sub
    If objectCount <= 0 Then Exit Sub

    Dim slotLimit As Long
    slotLimit = objectCount
    If slotLimit > MAX_OBJECT_SLOTS Then
        slotLimit = MAX_OBJECT_SLOTS
    End If

    Dim hits As Collection
    Set hits = New Collection

    Dim slotIndex As Long
    For slotIndex = 1 To slotLimit
        Dim targetName As String
        targetName = objectPrefix & CStr(slotIndex)
        Dim targetShape As Shape
        Set targetShape = ShapeIfExists(ws, targetName)

        If Not targetShape Is Nothing Then
            If targetShape.Visible And ShapesOverlap(swordShape, targetShape) Then
                hits.Add targetName
            End If
        End If
    Next slotIndex

    Dim nameVariant As Variant
    For Each nameVariant In hits
        ResetSingleObject ws, CStr(nameVariant)
    Next nameVariant

    ForceObjectRefresh
End Sub

Public Sub ResetObjects(Optional ByVal target As String = "All", Optional ByVal objectPrefix As String = "Bush", Optional ByVal objectCount As Long = MAX_OBJECT_SLOTS)
    Dim ws As Worksheet
    Set ws = CurrentScreenSheet()
    If ws Is Nothing Then Exit Sub

    If UCase$(target) = "ALL" Then
        If objectCount <= 0 Then Exit Sub

        Dim slotLimit As Long
        slotLimit = objectCount
        If slotLimit > MAX_OBJECT_SLOTS Then
            slotLimit = MAX_OBJECT_SLOTS
        End If

        Dim slotIndex As Long
        For slotIndex = 1 To slotLimit
            ResetSingleObject ws, objectPrefix & CStr(slotIndex)
        Next slotIndex
    Else
        ResetSingleObject ws, target
    End If

    ForceObjectRefresh
End Sub

Public Sub PositionObject(ByVal objectName As String, ByVal anchorAddress As String, Optional ByVal fillValue As Variant)
    Dim ws As Worksheet
    Set ws = CurrentScreenSheet()
    If ws Is Nothing Then Exit Sub

    Dim anchorCell As Range
    On Error Resume Next
    Set anchorCell = ws.Range(anchorAddress)
    On Error GoTo 0
    If anchorCell Is Nothing Then Exit Sub

    Dim objShape As Shape
    Set objShape = ShapeIfExists(ws, objectName)
    If objShape Is Nothing Then Exit Sub

    objShape.Top = anchorCell.Top
    objShape.Left = anchorCell.Left
    objShape.Visible = True

    If Not IsMissing(fillValue) Then
        FillObjectArea anchorCell, fillValue
    End If

    ForceObjectRefresh
End Sub

Public Sub PositionMultiple(ByVal objectType As String, ParamArray locations() As Variant)
    Dim ws As Worksheet
    Set ws = CurrentScreenSheet()
    If ws Is Nothing Then Exit Sub

    Dim slotIndex As Long
    For slotIndex = 1 To MAX_OBJECT_SLOTS
        Dim shapeName As String
        shapeName = objectType & CStr(slotIndex)

        Dim locationValue As String
        If slotIndex - 1 <= UBound(locations) Then
            locationValue = CStr(locations(slotIndex - 1))
        Else
            locationValue = ""
        End If

        AlignObject ws, shapeName, locationValue
    Next slotIndex

    ForceObjectRefresh
End Sub

'===================================================================================
'                              HEART PIECE EVENTS
'===================================================================================

Public Sub CollectHeartPiece(ByVal heartIndex As Long)
    Dim ws As Worksheet
    Set ws = CurrentScreenSheet()
    If ws Is Nothing Then Exit Sub

    Dim heartPiece As Shape
    Set heartPiece = ShapeIfExists(ws, "HeartPiece")
    If heartPiece Is Nothing Then Exit Sub

    Dim spriteMgr As SpriteManager
    Set spriteMgr = SpriteManagerInstance()
    If spriteMgr Is Nothing Then Exit Sub

    Dim linkSprite As Shape
    Set linkSprite = spriteMgr.PlayerSprite
    If linkSprite Is Nothing Then Exit Sub

    Dim linkWin As Shape
    Set linkWin = ShapeIfExists(ws, "LinkWin")
    If linkWin Is Nothing Then Exit Sub

    ClearObjectArea heartPiece
    heartPiece.Visible = False

    linkWin.Top = linkSprite.Top
    linkWin.Left = linkSprite.Left
    linkSprite.Visible = False
    linkWin.Visible = True

    heartPiece.Top = linkWin.Top - 45
    heartPiece.Left = linkWin.Left
    heartPiece.Visible = True

    TickFrame ws, 0
    Sleep 2000

    Dim dataSheet As Worksheet
    Set dataSheet = DataSheet()
    If dataSheet Is Nothing Then GoTo Cleanup

    dataSheet.Range("AB" & heartIndex).Value = "Y"
    dataSheet.Range("AB1").Value = dataSheet.Range("AB1").Value + 1

    If dataSheet.Range("AB1").Value = 4 Then
        ShowDialogue dataSheet.Range("R8").Value
        dataSheet.Range("AB1").Value = 0
        ' TODO: increment Link health here when system is ready
    Else
        ShowDialogue dataSheet.Range("R7").Value
    End If

Cleanup:
    heartPiece.Visible = False
    linkWin.Visible = False
    linkSprite.Visible = True
    ForceObjectRefresh
End Sub

'===================================================================================
'                              INTERNAL HELPERS
'===================================================================================

Private Sub ResetSingleObject(ByVal ws As Worksheet, ByVal shapeName As String)
    Dim objShape As Shape
    Set objShape = ShapeIfExists(ws, shapeName)
    If objShape Is Nothing Then Exit Sub

    ClearObjectArea objShape
    objShape.Visible = False
    objShape.Rotation = 0
End Sub

Private Sub ClearObjectArea(ByVal objShape As Shape)
    On Error Resume Next
    Dim anchorCell As Range
    Set anchorCell = objShape.TopLeftCell
    If anchorCell Is Nothing Then Exit Sub
    anchorCell.Resize(4, 6).Value = ""
    On Error GoTo 0
End Sub

Private Sub FillObjectArea(ByVal anchorCell As Range, ByVal value As Variant)
    On Error Resume Next
    anchorCell.Resize(4, 6).Value = value
    On Error GoTo 0
End Sub

Private Sub AlignObject(ByVal ws As Worksheet, ByVal shapeName As String, ByVal cellAddress As String)
    Dim objShape As Shape
    Set objShape = ShapeIfExists(ws, shapeName)
    If objShape Is Nothing Then Exit Sub

    If Len(Trim$(cellAddress)) = 0 Then
        objShape.Visible = False
        Exit Sub
    End If

    Dim anchorCell As Range
    On Error Resume Next
    Set anchorCell = ws.Range(cellAddress)
    On Error GoTo 0
    If anchorCell Is Nothing Then
        objShape.Visible = False
        Exit Sub
    End If

    objShape.Top = anchorCell.Top
    objShape.Left = anchorCell.Left
    objShape.Visible = True
End Sub

Private Function ShapesOverlap(ByVal a As Shape, ByVal b As Shape) As Boolean
    Dim aLeft As Double, aRight As Double, aTop As Double, aBottom As Double
    Dim bLeft As Double, bRight As Double, bTop As Double, bBottom As Double

    aLeft = a.Left
    aRight = a.Left + a.Width
    aTop = a.Top
    aBottom = a.Top + a.Height

    bLeft = b.Left
    bRight = b.Left + b.Width
    bTop = b.Top
    bBottom = b.Top + b.Height

    ShapesOverlap = (aLeft < bRight) And (aRight > bLeft) And _
                    (aTop < bBottom) And (aBottom > bTop)
End Function

Private Function CurrentScreenSheet() As Worksheet
    On Error Resume Next
    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Or gs.CurrentScreen = "" Then Exit Function
    Set CurrentScreenSheet = Sheets(gs.CurrentScreen)
    On Error GoTo 0
End Function

Private Function DataSheet() As Worksheet
    On Error Resume Next
    Set DataSheet = GameRegistryInstance().GetGameDataSheet()
    On Error GoTo 0
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

Private Sub TickFrame(ByVal ws As Worksheet, Optional ByVal delayMs As Long = 25)
    On Error Resume Next
    ws.Range("A1").Copy ws.Range("A2")
    On Error GoTo 0
    If delayMs > 0 Then Sleep delayMs
End Sub

Private Sub ForceObjectRefresh()
    Dim previousState As Boolean
    previousState = Application.ScreenUpdating
    Application.ScreenUpdating = True
    Application.ScreenUpdating = previousState
End Sub

Private Sub ShowDialogue(ByVal messageText As String)
    Dim dataSheet As Worksheet
    Set dataSheet = DataSheet()
    If dataSheet Is Nothing Then Exit Sub

    dataSheet.Range("C42").Value = messageText
    DialogueForm.Show
End Sub