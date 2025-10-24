Option Explicit

Private Const SCREEN_ROW_LABEL_COLUMN As Long = 7
Private Const SCREEN_ROW_OFFSET_COLUMN As Long = 8
Private Const SCREEN_COLUMN_LABEL_ROW As Long = 1
Private Const SCREEN_COLUMN_OFFSET_ROW As Long = 2
Private Const DEFAULT_VISIBLE_ROWS As Long = 22
Private Const DEFAULT_VISIBLE_COLUMNS As Long = 32

Private Type ViewportData
    VisibleRows As Long
    VisibleColumns As Long
    LastAnchor As String
End Type

Private m_View As ViewportData

'===================================================================================
'                              LIFECYCLE
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
    ResetState
End Sub

Public Sub Destroy()
    ResetState
End Sub

Private Sub ResetState()
    m_View.VisibleRows = 0
    m_View.VisibleColumns = 0
    m_View.LastAnchor = ""
End Sub

'===================================================================================
'                              PUBLIC API
'===================================================================================
Public Sub AlignToLink()
    On Error Resume Next
    Dim spriteManager As SpriteManager
    Set spriteManager = SpriteManagerInstance()
    If spriteManager Is Nothing Then Exit Sub

    Dim linkSprite As Shape
    Set linkSprite = spriteManager.LinkSprite
    If linkSprite Is Nothing Then Exit Sub

    Dim linkCell As Range
    Set linkCell = linkSprite.TopLeftCell
    If linkCell Is Nothing Then Exit Sub
    On Error GoTo 0

    ApplyViewportForCell linkCell
End Sub

Public Sub FocusOnScreen(ByVal screenCode As String)
    Dim code As String
    code = NormalizeScreenCode(screenCode)
    If code = "" Then Exit Sub

    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    Dim sheetName As String
    sheetName = gs.CurrentScreen
    If sheetName = "" Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = FindAnchorForScreen(ws, code)
    If anchor Is Nothing Then Exit Sub

    GoToViewport anchor
End Sub

Public Sub RefreshVisibleDimensions()
    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub
    If gs.CurrentScreen = "" Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets(gs.CurrentScreen)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    CaptureVisibleDimensions ws
End Sub

'===================================================================================
'                              CORE HELPERS
'===================================================================================
Private Sub ApplyViewportForCell(ByVal cell As Range)
    If cell Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = ComputeAnchorForCell(cell)
    If anchor Is Nothing Then Exit Sub

    GoToViewport anchor
End Sub

Private Function ComputeAnchorForCell(ByVal cell As Range) As Range
    If cell Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = cell.Worksheet

    EnsureVisibleDimensions ws

    Dim rowOffset As Long
    Dim colOffset As Long

    rowOffset = SafeOffset(ws.Cells(cell.Row, SCREEN_ROW_OFFSET_COLUMN).Value, m_View.VisibleRows \ 2)
    colOffset = SafeOffset(ws.Cells(SCREEN_COLUMN_OFFSET_ROW, cell.Column).Value, m_View.VisibleColumns \ 2)

    Dim topRow As Long
    Dim leftColumn As Long

    topRow = cell.Row - rowOffset + 1
    leftColumn = cell.Column - colOffset + 1

    If topRow < 1 Then topRow = 1
    If leftColumn < 1 Then leftColumn = 1

    Set ComputeAnchorForCell = ws.Cells(topRow, leftColumn)
End Function

Private Function FindAnchorForScreen(ByVal ws As Worksheet, ByVal screenCode As String) As Range
    Dim code As String
    code = NormalizeScreenCode(screenCode)
    If Len(code) <> 2 Then Exit Function

    Dim rowLabel As String
    Dim columnLabel As String
    rowLabel = Mid$(code, 1, 1)
    columnLabel = Mid$(code, 2, 1)

    Dim rowCell As Range
    Set rowCell = FindFirstMatch(ws.Columns(SCREEN_ROW_LABEL_COLUMN), rowLabel)

    Dim colCell As Range
    Set colCell = FindFirstMatch(ws.Rows(SCREEN_COLUMN_LABEL_ROW), columnLabel)

    If rowCell Is Nothing Or colCell Is Nothing Then Exit Function

    Dim candidate As Range
    Set candidate = ws.Cells(rowCell.Row, colCell.Column)

    Set FindAnchorForScreen = ComputeAnchorForCell(candidate)
End Function

Private Function FindFirstMatch(ByVal searchRange As Range, ByVal term As String) As Range
    If searchRange Is Nothing Then Exit Function

    Dim matchCell As Range
    On Error Resume Next
    Set matchCell = searchRange.Find(What:=term, LookIn:=xlValues, LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True)
    On Error GoTo 0

    If Not matchCell Is Nothing Then
        Set FindFirstMatch = matchCell
    End If
End Function

Private Sub GoToViewport(ByVal anchor As Range)
    If anchor Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = anchor.Worksheet

    EnsureVisibleDimensions ws

    Dim rowsVisible As Long
    Dim colsVisible As Long
    rowsVisible = m_View.VisibleRows
    colsVisible = m_View.VisibleColumns

    If rowsVisible <= 0 Then rowsVisible = DEFAULT_VISIBLE_ROWS
    If colsVisible <= 0 Then colsVisible = DEFAULT_VISIBLE_COLUMNS

    ws.Activate

    Dim maxRows As Long
    Dim maxCols As Long
    maxRows = ws.Rows.Count
    maxCols = ws.Columns.Count

    If anchor.Row + rowsVisible - 1 > maxRows Then
        rowsVisible = maxRows - anchor.Row + 1
    End If
    If anchor.Column + colsVisible - 1 > maxCols Then
        colsVisible = maxCols - anchor.Column + 1
    End If

    Dim targetRange As Range
    Set targetRange = anchor.Resize(rowsVisible, colsVisible)

    m_View.LastAnchor = anchor.Address(False, False)

    Dim previousZoom As Variant
    previousZoom = CaptureWindowZoom()

    On Error Resume Next
    Application.Goto targetRange, True
    On Error GoTo 0

    RestoreWindowZoom previousZoom

    CaptureVisibleDimensions ws
End Sub

Private Sub EnsureVisibleDimensions(ByVal ws As Worksheet)
    If m_View.VisibleRows > 0 And m_View.VisibleColumns > 0 Then Exit Sub
    CaptureVisibleDimensions ws
    If m_View.VisibleRows <= 0 Then m_View.VisibleRows = DEFAULT_VISIBLE_ROWS
    If m_View.VisibleColumns <= 0 Then m_View.VisibleColumns = DEFAULT_VISIBLE_COLUMNS
End Sub

Private Sub CaptureVisibleDimensions(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    Dim visible As Range
    Set visible = ActiveWindow.VisibleRange
    If Not visible Is Nothing Then
        m_View.VisibleRows = visible.Rows.Count
        m_View.VisibleColumns = visible.Columns.Count
    End If
    On Error GoTo 0
End Sub

Private Function SafeOffset(ByVal value As Variant, ByVal fallback As Long) As Long
    Dim parsed As Long
    On Error Resume Next
    parsed = CLng(Val(value))
    On Error GoTo 0
    If parsed <= 0 Then parsed = fallback
    If parsed <= 0 Then parsed = 1
    SafeOffset = parsed
End Function

Private Function CaptureWindowZoom() As Variant
    On Error Resume Next
    CaptureWindowZoom = ActiveWindow.Zoom
    If Err.Number <> 0 Then
        CaptureWindowZoom = Null
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub RestoreWindowZoom(ByVal zoomValue As Variant)
    If IsEmpty(zoomValue) Then Exit Sub
    If IsNull(zoomValue) Then Exit Sub
    If VarType(zoomValue) = vbBoolean Then Exit Sub
    If Not IsNumeric(zoomValue) Then Exit Sub

    On Error Resume Next
    ActiveWindow.Zoom = CDbl(zoomValue)
    On Error GoTo 0
End Sub

Private Function NormalizeScreenCode(ByVal screenCode As String) As String
    NormalizeScreenCode = UCase$(Trim$(screenCode))
End Function