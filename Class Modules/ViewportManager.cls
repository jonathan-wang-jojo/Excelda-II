Option Explicit

Private Const SCREEN_ROW_LABEL_COLUMN As Long = 7
Private Const SCREEN_ROW_OFFSET_COLUMN As Long = 8
Private Const SCREEN_COLUMN_LABEL_ROW As Long = 1
Private Const SCREEN_COLUMN_OFFSET_ROW As Long = 2
Private Const DEFAULT_VISIBLE_ROWS As Long = 33
Private Const DEFAULT_VISIBLE_COLUMNS As Long = 66

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
    m_View.VisibleRows = DEFAULT_VISIBLE_ROWS
    m_View.VisibleColumns = DEFAULT_VISIBLE_COLUMNS
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

    ScrollToAnchor anchor

    CaptureVisibleDimensions ws, targetRange
End Sub

Private Sub EnsureVisibleDimensions(ByVal ws As Worksheet)
    If m_View.VisibleRows <= 0 Then m_View.VisibleRows = DEFAULT_VISIBLE_ROWS
    If m_View.VisibleColumns <= 0 Then m_View.VisibleColumns = DEFAULT_VISIBLE_COLUMNS
    CaptureVisibleDimensions ws
    If m_View.VisibleRows <= 0 Then m_View.VisibleRows = DEFAULT_VISIBLE_ROWS
    If m_View.VisibleColumns <= 0 Then m_View.VisibleColumns = DEFAULT_VISIBLE_COLUMNS
End Sub

Private Sub CaptureVisibleDimensions(ByVal ws As Worksheet, Optional ByVal desiredViewport As Range)
    Dim activeWin As Object
    On Error Resume Next
    Set activeWin = ActiveWindow
    On Error GoTo 0
    If activeWin Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Activate
    On Error GoTo 0

    Dim visible As Range
    On Error Resume Next
    Set visible = activeWin.VisibleRange
    On Error GoTo 0

    If visible Is Nothing Then Exit Sub

    Dim actualRows As Long
    Dim actualColumns As Long
    actualRows = visible.Rows.Count
    actualColumns = visible.Columns.Count

    If actualRows <= 0 Then actualRows = DEFAULT_VISIBLE_ROWS
    If actualColumns <= 0 Then actualColumns = DEFAULT_VISIBLE_COLUMNS

    m_View.VisibleRows = actualRows
    m_View.VisibleColumns = actualColumns

    ReapplyViewportAnchor ws, desiredViewport
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

Private Sub ReapplyViewportAnchor(ByVal ws As Worksheet, ByVal desiredViewport As Range)
    Dim anchorCell As Range

    If Not desiredViewport Is Nothing Then
        Set anchorCell = desiredViewport.Cells(1, 1)
    ElseIf m_View.LastAnchor <> "" Then
        On Error Resume Next
        Set anchorCell = ws.Range(m_View.LastAnchor)
        On Error GoTo 0
    End If

    ScrollToAnchor anchorCell
End Sub

Private Sub ScrollToAnchor(ByVal anchorCell As Range)
    If anchorCell Is Nothing Then Exit Sub

    On Error Resume Next
    Application.Goto anchorCell, False
    On Error GoTo 0

    Dim win As Object
    On Error Resume Next
    Set win = ActiveWindow
    On Error GoTo 0
    If win Is Nothing Then Exit Sub

    On Error Resume Next
    win.ScrollRow = anchorCell.Row
    win.ScrollColumn = anchorCell.Column
    anchorCell.Select
    On Error GoTo 0
End Sub

Private Function NormalizeScreenCode(ByVal screenCode As String) As String
    NormalizeScreenCode = UCase$(Trim$(screenCode))
End Function