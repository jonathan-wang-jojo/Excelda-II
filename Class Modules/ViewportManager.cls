Option Explicit

Private Const SCREEN_ROW_LABEL_COLUMN As Long = 7
Private Const SCREEN_ROW_OFFSET_COLUMN As Long = 8
Private Const SCREEN_COLUMN_LABEL_ROW As Long = 1
Private Const SCREEN_COLUMN_OFFSET_ROW As Long = 2
Private Const DEFAULT_VISIBLE_ROWS As Long = 33
Private Const DEFAULT_VISIBLE_COLUMNS As Long = 66
Private Const MINOTAUR_TARGET_ZOOM As Long = 140

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
    Dim spriteManager As SpriteManager
    Set spriteManager = SpriteManagerInstance()
    If spriteManager Is Nothing Then Exit Sub

    Dim linkSprite As Shape
    Set linkSprite = spriteManager.LinkSprite
    If linkSprite Is Nothing Then Exit Sub

    Dim linkCell As Range
    Set linkCell = linkSprite.TopLeftCell
    If linkCell Is Nothing Then Exit Sub

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

    Dim win As Object
    On Error Resume Next
    Set win = ActiveWindow
    On Error GoTo 0
    If win Is Nothing Then Exit Sub

    Dim previousUpdating As Boolean
    previousUpdating = Application.ScreenUpdating
    If previousUpdating Then Application.ScreenUpdating = False

    Dim visibleRows As Long
    Dim visibleColumns As Long
    Dim currentVisible As Range
    On Error Resume Next
    Set currentVisible = win.VisibleRange
    On Error GoTo 0
    If Not currentVisible Is Nothing Then
        visibleRows = currentVisible.Rows.Count
        visibleColumns = currentVisible.Columns.Count
    End If
    If visibleRows <= 0 Then visibleRows = DEFAULT_VISIBLE_ROWS
    If visibleColumns <= 0 Then visibleColumns = DEFAULT_VISIBLE_COLUMNS

    Dim sheetMaxTop As Long
    Dim sheetMaxLeft As Long
    sheetMaxTop = anchorCell.Worksheet.Rows.Count - visibleRows + 1
    sheetMaxLeft = anchorCell.Worksheet.Columns.Count - visibleColumns + 1
    If sheetMaxTop < 1 Then sheetMaxTop = 1
    If sheetMaxLeft < 1 Then sheetMaxLeft = 1

    Dim targetTop As Long
    Dim targetLeft As Long
    targetTop = ClampLong(anchorCell.Row, 1, sheetMaxTop)
    targetLeft = ClampLong(anchorCell.Column, 1, sheetMaxLeft)

    ApplyCameraBoundsForSheet anchorCell.Worksheet, visibleRows, visibleColumns, targetTop, targetLeft

    targetTop = ClampLong(targetTop, 1, sheetMaxTop)
    targetLeft = ClampLong(targetLeft, 1, sheetMaxLeft)

    On Error Resume Next
    If win.ScrollRow <> targetTop Then win.ScrollRow = targetTop
    If win.ScrollColumn <> targetLeft Then win.ScrollColumn = targetLeft
    On Error GoTo 0

    On Error Resume Next
    m_View.LastAnchor = anchorCell.Worksheet.Cells(targetTop, targetLeft).Address(False, False)
    On Error GoTo 0

    If previousUpdating Then Application.ScreenUpdating = True
End Sub

Public Sub MaintainLinkViewport(Optional ByVal rowMargin As Long = 6, Optional ByVal columnMargin As Long = 10)
    Dim spriteManager As SpriteManager
    Set spriteManager = SpriteManagerInstance()
    If spriteManager Is Nothing Then Exit Sub

    Dim linkShape As Shape
    Set linkShape = spriteManager.LinkSprite
    If linkShape Is Nothing Then Exit Sub

    Dim linkCell As Range
    Set linkCell = linkShape.TopLeftCell
    If linkCell Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = linkCell.Worksheet

    Dim win As Window
    On Error Resume Next
    Set win = ActiveWindow
    On Error GoTo 0
    If win Is Nothing Then Exit Sub

    If Not ws Is ActiveSheet Then
        ws.Activate
        On Error Resume Next
        Set win = ActiveWindow
        On Error GoTo 0
        If win Is Nothing Then Exit Sub
    End If

    If ws Is Sheet9 Then
        EnsureWindowZoom win, MINOTAUR_TARGET_ZOOM
    End If

    Dim visible As Range
    On Error Resume Next
    Set visible = win.VisibleRange
    On Error GoTo 0
    If visible Is Nothing Then Exit Sub

    Dim topRow As Long
    Dim leftColumn As Long
    Dim visibleRows As Long
    Dim visibleColumns As Long
    topRow = win.ScrollRow
    leftColumn = win.ScrollColumn
    visibleRows = visible.Rows.Count
    visibleColumns = visible.Columns.Count
    If visibleRows <= 0 Then visibleRows = DEFAULT_VISIBLE_ROWS
    If visibleColumns <= 0 Then visibleColumns = DEFAULT_VISIBLE_COLUMNS

    Dim bottomRow As Long
    Dim rightColumn As Long
    bottomRow = topRow + visibleRows - 1
    rightColumn = leftColumn + visibleColumns - 1

    Dim desiredTop As Long
    Dim desiredLeft As Long

    If ws Is Sheet9 Then
        Dim rowHeight As Double
        Dim columnWidth As Double
        rowHeight = linkCell.Height
        columnWidth = linkCell.Width
        If rowHeight <= 0# Then rowHeight = 1#
        If columnWidth <= 0# Then columnWidth = 1#

        Dim spriteHeightInRows As Double
        Dim spriteWidthInColumns As Double
        spriteHeightInRows = linkShape.Height / rowHeight
        spriteWidthInColumns = linkShape.Width / columnWidth
        If spriteHeightInRows <= 0# Then spriteHeightInRows = 1#
        If spriteWidthInColumns <= 0# Then spriteWidthInColumns = 1#

        Dim centerRowZero As Double
        Dim centerColumnZero As Double
        centerRowZero = (linkCell.Row - 1) + (linkShape.Top - linkCell.Top) / rowHeight + (spriteHeightInRows / 2#)
        centerColumnZero = (linkCell.Column - 1) + (linkShape.Left - linkCell.Left) / columnWidth + (spriteWidthInColumns / 2#)

        Dim desiredTopZero As Double
        Dim desiredLeftZero As Double
        desiredTopZero = centerRowZero - (visibleRows / 2#)
        desiredLeftZero = centerColumnZero - (visibleColumns / 2#)

        desiredTop = CLng(Int(desiredTopZero) + 1)
        desiredLeft = CLng(Int(desiredLeftZero) + 1)
    Else
        desiredTop = topRow
        desiredLeft = leftColumn

        If linkCell.Row < topRow + rowMargin Then
            desiredTop = linkCell.Row - rowMargin
        ElseIf linkCell.Row > bottomRow - rowMargin Then
            desiredTop = linkCell.Row + rowMargin - visibleRows + 1
        End If

        If linkCell.Column < leftColumn + columnMargin Then
            desiredLeft = linkCell.Column - columnMargin
        ElseIf linkCell.Column > rightColumn - columnMargin Then
            desiredLeft = linkCell.Column + columnMargin - visibleColumns + 1
        End If
    End If

    Dim maxTop As Long
    Dim maxLeft As Long
    maxTop = ws.Rows.Count - visibleRows + 1
    maxLeft = ws.Columns.Count - visibleColumns + 1
    If maxTop < 1 Then maxTop = 1
    If maxLeft < 1 Then maxLeft = 1

    desiredTop = ClampLong(desiredTop, 1, maxTop)
    desiredLeft = ClampLong(desiredLeft, 1, maxLeft)

    ApplyCameraBoundsForSheet ws, visibleRows, visibleColumns, desiredTop, desiredLeft

    desiredTop = ClampLong(desiredTop, 1, maxTop)
    desiredLeft = ClampLong(desiredLeft, 1, maxLeft)

    If desiredTop <> topRow Or desiredLeft <> leftColumn Then
        Dim previousUpdating As Boolean
        previousUpdating = Application.ScreenUpdating
        If previousUpdating Then Application.ScreenUpdating = False

        On Error Resume Next
        If desiredTop <> topRow Then win.ScrollRow = desiredTop
        If desiredLeft <> leftColumn Then win.ScrollColumn = desiredLeft
        On Error GoTo 0

        If previousUpdating Then Application.ScreenUpdating = True

        On Error Resume Next
        m_View.LastAnchor = ws.Cells(desiredTop, desiredLeft).Address(False, False)
        On Error GoTo 0
    End If
End Sub

Private Function ClampLong(ByVal value As Long, ByVal minValue As Long, ByVal maxValue As Long) As Long
    If maxValue < minValue Then
        ClampLong = minValue
        Exit Function
    End If
    If value < minValue Then value = minValue
    If value > maxValue Then value = maxValue
    ClampLong = value
End Function

Private Sub ApplyCameraBoundsForSheet(ByVal ws As Worksheet, ByVal visibleRows As Long, ByVal visibleColumns As Long, ByRef topRow As Long, ByRef leftColumn As Long)
    If ws Is Sheet9 Then
        Const MIN_COLUMN As Long = 101 ' CW
        Const MAX_COLUMN As Long = 287 ' KA
        Const MIN_ROW As Long = 17
        Const MAX_ROW As Long = 82

        Dim allowedMinLeft As Long
        Dim allowedMaxLeft As Long
        allowedMinLeft = MIN_COLUMN
        allowedMaxLeft = MAX_COLUMN - visibleColumns + 1
        If allowedMaxLeft < allowedMinLeft Then allowedMaxLeft = allowedMinLeft

        leftColumn = ClampLong(leftColumn, allowedMinLeft, allowedMaxLeft)

        Dim allowedMinTop As Long
        Dim allowedMaxTop As Long
        allowedMinTop = MIN_ROW
        allowedMaxTop = MAX_ROW - visibleRows + 1
        If allowedMaxTop < allowedMinTop Then allowedMaxTop = allowedMinTop

        topRow = ClampLong(topRow, allowedMinTop, allowedMaxTop)
    End If
End Sub

Private Function NormalizeScreenCode(ByVal screenCode As String) As String
    NormalizeScreenCode = UCase$(Trim$(screenCode))
End Function

Private Sub EnsureWindowZoom(ByVal win As Window, ByVal targetZoom As Long)
    If win Is Nothing Then Exit Sub
    If targetZoom <= 0 Then Exit Sub

    On Error Resume Next
    If win.Zoom <> targetZoom Then
        win.Zoom = targetZoom
    End If
    On Error GoTo 0
End Sub