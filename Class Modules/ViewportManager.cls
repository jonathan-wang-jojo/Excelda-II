Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' VIEWPORT MANAGER - Camera & Scrolling System
'═══════════════════════════════════════════════════════════════════════════════
' Handles viewport positioning, camera following, and screen navigation.
' Uses instant camera snap for jitter-free scrolling (sprite interpolation
' provides visual smoothness via game loop alpha).
'═══════════════════════════════════════════════════════════════════════════════

'──────────────────────────────────────────────────────────────────────────────
' Screen Grid Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Const SCREEN_ROW_LABEL_COLUMN As Long = 7      ' Column containing row labels (A, B, C...)
Private Const SCREEN_ROW_OFFSET_COLUMN As Long = 8     ' Column containing row offsets
Private Const SCREEN_COLUMN_LABEL_ROW As Long = 1      ' Row containing column labels (1, 2, 3...)
Private Const SCREEN_COLUMN_OFFSET_ROW As Long = 2     ' Row containing column offsets

'──────────────────────────────────────────────────────────────────────────────
' Viewport Defaults
'──────────────────────────────────────────────────────────────────────────────
Private Const DEFAULT_VISIBLE_ROWS As Long = 33
Private Const DEFAULT_VISIBLE_COLUMNS As Long = 66
Private Const DEFAULT_ROW_MARGIN As Long = 6            ' Margin-based camera mode
Private Const DEFAULT_COLUMN_MARGIN As Long = 10

'──────────────────────────────────────────────────────────────────────────────
' Camera Modes
'──────────────────────────────────────────────────────────────────────────────
Private Enum CameraMode
    CameraInstant = 0       ' Center on player instantly (default, Unity-style)
    CameraMargins = 1       ' Only scroll when player near edges (platformer-style)
    CameraGrid = 2          ' Snap to grid quadrants (Zelda-style)
End Enum

'──────────────────────────────────────────────────────────────────────────────
' Viewport State
'──────────────────────────────────────────────────────────────────────────────
Private Type ViewportData
    VisibleRows As Long
    VisibleColumns As Long
    LastAnchor As String
    Mode As CameraMode
    ' Configurable bounds (0 = no restriction)
    BoundsMinRow As Long
    BoundsMaxRow As Long
    BoundsMinColumn As Long
    BoundsMaxColumn As Long
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
    m_View.Mode = CameraInstant
    m_View.BoundsMinRow = 0
    m_View.BoundsMaxRow = 0
    m_View.BoundsMinColumn = 0
    m_View.BoundsMaxColumn = 0
End Sub

'===================================================================================
'                              PUBLIC API - Configuration
'===================================================================================
Public Sub SetCameraMode(ByVal mode As Long)
    If mode < 0 Or mode > 2 Then Exit Sub
    m_View.Mode = mode
End Sub

Public Sub SetCameraBounds(ByVal minRow As Long, ByVal maxRow As Long, _
                          ByVal minColumn As Long, ByVal maxColumn As Long)
    m_View.BoundsMinRow = minRow
    m_View.BoundsMaxRow = maxRow
    m_View.BoundsMinColumn = minColumn
    m_View.BoundsMaxColumn = maxColumn
End Sub

Public Sub ClearCameraBounds()
    m_View.BoundsMinRow = 0
    m_View.BoundsMaxRow = 0
    m_View.BoundsMinColumn = 0
    m_View.BoundsMaxColumn = 0
End Sub

'===================================================================================
'                              PUBLIC API - Navigation
'===================================================================================
Public Sub AlignToPlayer()
    Dim spriteManager As SpriteManager
    Set spriteManager = SpriteManagerInstance()
    If spriteManager Is Nothing Then Exit Sub

    Dim playerSprite As Shape
    Set playerSprite = spriteManager.PlayerSprite
    If playerSprite Is Nothing Then Exit Sub

    Dim playerCell As Range
    Set playerCell = playerSprite.TopLeftCell
    If playerCell Is Nothing Then Exit Sub

    ApplyViewportForCell playerCell
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

    ' Apply configured bounds if set
    If m_View.BoundsMinRow > 0 Or m_View.BoundsMaxRow > 0 Or _
       m_View.BoundsMinColumn > 0 Or m_View.BoundsMaxColumn > 0 Then
        ApplyConfiguredBounds visibleRows, visibleColumns, targetTop, targetLeft
    End If

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

'===================================================================================
'                              PUBLIC API - Scrolling
'===================================================================================
Public Sub HandleScrollTransition(ByVal scrollCode As String)
    On Error GoTo ErrorHandler

    Dim normalizedCode As String
    normalizedCode = Trim$(scrollCode)
    If normalizedCode = "" Then Exit Sub

    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    Dim primaryDir As String
    primaryDir = ResolveScrollDirection(normalizedCode, gs.MoveDir, gs.LastDir)
    If primaryDir = "" Then Exit Sub

    If ShouldPreventRescroll(gs.PlayerCellAddress, primaryDir) Then Exit Sub

    PerformWindowScroll normalizedCode, primaryDir

    DataCacheInstance().SetValue RANGE_PREVIOUS_CELL, gs.PlayerCellAddress
    DataCacheInstance().SetValue RANGE_PREVIOUS_SCROLL, primaryDir
    DataCacheInstance().SetValue RANGE_SCROLL_DIRECTION, primaryDir

    UpdateScreenLocation normalizedCode, primaryDir

    Dim setupMacro As String
    setupMacro = gs.CurrentScreenCode
    If setupMacro = "" Then setupMacro = gs.CurrentScreen
    If setupMacro <> "" Then
        On Error GoTo ScreenSetupError
        SceneManagerInstance().ApplyScreen setupMacro
        DataCacheInstance().Flush GameRegistryInstance().GetGameDataSheet()
        On Error GoTo ErrorHandler
    End If

    Exit Sub

ScreenSetupError:
    MsgBox "Screen setup macro not found: " & setupMacro, vbCritical, "Screen Setup Error"
    On Error GoTo ErrorHandler
    Exit Sub

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Scroll Error: " & Err.Description, vbCritical, "Scroll Error"
    End If
End Sub

Public Sub UpdateScreenLocation(ByVal scrollDir As String, ByVal direction As String)
    On Error GoTo ErrorHandler

    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    If gs.CurrentScreen = "" Or gs.PlayerCellAddress = "" Then Exit Sub

    Dim mapSheet As Worksheet
    Set mapSheet = Sheets(gs.CurrentScreen)

    Dim baseCell As Range
    Set baseCell = mapSheet.Range(gs.PlayerCellAddress)

    Dim baseRow As Long
    Dim baseColumn As Long
    baseRow = baseCell.Row
    baseColumn = baseCell.Column

    Dim scrollCode As String
    scrollCode = Trim$(scrollDir)

    Dim orientation As String
    orientation = ScrollOrientation(scrollCode)

    Dim primaryDir As String
    primaryDir = ExtractDirectionalComponent(direction, orientation)
    If primaryDir = "" Then
        primaryDir = ScrollCodeToDirectionLetter(scrollCode)
        If primaryDir = "" Then
            primaryDir = ExtractDirectionalComponent(gs.lastDir, orientation)
        End If
    End If

    Dim targetRow As Long
    Dim targetColumn As Long
    targetRow = baseRow
    targetColumn = baseColumn

    Select Case scrollCode
        Case SCROLL_CODE_VERTICAL
            Select Case primaryDir
                Case "D"
                    targetRow = baseRow + 5
                Case Else
                    targetRow = baseRow
            End Select
            targetColumn = baseColumn
        Case SCROLL_CODE_HORIZONTAL
            Select Case primaryDir
                Case "R"
                    targetColumn = baseColumn + 2
                Case "L"
                    targetColumn = baseColumn - 2
                Case Else
                    targetColumn = baseColumn
            End Select
            targetRow = baseRow
        Case SCROLL_CODE_DIRECT_DOWN
            targetColumn = baseColumn
            targetRow = baseRow + 5
        Case SCROLL_CODE_DIRECT_UP
            targetColumn = baseColumn
            targetRow = baseRow
        Case ""
            targetColumn = baseColumn
            targetRow = baseRow
        Case Else
            Select Case primaryDir
                Case "D"
                    targetColumn = baseColumn
                    targetRow = baseRow + 5
                Case "U"
                    targetColumn = baseColumn
                    targetRow = baseRow
                Case "L"
                    targetRow = baseRow
                    targetColumn = baseColumn - 2
                Case "R"
                    targetRow = baseRow
                    targetColumn = baseColumn + 2
                Case Else
                    targetColumn = baseColumn
                    targetRow = baseRow
            End Select
    End Select

    If targetColumn < 1 Then targetColumn = 1
    If targetRow < 1 Then targetRow = 1
    If targetColumn > mapSheet.Columns.Count Then targetColumn = mapSheet.Columns.Count
    If targetRow > mapSheet.Rows.Count Then targetRow = mapSheet.Rows.Count

    Dim rowLabel As String
    Dim columnLabel As String
    rowLabel = Trim$(CStr(mapSheet.Cells(targetRow, 7).Value))
    columnLabel = Trim$(CStr(mapSheet.Cells(1, targetColumn).Value))

    Dim screenCode As String
    screenCode = UCase$(Trim$(rowLabel & columnLabel))

    gs.CurrentScreenCode = screenCode

    DataCacheInstance().SetValue RANGE_SCREEN_ROW, rowLabel
    DataCacheInstance().SetValue RANGE_SCREEN_COLUMN, columnLabel

    FocusOnScreen screenCode

    Exit Sub

ErrorHandler:
    MsgBox "Error in UpdateScreenLocation: " & Err.Description, vbCritical, "Screen Location Error"
End Sub

'===================================================================================
'                              PUBLIC API - Camera Following
'===================================================================================
''' <summary>
''' Maintains viewport centered on player. Call every frame from game loop.
''' Uses instant snap for jitter-free scrolling (sprite interpolation provides smoothness).
''' </summary>
Public Sub MaintainPlayerViewport(Optional ByVal rowMargin As Long = -1, Optional ByVal columnMargin As Long = -1)
    ' Get player sprite
    Dim spriteManager As SpriteManager
    Set spriteManager = SpriteManagerInstance()
    If spriteManager Is Nothing Then Exit Sub

    Dim playerSprite As Shape
    Set playerSprite = spriteManager.PlayerSprite
    If playerSprite Is Nothing Then Exit Sub

    Dim playerCell As Range
    Set playerCell = playerSprite.TopLeftCell
    If playerCell Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = playerCell.Worksheet

    ' Ensure worksheet is active
    Dim win As Window
    If Not ws Is ActiveSheet Then ws.Activate
    
    On Error Resume Next
    Set win = ActiveWindow
    On Error GoTo 0
    If win Is Nothing Then Exit Sub

    Dim previousAnchor As String
    previousAnchor = m_View.LastAnchor

    ' Get current viewport state
    Dim visible As Range
    On Error Resume Next
    Set visible = win.VisibleRange
    On Error GoTo 0
    If visible Is Nothing Then Exit Sub

    Dim visibleRows As Long, visibleColumns As Long
    visibleRows = visible.Rows.Count
    visibleColumns = visible.Columns.Count
    If visibleRows <= 0 Then visibleRows = DEFAULT_VISIBLE_ROWS
    If visibleColumns <= 0 Then visibleColumns = DEFAULT_VISIBLE_COLUMNS

    ' Calculate desired camera position based on mode
    Dim desiredTop As Long, desiredLeft As Long
    Dim pendingScreenCode As String
    Dim shouldTrackScreen As Boolean
    
    Select Case m_View.Mode
        Case CameraInstant
            ' Center camera on player instantly (default, Unity-style)
            desiredTop = playerCell.Row - (visibleRows \ 2)
            desiredLeft = playerCell.Column - (visibleColumns \ 2)
            
        Case CameraMargins
            ' Only scroll when player approaches viewport edges (platformer-style)
            Dim rMargin As Long, cMargin As Long
            rMargin = IIf(rowMargin < 0, DEFAULT_ROW_MARGIN, rowMargin)
            cMargin = IIf(columnMargin < 0, DEFAULT_COLUMN_MARGIN, columnMargin)
            
            Dim topRow As Long, leftColumn As Long
            topRow = win.ScrollRow
            leftColumn = win.ScrollColumn
            
            desiredTop = topRow
            desiredLeft = leftColumn
            
            ' Adjust if player near edges
            If playerCell.Row < topRow + rMargin Then
                desiredTop = playerCell.Row - rMargin
            ElseIf playerCell.Row > topRow + visibleRows - rMargin - 1 Then
                desiredTop = playerCell.Row + rMargin - visibleRows + 1
            End If
            
            If playerCell.Column < leftColumn + cMargin Then
                desiredLeft = playerCell.Column - cMargin
            ElseIf playerCell.Column > leftColumn + visibleColumns - cMargin - 1 Then
                desiredLeft = playerCell.Column + cMargin - visibleColumns + 1
            End If
            
        Case CameraGrid
            ' Snap to predefined screen anchor using row/column offset metadata
            Dim rowOffset As Long
            Dim columnOffset As Long
            rowOffset = SafeOffset(ws.Cells(playerCell.Row, SCREEN_ROW_OFFSET_COLUMN).Value, visibleRows \ 2)
            columnOffset = SafeOffset(ws.Cells(SCREEN_COLUMN_OFFSET_ROW, playerCell.Column).Value, visibleColumns \ 2)

            desiredTop = playerCell.Row - rowOffset + 1
            desiredLeft = playerCell.Column - columnOffset + 1
            pendingScreenCode = ComputeScreenCodeForCell(playerCell)
            shouldTrackScreen = True
    End Select

    ' Apply configured bounds if set
    If m_View.BoundsMinRow > 0 Or m_View.BoundsMaxRow > 0 Or _
       m_View.BoundsMinColumn > 0 Or m_View.BoundsMaxColumn > 0 Then
        ApplyConfiguredBounds visibleRows, visibleColumns, desiredTop, desiredLeft
    End If

    ' Clamp to sheet bounds
    Dim maxTop As Long, maxLeft As Long
    maxTop = ws.Rows.Count - visibleRows + 1
    maxLeft = ws.Columns.Count - visibleColumns + 1
    If maxTop < 1 Then maxTop = 1
    If maxLeft < 1 Then maxLeft = 1
    
    desiredTop = ClampLong(desiredTop, 1, maxTop)
    desiredLeft = ClampLong(desiredLeft, 1, maxLeft)

    Dim finalAnchorAddress As String
    On Error Resume Next
    finalAnchorAddress = ws.Cells(desiredTop, desiredLeft).Address(False, False)
    On Error GoTo 0

    If shouldTrackScreen Then
        If previousAnchor <> finalAnchorAddress Then
            Dim resolvedScreenCode As String
            resolvedScreenCode = pendingScreenCode
            If resolvedScreenCode = "" Then
                resolvedScreenCode = ComputeScreenCodeForCell(playerCell)
            End If
            UpdateScreenTracking playerCell, resolvedScreenCode
        End If
    End If

    ' Apply scroll if changed (instant snap, no smoothing)
    If desiredTop <> win.ScrollRow Or desiredLeft <> win.ScrollColumn Then
        Dim prevUpdating As Boolean
        prevUpdating = Application.ScreenUpdating
        If prevUpdating Then Application.ScreenUpdating = False

        On Error Resume Next
        If desiredTop <> win.ScrollRow Then win.ScrollRow = desiredTop
        If desiredLeft <> win.ScrollColumn Then win.ScrollColumn = desiredLeft
        On Error GoTo 0

        If prevUpdating Then Application.ScreenUpdating = True
        
        On Error Resume Next
        m_View.LastAnchor = ws.Cells(desiredTop, desiredLeft).Address(False, False)
        On Error GoTo 0
    End If
End Sub

'===================================================================================
'                              PRIVATE HELPERS - Camera
'===================================================================================
Private Sub ApplyConfiguredBounds(ByVal visibleRows As Long, ByVal visibleColumns As Long, _
                                  ByRef topRow As Long, ByRef leftColumn As Long)
    ' Apply user-configured camera bounds
    If m_View.BoundsMinColumn > 0 And m_View.BoundsMaxColumn > 0 Then
        Dim minLeft As Long, maxLeft As Long
        minLeft = m_View.BoundsMinColumn
        maxLeft = m_View.BoundsMaxColumn - visibleColumns + 1
        If maxLeft < minLeft Then maxLeft = minLeft
        leftColumn = ClampLong(leftColumn, minLeft, maxLeft)
    End If
    
    If m_View.BoundsMinRow > 0 And m_View.BoundsMaxRow > 0 Then
        Dim minTop As Long, maxTop As Long
        minTop = m_View.BoundsMinRow
        maxTop = m_View.BoundsMaxRow - visibleRows + 1
        If maxTop < minTop Then maxTop = minTop
        topRow = ClampLong(topRow, minTop, maxTop)
    End If
End Sub

Private Function ComputeScreenCodeForCell(ByVal cell As Range) As String
    If cell Is Nothing Then Exit Function

    Dim ws As Worksheet
    Set ws = cell.Worksheet

    Dim rowLabel As String
    Dim columnLabel As String

    rowLabel = UCase$(Trim$(CStr(ws.Cells(cell.Row, SCREEN_ROW_LABEL_COLUMN).Value)))
    columnLabel = UCase$(Trim$(CStr(ws.Cells(SCREEN_COLUMN_LABEL_ROW, cell.Column).Value)))

    If rowLabel <> "" And columnLabel <> "" Then
        ComputeScreenCodeForCell = rowLabel & columnLabel
    End If
End Function

Private Sub UpdateScreenTracking(ByVal playerCell As Range, ByVal screenCode As String)
    If playerCell Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = playerCell.Worksheet

    Dim rowLabel As String
    Dim columnLabel As String

    rowLabel = UCase$(Trim$(CStr(ws.Cells(playerCell.Row, SCREEN_ROW_LABEL_COLUMN).Value)))
    columnLabel = UCase$(Trim$(CStr(ws.Cells(SCREEN_COLUMN_LABEL_ROW, playerCell.Column).Value)))

    Dim cache As DataCache
    Set cache = DataCacheInstance()
    If Not cache Is Nothing Then
        If rowLabel <> "" Then cache.SetValue RANGE_SCREEN_ROW, rowLabel
        If columnLabel <> "" Then cache.SetValue RANGE_SCREEN_COLUMN, columnLabel
    End If

    Dim gs As GameState
    Set gs = GameStateInstance()
    If Not gs Is Nothing Then
        If screenCode <> "" Then gs.CurrentScreenCode = screenCode
    End If
End Sub

'===================================================================================
'                              PRIVATE HELPERS - Math
'===================================================================================

Private Function ClampLong(ByVal value As Long, ByVal minValue As Long, ByVal maxValue As Long) As Long
    If maxValue < minValue Then
        ClampLong = minValue
        Exit Function
    End If
    If value < minValue Then value = minValue
    If value > maxValue Then value = maxValue
    ClampLong = value
End Function

Private Function ClampDouble(ByVal value As Double, ByVal minValue As Double, ByVal maxValue As Double) As Double
    If maxValue < minValue Then
        ClampDouble = minValue
        Exit Function
    End If
    If value < minValue Then value = minValue
    If value > maxValue Then value = maxValue
    ClampDouble = value
End Function

Private Function NormalizeScreenCode(ByVal screenCode As String) As String
    NormalizeScreenCode = UCase$(Trim$(screenCode))
End Function

'===================================================================================
'                              PRIVATE HELPERS - Scrolling
'===================================================================================
Private Function ResolveScrollDirection(ByVal scrollCode As String, ByVal moveDir As String, ByVal lastDir As String) As String
    Dim mapped As String
    mapped = ScrollCodeToDirectionLetter(scrollCode)
    If mapped <> "" Then
        ResolveScrollDirection = mapped
        Exit Function
    End If

    Dim orientation As String
    orientation = ScrollOrientation(scrollCode)

    Dim moveCandidate As String
    Dim lastCandidate As String
    Dim prevCandidate As String
    Dim resolved As String

    Dim liveMove As String
    liveMove = NormalizeDirectionCandidate(DataCacheInstance().GetValue(RANGE_MOVE_DIR))

    moveCandidate = NormalizeDirectionCandidate(moveDir)
    lastCandidate = NormalizeDirectionCandidate(lastDir)
    prevCandidate = NormalizeDirectionCandidate(DataCacheInstance().GetValue(RANGE_PREVIOUS_SCROLL))

    resolved = ExtractDirectionalComponent(liveMove, orientation)
    If resolved <> "" Then
        ResolveScrollDirection = resolved
        Exit Function
    End If

    resolved = ExtractDirectionalComponent(moveCandidate, orientation)
    If resolved <> "" Then
        ResolveScrollDirection = resolved
        Exit Function
    End If

    resolved = ExtractDirectionalComponent(lastCandidate, orientation)
    If resolved <> "" Then
        ResolveScrollDirection = resolved
        Exit Function
    End If

    resolved = ExtractDirectionalComponent(prevCandidate, orientation)
    If resolved <> "" Then
        ResolveScrollDirection = resolved
        Exit Function
    End If

    resolved = ExtractDirectionalComponent(liveMove, "")
    If resolved <> "" Then
        ResolveScrollDirection = resolved
        Exit Function
    End If

    resolved = ExtractDirectionalComponent(moveCandidate, "")
    If resolved <> "" Then
        ResolveScrollDirection = resolved
        Exit Function
    End If

    resolved = ExtractDirectionalComponent(lastCandidate, "")
    If resolved <> "" Then
        ResolveScrollDirection = resolved
        Exit Function
    End If

    ResolveScrollDirection = ExtractDirectionalComponent(prevCandidate, "")
End Function

Private Function ScrollCodeToDirectionLetter(ByVal scrollCode As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(scrollCode))

    Select Case normalized
        Case SCROLL_CODE_DIRECT_DOWN, "3", "D"
            ScrollCodeToDirectionLetter = "D"
        Case SCROLL_CODE_DIRECT_UP, "4", "U"
            ScrollCodeToDirectionLetter = "U"
        Case "R"
            ScrollCodeToDirectionLetter = "R"
        Case "L"
            ScrollCodeToDirectionLetter = "L"
        Case Else
            ScrollCodeToDirectionLetter = ""
    End Select
End Function

Private Function ScrollOrientation(ByVal scrollCode As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(scrollCode))

    Select Case normalized
        Case SCROLL_CODE_HORIZONTAL, "2", "R", "L"
            ScrollOrientation = "H"
        Case SCROLL_CODE_VERTICAL, "1", _
             SCROLL_CODE_DIRECT_DOWN, "3", _
             SCROLL_CODE_DIRECT_UP, "4", _
             "U", "D"
            ScrollOrientation = "V"
        Case Else
            ScrollOrientation = ""
    End Select
End Function

Private Function NormalizeDirectionCandidate(ByVal value As String) As String
    NormalizeDirectionCandidate = UCase$(Trim$(value))
End Function

Private Function ExtractDirectionalComponent(ByVal direction As String, ByVal orientation As String) As String
    Dim normalized As String
    normalized = NormalizeDirectionCandidate(direction)
    If normalized = "" Then Exit Function

    Dim index As Long
    Dim ch As String

    If orientation <> "" Then
        For index = 1 To Len(normalized)
            ch = Mid$(normalized, index, 1)
            Select Case orientation
                Case "H"
                    If ch = "L" Or ch = "R" Then
                        ExtractDirectionalComponent = ch
                        Exit Function
                    End If
                Case "V"
                    If ch = "U" Or ch = "D" Then
                        ExtractDirectionalComponent = ch
                        Exit Function
                    End If
            End Select
        Next index
    End If

    For index = 1 To Len(normalized)
        ch = Mid$(normalized, index, 1)
        Select Case ch
            Case "U", "D", "L", "R"
                ExtractDirectionalComponent = ch
                Exit Function
        End Select
    Next index
End Function

Private Function ShouldPreventRescroll(ByVal currentCell As String, ByVal newDirection As String) As Boolean
    Dim previousCell As String
    Dim previousDir As String
    Dim currentDir As String

    previousCell = DataCacheInstance().GetValue(RANGE_PREVIOUS_CELL)
    previousDir = DataCacheInstance().GetValue(RANGE_PREVIOUS_SCROLL)
    currentDir = DataCacheInstance().GetValue(RANGE_SCROLL_DIRECTION)

    Dim normalizedNewDir As String
    normalizedNewDir = NormalizeDirectionCandidate(newDirection)
    Dim comparisonDir As String
    comparisonDir = NormalizeDirectionCandidate(currentDir)
    If normalizedNewDir <> "" Then comparisonDir = normalizedNewDir

    Dim normalizedPrev As String
    normalizedPrev = NormalizeDirectionCandidate(previousDir)

    If NormalizeDirectionCandidate(currentCell) = NormalizeDirectionCandidate(previousCell) Then
        If normalizedPrev <> "" And comparisonDir <> "" Then
            If normalizedPrev = comparisonDir Then
                ShouldPreventRescroll = True
                Exit Function
            End If
        End If
    End If

    ShouldPreventRescroll = False
End Function

Private Sub PerformWindowScroll(ByVal scrollCode As String, ByVal direction As String)
    Select Case Trim$(scrollCode)
        Case SCROLL_CODE_HORIZONTAL
            Select Case UCase$(direction)
                Case "R": ActiveWindow.SmallScroll toRight:=SCROLL_AMOUNT_HORIZONTAL
                Case "L": ActiveWindow.SmallScroll toLeft:=SCROLL_AMOUNT_HORIZONTAL
            End Select
        Case SCROLL_CODE_VERTICAL
            Select Case UCase$(direction)
                Case "D": ActiveWindow.SmallScroll Down:=SCROLL_AMOUNT_VERTICAL
                Case "U": ActiveWindow.SmallScroll Up:=SCROLL_AMOUNT_VERTICAL
            End Select
        Case SCROLL_CODE_DIRECT_DOWN
            ActiveWindow.SmallScroll Down:=SCROLL_AMOUNT_VERTICAL
        Case SCROLL_CODE_DIRECT_UP
            ActiveWindow.SmallScroll Up:=SCROLL_AMOUNT_VERTICAL
        Case Else
            Select Case UCase$(direction)
                Case "R": ActiveWindow.SmallScroll toRight:=SCROLL_AMOUNT_HORIZONTAL
                Case "L": ActiveWindow.SmallScroll toLeft:=SCROLL_AMOUNT_HORIZONTAL
                Case "D": ActiveWindow.SmallScroll Down:=SCROLL_AMOUNT_VERTICAL
                Case "U": ActiveWindow.SmallScroll Up:=SCROLL_AMOUNT_VERTICAL
            End Select
    End Select
End Sub