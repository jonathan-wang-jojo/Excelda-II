
'Attribute VB_Name = "AB_GameMechanics"
Option Explicit

'###################################################################################
'                              EXCELDA II - GAME MECHANICS
'###################################################################################

Sub myScroll(ByVal scrollDir As String)
    On Error GoTo ErrorHandler

    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Sub

    Dim scrollCode As String
    scrollCode = Trim$(CStr(scrollDir))
    If scrollCode = "" Then Exit Sub

    ' Determine the intended scroll direction (U/D/L/R)
    Dim primaryDir As String
    primaryDir = ResolveScrollDirection(scrollCode, gs.MoveDir, gs.LastDir)
    If primaryDir = "" Then Exit Sub

    If ShouldPreventRescroll(gs.LinkCellAddress, primaryDir) Then Exit Sub

    ' Perform the viewport scroll
    PerformWindowScroll scrollCode, primaryDir

    ' Persist scroll direction and recalc screen code
    Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_CELL).Value = gs.LinkCellAddress
    Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_SCROLL).Value = primaryDir
    Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value = primaryDir
    calculateScreenLocation scrollCode, primaryDir

    ' Execute the target screen's setup routine
    On Error GoTo ScreenSetupError
    Dim setupMacro As String
    setupMacro = gs.CurrentScreenCode
    If setupMacro = "" Then setupMacro = gs.CurrentScreen
    If setupMacro <> "" Then
        SceneManagerInstance().ApplyScreen setupMacro
    End If

    Exit Sub

ScreenSetupError:
    MsgBox "Screen setup macro not found: " & setupMacro, vbCritical, "Screen Setup Error"
    Exit Sub

ErrorHandler:
    MsgBox "Error in myScroll: " & Err.Description, vbCritical, "Scroll Error"
End Sub

'###################################################################################
'                              Helper Functions
'###################################################################################

Private Function ResolveScrollDirection(ByVal scrollCode As String, ByVal moveDir As String, ByVal lastDir As String) As String
    Dim mapped As String
    mapped = ScrollCodeToDirectionLetter(scrollCode)
    If mapped <> "" Then
        ResolveScrollDirection = mapped
        Exit Function
    End If

    Dim candidate As String
    candidate = NormalizeDirectionCandidate(moveDir)
    If candidate = "" Then candidate = NormalizeDirectionCandidate(lastDir)
    If candidate = "" Then Exit Function

    ResolveScrollDirection = PrimaryDirectionLetter(candidate)
End Function

Private Function ScrollCodeToDirectionLetter(ByVal scrollCode As String) As String
    Select Case UCase$(Trim$(scrollCode))
        Case SCROLL_CODE_RIGHT, "R"
            ScrollCodeToDirectionLetter = "R"
        Case SCROLL_CODE_LEFT, "L"
            ScrollCodeToDirectionLetter = "L"
        Case SCROLL_CODE_DOWN, "D"
            ScrollCodeToDirectionLetter = "D"
        Case SCROLL_CODE_UP, "U"
            ScrollCodeToDirectionLetter = "U"
        Case Else
            ScrollCodeToDirectionLetter = ""
    End Select
End Function

Private Function NormalizeDirectionCandidate(ByVal value As String) As String
    NormalizeDirectionCandidate = UCase$(Trim$(value))
End Function

Private Function PrimaryDirectionLetter(ByVal direction As String) As String
    Dim normalized As String
    normalized = NormalizeDirectionCandidate(direction)
    If normalized = "" Then Exit Function

    If InStr(normalized, "U") > 0 Then
        PrimaryDirectionLetter = "U"
    ElseIf InStr(normalized, "D") > 0 Then
        PrimaryDirectionLetter = "D"
    ElseIf InStr(normalized, "L") > 0 Then
        PrimaryDirectionLetter = "L"
    ElseIf InStr(normalized, "R") > 0 Then
        PrimaryDirectionLetter = "R"
    ElseIf Len(normalized) > 0 Then
        PrimaryDirectionLetter = Mid$(normalized, 1, 1)
    End If
End Function

Private Function ShouldPreventRescroll(ByVal currentCell As String, ByVal newDirection As String) As Boolean
    ' Simple rescroll prevention - check if we're in same general area
    Dim previousCell As String
    Dim previousDir As String
    Dim currentDir As String
    
    previousCell = Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_CELL).Value
    previousDir = Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_SCROLL).Value
    currentDir = Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value
    
    Dim normalizedNewDir As String
    normalizedNewDir = NormalizeDirectionCandidate(newDirection)
    Dim comparisonDir As String
    comparisonDir = NormalizeDirectionCandidate(currentDir)
    If normalizedNewDir <> "" Then comparisonDir = normalizedNewDir

    Dim normalizedPrev As String
    normalizedPrev = NormalizeDirectionCandidate(previousDir)

    ' If we're still in the same cell and direction hasn't changed, prevent rescroll
    If currentCell = previousCell Then
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
        Case SCROLL_CODE_RIGHT
            ActiveWindow.SmallScroll toRight:=SCROLL_AMOUNT_HORIZONTAL
        Case SCROLL_CODE_LEFT
            ActiveWindow.SmallScroll toLeft:=SCROLL_AMOUNT_HORIZONTAL
        Case SCROLL_CODE_DOWN
            ActiveWindow.SmallScroll Down:=SCROLL_AMOUNT_VERTICAL
        Case SCROLL_CODE_UP
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

'###################################################################################
'                              Screen Location & Alignment
'###################################################################################

Sub calculateScreenLocation(ByVal scrollDir As String, ByVal direction As String)
    On Error GoTo ErrorHandler
    
    Dim gs As GameState
    Set gs = GameStateInstance()
    
    Dim myColumn As Long, myRow As Long
    Dim baseRow As Long, baseColumn As Long
    Dim baseCell As Range
    Dim mapSheet As Worksheet
    
    If gs.CurrentScreen = "" Or gs.LinkCellAddress = "" Then Exit Sub
    Set mapSheet = Sheets(gs.CurrentScreen)
    
    ' Get base position from sprite
    Set baseCell = mapSheet.Range(gs.LinkCellAddress)
    baseRow = baseCell.Row
    baseColumn = baseCell.Column

    Dim scrollCode As String
    scrollCode = Trim$(scrollDir)

    Dim primaryDir As String
    primaryDir = NormalizeDirectionCandidate(direction)
    If primaryDir = "" Then
        primaryDir = ScrollCodeToDirectionLetter(scrollCode)
        If primaryDir = "" Then
            primaryDir = NormalizeDirectionCandidate(gs.LastDir)
        End If
    End If

    myRow = baseRow
    myColumn = baseColumn

    Select Case scrollCode
        Case SCROLL_CODE_DOWN
            myColumn = baseColumn
            myRow = baseRow + 5
        Case SCROLL_CODE_UP
            myColumn = baseColumn
            myRow = baseRow
        Case SCROLL_CODE_LEFT
            myRow = baseRow
            myColumn = baseColumn - 2
        Case SCROLL_CODE_RIGHT
            myRow = baseRow
            myColumn = baseColumn + 2
        Case ""
            myColumn = baseColumn
            myRow = baseRow
        Case Else
            Select Case primaryDir
                Case "D"
                    myColumn = baseColumn
                    myRow = baseRow + 5
                Case "U"
                    myColumn = baseColumn
                    myRow = baseRow
                Case "L"
                    myRow = baseRow
                    myColumn = baseColumn - 2
                Case "R"
                    myRow = baseRow
                    myColumn = baseColumn + 2
                Case Else
                    myColumn = baseColumn
                    myRow = baseRow
            End Select
    End Select

    If myColumn < 1 Then myColumn = 1
    If myRow < 1 Then myRow = 1
    If myColumn > mapSheet.Columns.Count Then myColumn = mapSheet.Columns.Count
    If myRow > mapSheet.Rows.Count Then myRow = mapSheet.Rows.Count
    
    ' Get screen identifiers from calculated position
    Dim rowLabel As String
    Dim columnLabel As String
    rowLabel = Trim$(CStr(mapSheet.Cells(myRow, 7).Value))
    columnLabel = Trim$(CStr(mapSheet.Cells(1, myColumn).Value))

    Dim screenCode As String
    screenCode = UCase$(Trim$(rowLabel & columnLabel))

    gs.CurrentScreenCode = screenCode

    On Error Resume Next
    Sheets(SHEET_DATA).Range(RANGE_SCREEN_ROW).Value = rowLabel
    Sheets(SHEET_DATA).Range(RANGE_SCREEN_COLUMN).Value = columnLabel
    On Error GoTo ErrorHandler

    ViewportManagerInstance().FocusOnScreen screenCode
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in calculateScreenLocation: " & Err.Description, vbCritical, "Calculate Error"
End Sub

Sub alignScreen()
    On Error GoTo ErrorHandler
    
    ViewportManagerInstance().AlignToLink
    ViewportManagerInstance().RefreshVisibleDimensions
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in alignScreen: " & Err.Description, vbCritical, "Align Error"
End Sub