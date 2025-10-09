'Attribute VB_Name = "AB_GameMechanics"
Option Explicit

'###################################################################################
'                              EXCELDA II - GAME MECHANICS
'###################################################################################
' Simplified scrolling and screen mechanics
'###################################################################################

Sub myScroll(ByVal scrollDir As String)
    On Error GoTo ErrorHandler
    
    Dim gs As GameState
    Set gs = GameStateInstance()
    
    Dim linkDirection As String
    linkDirection = gs.MoveDir
    
    ' Store previous cell for rescroll detection
    Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_CELL).Value = gs.LinkCellAddress
    
    ' Extract primary direction for scroll
    Dim primaryDir As String
    primaryDir = ExtractPrimaryDirection(linkDirection, scrollDir)
    If primaryDir = "" Then Exit Sub
    
    ' Store scroll direction for rescroll prevention
    Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_SCROLL).Value = primaryDir
    
    ' Check if we should prevent rescrolling
    If ShouldPreventRescroll(gs.LinkCellAddress) Then Exit Sub
    
    ' Perform the scroll
    PerformWindowScroll scrollDir, primaryDir
    
    ' Update scroll state
    Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value = primaryDir
    
    ' Calculate and set up new screen
    Call calculateScreenLocation(scrollDir, linkDirection)
    
    ' Run screen setup macro
    On Error GoTo ScreenSetupError
    mySub = CurrentScreen 'global
    Application.Run mySub
    
    Exit Sub
    
ScreenSetupError:
    MsgBox "Screen setup macro not found: " & mySub, vbCritical, "Screen Setup Error"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in myScroll: " & Err.Description, vbCritical, "Scroll Error"
End Sub

'###################################################################################
'                              Helper Functions
'###################################################################################

Private Function ExtractPrimaryDirection(ByVal linkDir As String, ByVal scrollType As String) As String
    ' Extract the primary direction based on scroll type
    ' Returns single character: U, D, L, R, or empty string
    
    If scrollType = SCROLL_VERTICAL Then
        ' Vertical scroll - check for U or D in direction
        If InStr(linkDir, "U") > 0 Then
            ExtractPrimaryDirection = "U"
        ElseIf InStr(linkDir, "D") > 0 Then
            ExtractPrimaryDirection = "D"
        Else
            ExtractPrimaryDirection = ""
        End If
    ElseIf scrollType = SCROLL_HORIZONTAL Then
        ' Horizontal scroll - check for L or R in direction
        If InStr(linkDir, "L") > 0 Then
            ExtractPrimaryDirection = "L"
        ElseIf InStr(linkDir, "R") > 0 Then
            ExtractPrimaryDirection = "R"
        Else
            ExtractPrimaryDirection = ""
        End If
    Else
        ExtractPrimaryDirection = ""
    End If
End Function

Private Function ShouldPreventRescroll(ByVal currentCell As String) As Boolean
    ' Simple rescroll prevention - check if we're in same general area
    Dim previousCell As String
    Dim previousDir As String
    Dim currentDir As String
    
    previousCell = Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_CELL).Value
    previousDir = Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_SCROLL).Value
    currentDir = Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value
    
    ' If we're still in the same cell and direction hasn't changed, prevent rescroll
    If currentCell = previousCell Then
        If previousDir = currentDir Or previousDir = "" Then
            ShouldPreventRescroll = True
            Exit Function
        End If
    End If
    
    ShouldPreventRescroll = False
End Function

Private Sub PerformWindowScroll(ByVal scrollType As String, ByVal direction As String)
    ' Simple, unified scroll logic
    
    If scrollType = SCROLL_VERTICAL Then
        If direction = "D" Then
            ActiveWindow.SmallScroll Down:=SCROLL_AMOUNT_VERTICAL
        ElseIf direction = "U" Then
            ActiveWindow.SmallScroll Up:=SCROLL_AMOUNT_VERTICAL
        End If
    ElseIf scrollType = SCROLL_HORIZONTAL Then
        If direction = "L" Then
            ActiveWindow.SmallScroll toLeft:=SCROLL_AMOUNT_HORIZONTAL
        ElseIf direction = "R" Then
            ActiveWindow.SmallScroll toRight:=SCROLL_AMOUNT_HORIZONTAL
        End If
    End If
End Sub

'###################################################################################
'                              Screen Location & Alignment
'###################################################################################

Sub calculateScreenLocation(ByVal scrollDir As String, ByVal linkDirection As String)
    On Error GoTo ErrorHandler
    
    Dim gs As GameState
    Set gs = GameStateInstance()
    
    Dim myColumn As Long, myRow As Long
    Dim baseRow As Long, baseColumn As Long
    Dim baseCell As Range
    
    ' Get base position from sprite
    Set baseCell = Range(gs.LinkCellAddress)
    baseRow = baseCell.Row
    baseColumn = baseCell.Column
    
    ' Calculate position based on scroll type
    If scrollDir = SCROLL_VERTICAL Then
        myColumn = baseColumn
        ' Vertical: adjust row based on direction
        If InStr(linkDirection, "U") > 0 Then
            myRow = baseRow
        ElseIf InStr(linkDirection, "D") > 0 Then
            myRow = baseRow + 5
        Else
            myRow = baseRow
        End If
    ElseIf scrollDir = SCROLL_HORIZONTAL Then
        myRow = baseRow
        ' Horizontal: adjust column based on direction
        If InStr(linkDirection, "L") > 0 Then
            myColumn = baseColumn - 2
        ElseIf InStr(linkDirection, "R") > 0 Then
            myColumn = baseColumn + 2
        Else
            myColumn = baseColumn
        End If
    Else
        ' No scroll - use current position
        myColumn = baseColumn
        myRow = baseRow
    End If
    
    ' Get screen identifiers from calculated position
    Dim myRowValue As String, myColumnValue As String
    myRowValue = Cells(myRow, 7).Value
    myColumnValue = Cells(1, myColumn).Value
    
    ' Set current screen
    CurrentScreen = myRowValue & myColumnValue
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in calculateScreenLocation: " & Err.Description, vbCritical, "Calculate Error"
End Sub

Sub alignScreen()
    On Error GoTo ErrorHandler
    
    Dim myColumn As Long, myRow As Long
    Dim offsetRow As Long, offsetColumn As Long
    Dim myTopLeft As String
    
    ' Get current position
    myColumn = ActiveCell.Column
    myRow = ActiveCell.Row
    
    ' Get screen offset values from data sheet
    offsetRow = Cells(myRow, 8).Value
    offsetColumn = Cells(2, myColumn).Value
    
    ' Calculate top-left position for screen alignment
    myTopLeft = ActiveCell.Offset(-offsetRow + 1, -offsetColumn + 1).Address
    
    ' Navigate to calculated position
    Application.GoTo ActiveSheet.Range(myTopLeft), True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in alignScreen: " & Err.Description, vbCritical, "Align Error"
End Sub
