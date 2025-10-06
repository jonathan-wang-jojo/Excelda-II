Option Explicit

'###################################################################################
'                              EXCELDA II - GAME MECHANICS
'###################################################################################

Sub myScroll(ByVal scrollDir As String)
    On Error GoTo ScrollError
    
    Dim linkDirection As String
    linkDirection = Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value
    
    ' Set current cell to compare for next scroll
    Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_CELL).Value = linkCellAddress
    
    ' Set current direction of travel to prevent rescrolling
    SetScrollDirection scrollDir, linkDirection
    
    ' Check if we should prevent rescrolling
    If ShouldPreventRescroll() Then Exit Sub
    
    ' Set up scroll coordinates
    SetupScrollCoordinates
    
    ' Perform the actual scroll
    PerformScroll scrollDir, linkDirection
    
    ' Calculate and set up new screen
    Call calculateScreenLocation(scrollDir, linkDirection)
    
    ' Run screen setup macro
    On Error GoTo ScreenSetupError
    mySub = currentScreen 'global
    Application.Run mySub
    
    Exit Sub
    
ScreenSetupError:
    MsgBox "Screen setup macro not found: " & mySub, vbCritical, "Screen Setup Error"
    Exit Sub
    
ScrollError:
    MsgBox "Error in myScroll: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Scroll Error"
End Sub

Private Sub SetScrollDirection(ByVal scrollDir As String, ByVal linkDirection As String)
    Dim directionValue As String
    
    Select Case scrollDir
        Case SCROLL_VERTICAL
            directionValue = GetVerticalScrollDirection(linkDirection)
        Case SCROLL_HORIZONTAL
            directionValue = GetHorizontalScrollDirection(linkDirection)
    End Select
    
    Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_SCROLL).Value = directionValue
End Sub

Private Function GetVerticalScrollDirection(ByVal linkDirection As String) As String
    Select Case linkDirection
        Case "U", "RU", "LU"
            GetVerticalScrollDirection = "U"
        Case "D", "RD", "LD"
            GetVerticalScrollDirection = "D"
        Case Else
            GetVerticalScrollDirection = ""
    End Select
End Function

Private Function GetHorizontalScrollDirection(ByVal linkDirection As String) As String
    Select Case linkDirection
        Case "U"
            GetHorizontalScrollDirection = ""
        Case "RU", "R", "RD"
            GetHorizontalScrollDirection = "R"
        Case "LU", "L", "LD"
            GetHorizontalScrollDirection = "L"
        Case "D"
            GetHorizontalScrollDirection = ""
        Case Else
            GetHorizontalScrollDirection = ""
    End Select
End Function

Private Function ShouldPreventRescroll() As Boolean
    Dim currentCell As String
    Dim previousCell As String
    Dim currentDirection As String
    Dim previousDirection As String
    
    currentCell = Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value
    previousCell = Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_CELL).Value
    currentDirection = Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value
    previousDirection = Sheets(SHEET_DATA).Range(RANGE_PREVIOUS_SCROLL).Value
    
    ' Check if we're in the same cell or adjacent cells
    If IsSameOrAdjacentCell(currentCell, previousCell) Then
        ' Check if direction is the same or empty
        If previousDirection = currentDirection Or previousDirection = "" Then
            ShouldPreventRescroll = True
        End If
    End If
End Function

Private Function IsSameOrAdjacentCell(ByVal currentCell As String, ByVal previousCell As String) As Boolean
    ' Check if current cell matches any of the stored adjacent cells
    Dim i As Integer
    For i = 8 To 12 ' Columns E through L (C8 to L8)
        If Sheets(SHEET_DATA).Cells(8, i).Value = previousCell Then
            IsSameOrAdjacentCell = True
            Exit Function
        End If
    Next i
    
    IsSameOrAdjacentCell = (currentCell = previousCell)
End Function

Private Sub SetupScrollCoordinates()
    Dim baseCell As Range
    Set baseCell = Range(Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value)
    
    ' Set current cell
    Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = linkCellAddress
    
    ' Set up adjacent cells for scroll detection
    Sheets(SHEET_DATA).Range("E8").Value = baseCell.Offset(0, 1).Address   ' Right
    Sheets(SHEET_DATA).Range("F8").Value = baseCell.Offset(0, 2).Address   ' Right+1
    Sheets(SHEET_DATA).Range("G8").Value = baseCell.Offset(0, -1).Address  ' Left
    Sheets(SHEET_DATA).Range("H8").Value = baseCell.Offset(0, -2).Address    ' Left+1
    Sheets(SHEET_DATA).Range("I8").Value = baseCell.Offset(-1, 0).Address  ' Up
    Sheets(SHEET_DATA).Range("J8").Value = baseCell.Offset(-2, 0).Address  ' Up+1
    Sheets(SHEET_DATA).Range("K8").Value = baseCell.Offset(1, 0).Address   ' Down
    Sheets(SHEET_DATA).Range("L8").Value = baseCell.Offset(2, 0).Address   ' Down+1
End Sub

Private Sub PerformScroll(ByVal scrollDir As String, ByVal linkDirection As String)
    Select Case scrollDir
        Case SCROLL_VERTICAL
            PerformVerticalScroll linkDirection
        Case SCROLL_HORIZONTAL
            PerformHorizontalScroll linkDirection
    End Select
End Sub

Private Sub PerformVerticalScroll(ByVal linkDirection As String)
    Select Case linkDirection
        Case "D", "RD", "LD"
            ActiveWindow.SmallScroll Down:=SCROLL_AMOUNT_VERTICAL
            Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value = "D"
        Case "U", "RU", "LU"
            ActiveWindow.SmallScroll Up:=SCROLL_AMOUNT_VERTICAL
            Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value = "U"
    End Select
End Sub

Private Sub PerformHorizontalScroll(ByVal linkDirection As String)
    Select Case linkDirection
        Case "L", "LU", "LD"
            ActiveWindow.SmallScroll toLeft:=SCROLL_AMOUNT_HORIZONTAL
            Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value = "L"
        Case "R", "RU", "RD"
            ActiveWindow.SmallScroll toRight:=SCROLL_AMOUNT_HORIZONTAL
            Sheets(SHEET_DATA).Range(RANGE_SCROLL_DIRECTION).Value = "R"
    End Select
End Sub

'###################################################################################
'                              Screen Location Calculation
'###################################################################################

Sub calculateScreenLocation(ByVal scrollDir As String, ByVal linkDirection As String)
    On Error GoTo CalculateError
    
    Dim myColumn As Long, myRow As Long
    Dim myColumnValue As String, myRowValue As String
    
    linkCellAddress = SpriteManager.Instance.LinkSprite.TopLeftCell.Address
    
    ' Calculate screen position based on scroll direction
    Select Case scrollDir
        Case SCROLL_VERTICAL
            CalculateVerticalPosition linkDirection, myColumn, myRow
        Case SCROLL_HORIZONTAL
            CalculateHorizontalPosition linkDirection, myColumn, myRow
    End Select
    
    ' Get screen identifiers
    myRowValue = Cells(myRow, 7).Value
    myColumnValue = Cells(1, myColumn).Value
    
    ' Set current screen
    currentScreen = myRowValue & myColumnValue
    
    Exit Sub
    
CalculateError:
    MsgBox "Error in calculateScreenLocation: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Calculate Error"
End Sub

Private Sub CalculateVerticalPosition(ByVal linkDirection As String, ByRef myColumn As Long, ByRef myRow As Long)
    Dim baseRow As Long
    baseRow = Range(linkCellAddress).Row
    
    Select Case linkDirection
        Case "U", "LU", "RU"
            myColumn = Range(linkCellAddress).Column
            myRow = baseRow
        Case "D", "RD", "LD"
            myColumn = Range(linkCellAddress).Column
            myRow = baseRow + 5
        Case Else
            myColumn = Range(linkCellAddress).Column
            myRow = baseRow
    End Select
End Sub

Private Sub CalculateHorizontalPosition(ByVal linkDirection As String, ByRef myColumn As Long, ByRef myRow As Long)
    Dim baseColumn As Long
    baseColumn = Range(linkCellAddress).Column
    
    Select Case linkDirection
        Case "LU", "LD", "L"
            myColumn = baseColumn - 2
            myRow = Range(linkCellAddress).Row
        Case "RU", "R", "RD"
            myColumn = baseColumn + 2
            myRow = Range(linkCellAddress).Row
        Case Else
            myColumn = baseColumn
            myRow = Range(linkCellAddress).Row
    End Select
End Sub

'###################################################################################
'                              Screen Alignment
'###################################################################################

Sub alignScreen()
    On Error GoTo AlignError
    
    Dim myColumn As Long, myRow As Long
    Dim myColumnValue As Long, myRowValue As Long
    Dim myTopLeft As String
    
    myColumn = ActiveCell.Column
    myRow = ActiveCell.Row
    
    ' Get screen offset values
    myRowValue = Cells(myRow, 8).Value
    myColumnValue = Cells(2, myColumn).Value
    
    ' Calculate top-left position
    myTopLeft = ActiveCell.Offset(-myRowValue + 1, -myColumnValue + 1).Address
    
    ' Navigate to calculated position
    Application.GoTo ActiveSheet.Range(myTopLeft), True
    
    Exit Sub
    
AlignError:
    MsgBox "Error in alignScreen: " & Err.Description & " (Error " & Err.Number & ")", vbCritical, "Align Error"
End Sub