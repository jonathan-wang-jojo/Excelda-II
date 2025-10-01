Attribute VB_Name = "AB_GameMechanics"
'Scrolling

Sub myScroll(scrollDir)

Dim linkDirection
linkDirection = Sheets("Data").Range("C21").Value

'set current cell to compare for next scroll (allows scrolling in same direction but on opposite side of screen)
Sheets("Data").Range("D8").Value = linkCellAddress 'current


'set current direction of travel/intended scroll (to prevent rescrolling if in same cell and same direction)
Select Case scrollDir
    
    Case Is = 1
        Select Case linkDirection
            Case Is = "U"
                Sheets("Data").Range("D7").Value = "U"
            Case Is = "RU"
                Sheets("Data").Range("D7").Value = "U"
            Case Is = "LU"
                Sheets("Data").Range("D7").Value = "U"
            Case Is = "L"
                Sheets("Data").Range("D7").Value = ""
            Case Is = "R"
                Sheets("Data").Range("D7").Value = ""
            Case Is = "D"
                Sheets("Data").Range("D7").Value = "D"
            Case Is = "RD"
                Sheets("Data").Range("D7").Value = "D"
            Case Is = "LD"
                Sheets("Data").Range("D7").Value = "D"
            Case Else
                'MsgBox "Unknown linkDirection = " & linkDirection
        End Select
        
    Case Is = 2
         Select Case linkDirection
            Case Is = "U"
                Sheets("Data").Range("D7").Value = ""
            Case Is = "RU"
                Sheets("Data").Range("D7").Value = "R"
            Case Is = "LU"
                Sheets("Data").Range("D7").Value = "L"
            Case Is = "L"
                Sheets("Data").Range("D7").Value = "L"
            Case Is = "R"
                Sheets("Data").Range("D7").Value = "R"
            Case Is = "D"
                Sheets("Data").Range("D7").Value = ""
            Case Is = "RD"
                Sheets("Data").Range("D7").Value = "R"
            Case Is = "LD"
                Sheets("Data").Range("D7").Value = "L"
            Case Else
                'MsgBox " Unknown linkDirection = " & linkDirection
        End Select
    Case Else
        'MsgBox "Odd scrollDir"
    
End Select


'if the trigger cell and the direction of travel are the same, prevent 'rescrolling'
If Sheets("Data").Range("D8").Value = Sheets("Data").Range("C8").Value Or _
    Sheets("Data").Range("D8").Value = Sheets("Data").Range("E8").Value Or _
        Sheets("Data").Range("D8").Value = Sheets("Data").Range("F8").Value Or _
            Sheets("Data").Range("D8").Value = Sheets("Data").Range("G8").Value Or _
                Sheets("Data").Range("D8").Value = Sheets("Data").Range("H8").Value Or _
                    Sheets("Data").Range("D8").Value = Sheets("Data").Range("I8").Value Or _
                        Sheets("Data").Range("D8").Value = Sheets("Data").Range("J8").Value Or _
                            Sheets("Data").Range("D8").Value = Sheets("Data").Range("K8").Value Or _
                                Sheets("Data").Range("D8").Value = Sheets("Data").Range("L8").Value Then
    
    If Sheets("Data").Range("D7").Value = Sheets("Data").Range("C7").Value Or _
        Sheets("Data").Range("D7").Value = "" Then
        Exit Sub
        
    End If
    
End If


Sheets("Data").Range("C8").Value = linkCellAddress

'vertical scrolling (cells adjacent on same row)
Sheets("Data").Range("E8").Value = Range(Sheets("Data").Range("C8").Value).Offset(0, 1).Address
Sheets("Data").Range("F8").Value = Range(Sheets("Data").Range("C8").Value).Offset(0, 2).Address
Sheets("Data").Range("G8").Value = Range(Sheets("Data").Range("C8").Value).Offset(0, -1).Address
Sheets("Data").Range("H8").Value = Range(Sheets("Data").Range("C8").Value).Offset(0, -2).Address

'horizontal scrolling (cells adjacent on same column)
Sheets("Data").Range("I8").Value = Range(Sheets("Data").Range("C8").Value).Offset(-1, 0).Address
Sheets("Data").Range("J8").Value = Range(Sheets("Data").Range("C8").Value).Offset(-2, 0).Address
Sheets("Data").Range("K8").Value = Range(Sheets("Data").Range("C8").Value).Offset(1, 0).Address
Sheets("Data").Range("L8").Value = Range(Sheets("Data").Range("C8").Value).Offset(2, 0).Address


Select Case scrollDir

Case Is = "1"

    Select Case linkDirection

        Case Is = "D"
            ActiveWindow.SmallScroll Down:=32
            Sheets("Data").Range("C7").Value = "D"
            
        Case Is = "U"
            ActiveWindow.SmallScroll Up:=32
            Sheets("Data").Range("C7").Value = "U"

        Case Is = "RD"
            ActiveWindow.SmallScroll Down:=32
            Sheets("Data").Range("C7").Value = "D"

        Case Is = "LD"
            ActiveWindow.SmallScroll Down:=32
            Sheets("Data").Range("C7").Value = "D"

        Case Is = "RU"
            ActiveWindow.SmallScroll Up:=32
            Sheets("Data").Range("C7").Value = "U"

        Case Is = "LU"
            ActiveWindow.SmallScroll Up:=32
            Sheets("Data").Range("C7").Value = "U"
    
    End Select
        
Case Is = "2"

    Select Case linkDirection

        Case Is = "L"
            ActiveWindow.SmallScroll toleft:=60
            Sheets("Data").Range("C7").Value = "L"

        Case Is = "R"
            ActiveWindow.SmallScroll toRight:=60
            Sheets("Data").Range("C7").Value = "R"

        Case Is = "RD"
            ActiveWindow.SmallScroll toRight:=60
            Sheets("Data").Range("C7").Value = "R"
        'End If
        Case Is = "LD"
            ActiveWindow.SmallScroll toleft:=60
            Sheets("Data").Range("C7").Value = "L"
            
        Case Is = "RU"
            ActiveWindow.SmallScroll toRight:=60
            Sheets("Data").Range("C7").Value = "R"

        Case Is = "LU"
            ActiveWindow.SmallScroll toleft:=60
            Sheets("Data").Range("C7").Value = "L"

    End Select

Case Else
    'MsgBox ("odd scroll direction")

End Select

Call calculateScreenLocation(scrollDir, linkDirection)

On Error GoTo endPoint

mySub = currentScreen 'global
Application.Run mySub

Exit Sub

endPoint:
MsgBox "Screen setup macro not found: " & mySub

End Sub


'Working out which screen to set up


Sub calculateScreenLocation(scrollDir, linkDirection)

Dim myColumn, myColumnValue, myRow, myRowValue

linkCellAddress = LinkSprite.TopLeftCell.Address

Select Case scrollDir

    'Vertical
    Case Is = "1"

        Select Case linkDirection
        
            Case Is = "U"
                myColumn = Range(linkCellAddress).Column
                myRow = Range(linkCellAddress).Row
            Case Is = "LU"
                myColumn = Range(linkCellAddress).Column
                myRow = Range(linkCellAddress).Row
            Case Is = "RU"
                myColumn = Range(linkCellAddress).Column
                myRow = Range(linkCellAddress).Row
            Case Is = "D"
                myColumn = Range(linkCellAddress).Column
                myRow = Range(linkCellAddress).Row + 5
            Case Is = "RD"
                myColumn = Range(linkCellAddress).Column
                myRow = Range(linkCellAddress).Row + 5
            Case Is = "LD"
                myColumn = Range(linkCellAddress).Column
                myRow = Range(linkCellAddress).Row + 5
            Case Else
                'MsgBox ("Unknown linkDirection")
                myColumn = Range(linkCellAddress).Column
                myRow = Range(linkCellAddress).Row
        End Select
        
        
    'Horizontal
    Case Is = "2"
    
        Select Case linkDirection
        
            Case Is = "LU"
                myColumn = Range(linkCellAddress).Column - 2
                myRow = Range(linkCellAddress).Row
            Case Is = "LD"
                myColumn = Range(linkCellAddress).Column - 2
                myRow = Range(linkCellAddress).Row
            Case Is = "L"
                myColumn = Range(linkCellAddress).Column - 2
                myRow = Range(linkCellAddress).Row
            Case Is = "RU"
                myColumn = Range(linkCellAddress).Column + 2
                myRow = Range(linkCellAddress).Row
            Case Is = "R"
                myColumn = Range(linkCellAddress).Column + 2
                myRow = Range(linkCellAddress).Row
            Case Is = "RD"
                myColumn = Range(linkCellAddress).Column + 2
                myRow = Range(linkCellAddress).Row
            Case Else
                myColumn = Range(linkCellAddress).Column
                myRow = Range(linkCellAddress).Row
        End Select
        
End Select

myRowValue = Cells(myRow, 7).Value
myColumnValue = Cells(1, myColumn).Value

currentScreen = myRowValue & myColumnValue


End Sub


Sub alignScreen()

Dim myColumn, myColumnValue, myRow, myRowValue

myColumn = ActiveCell.Column
myRow = ActiveCell.Row

myRowValue = Cells(myRow, 8).Value
myColumnValue = Cells(2, myColumn).Value

Dim myTopLeft
myTopLeft = ActiveCell.Offset(-myRowValue + 1, -myColumnValue + 1).Address

Application.GoTo ActiveSheet.Range(myTopLeft), True

End Sub

