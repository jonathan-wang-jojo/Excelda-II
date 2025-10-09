'Attribute VB_Name = "AK_Objects"
'###################################################################################
'#
'#
'#      OBJECTS - Bushes, grass, rocks etc
'#
'#
'###################################################################################

Sub swordHitBush(swordImage)


Dim currentBush, bushNumber, bush1, bush2, bush3, bush4, bush5, bush6, bush7, bush8, bush9, bush10, bush11, bush13
Dim bush14, bush15, bush16, bush17, bush18, bush19, bush20, bush21, bush22, bush23, bush24, bush25, bush26, bush27, bush28, bush29, bush30

bushNumber = 1

For N = 1 To 30

currentBush = "Bush" & bushNumber

Dim overlap, sideOverlap, topOverlap As Boolean

overlap = False
sideOverlap = False
topOverlap = False

Set bushImage = ActiveSheet.Pictures(currentBush)

'check sides
If swordImage.Left < bushImage.Left And bushImage.Left <= swordImage.Left + swordImage.Width Then
    sideOverlap = True
ElseIf bushImage.Left < swordImage.Left And swordImage.Left <= bushImage.Left + bushImage.Width Then
    sideOverlap = True
End If

'check tops
If swordImage.Top < bushImage.Top And bushImage.Top <= swordImage.Top + swordImage.Height Then
    topOverlap = True
ElseIf bushImage.Top < swordImage.Top And swordImage.Top <= bushImage.Top + bushImage.Height Then
    topOverlap = True
End If

If sideOverlap And topOverlap Then
    overlap = True
End If

If overlap = True Then

    Select Case bushNumber
 
        Case Is = 1
            bush1 = True
        Case Is = 2
            bush2 = True
        Case Is = 3
            bush3 = True
        Case Is = 4
            bush4 = True
        Case Is = 5
            bush5 = True
        Case Is = 6
            bush6 = True
        Case Is = 7
            bush7 = True
        Case Is = 8
            bush8 = True
        Case Is = 9
            bush9 = True
        Case Is = 10
            bush10 = True
        Case Is = 11
            bush11 = True
        Case Is = 12
            bush12 = True
        Case Is = 13
            bush13 = True
        Case Is = 14
            bush14 = True
        Case Is = 15
            bush15 = True
        Case Is = 16
            bush16 = True
        Case Is = 17
            bush17 = True
        Case Is = 18
            bush18 = True
        Case Is = 19
            bush19 = True
        Case Is = 20
            bush20 = True
        Case Is = 21
            bush21 = True
        Case Is = 22
            bush22 = True
        Case Is = 23
            bush23 = True
        Case Is = 24
            bush24 = True
        Case Is = 25
            bush25 = True
        Case Is = 26
            bush26 = True
        Case Is = 27
            bush27 = True
        Case Is = 28
            bush28 = True
        Case Is = 29
            bush29 = True
        Case Is = 30
            bush30 = True
        Case Else
            MsgBox "unknown Bush - Macro:screenSetUps > swordHitBush"
    End Select
 
End If

bushNumber = bushNumber + 1

Next N

If bush1 = True Then
    Call resetBushes("Bush1")
End If

If bush2 = True Then
    Call resetBushes("Bush2")
End If

If bush3 = True Then
    Call resetBushes("Bush3")
End If

If bush4 = True Then
    Call resetBushes("Bush4")
End If

If bush5 = True Then
    Call resetBushes("Bush5")
End If

If bush6 = True Then
    Call resetBushes("Bush6")
End If

If bush7 = True Then
    Call resetBushes("Bush7")
End If

If bush8 = True Then
    Call resetBushes("Bush8")
End If

If bush9 = True Then
    Call resetBushes("Bush9")
End If

If bush10 = True Then
    Call resetBushes("Bush10")
End If

If bush11 = True Then
    Call resetBushes("Bush11")
End If

If bush12 = True Then
    Call resetBushes("Bush12")
End If

If bush13 = True Then
    Call resetBushes("Bush13")
End If

If bush14 = True Then
    Call resetBushes("Bush14")
End If

If bush15 = True Then
    Call resetBushes("Bush15")
End If

If bush16 = True Then
    Call resetBushes("Bush16")
End If

If bush17 = True Then
    Call resetBushes("Bush17")
End If

If bush18 = True Then
    Call resetBushes("Bush18")
End If

If bush19 = True Then
    Call resetBushes("Bush19")
End If

If bush20 = True Then
    Call resetBushes("Bush20")
End If

If bush21 = True Then
    Call resetBushes("Bush21")
End If

If bush22 = True Then
    Call resetBushes("Bush22")
End If

If bush23 = True Then
    Call resetBushes("Bush23")
End If

If bush24 = True Then
    Call resetBushes("Bush24")
End If

If bush25 = True Then
    Call resetBushes("Bush25")
End If

If bush26 = True Then
    Call resetBushes("Bush26")
End If

If bush27 = True Then
    Call resetBushes("Bush27")
End If

If bush28 = True Then
    Call resetBushes("Bush28")
End If

If bush29 = True Then
    Call resetBushes("Bush29")
End If

If bush30 = True Then
    Call resetBushes("Bush30")
End If


End Sub


Sub resetBushes(whichBush)


'MsgBox "resetting bushes"

Dim currentBush As String
Dim bushNumber As Integer
Dim bushCell

Select Case whichBush

    Case Is = "All"

        bushNumber = 1

        For N = 1 To 30

            currentBush = "Bush" & bushNumber

            bushCell = ActiveSheet.Pictures(currentBush).TopLeftCell.Address
            ActiveSheet.Range(ActiveSheet.Range(bushCell), ActiveSheet.Range(bushCell).Offset(3, 5)).Value = ""
            ActiveSheet.Shapes(currentBush).Visible = False
    
            bushNumber = bushNumber + 1

        Next N

    Case Else
        'currentBush = whichBush
        bushCell = ActiveSheet.Pictures(whichBush).TopLeftCell.Address
        ActiveSheet.Range(ActiveSheet.Range(bushCell), ActiveSheet.Range(bushCell).Offset(3, 5)).Value = ""
        ActiveSheet.Shapes(whichBush).Visible = False

End Select


End Sub


Sub positionObj(ObjName, ObjLocation, cellVal)

    Dim ObjCell
    
    ActiveSheet.Pictures(ObjName).Top = ActiveSheet.Range(ObjLocation).Top
    ActiveSheet.Pictures(ObjName).Left = ActiveSheet.Range(ObjLocation).Left
    ActiveSheet.Pictures(ObjName).Visible = True
    ObjCell = ActiveSheet.Pictures(ObjName).TopLeftCell.Address
    
    ActiveSheet.Range(ActiveSheet.Range(ObjCell), ActiveSheet.Range(ObjCell).Offset(3, 5)).Value = cellVal

End Sub

Sub positionMultiple(ObjectType, Obj1, Obj2, Obj3, Obj4, Obj5, Obj6, Obj7, Obj8, Obj9, Obj10, Obj11, Obj12, Obj13, Obj14, Obj15, Obj16, Obj17, Obj18, Obj19, Obj20, Obj21, Obj22, Obj23, Obj24, Obj25, Obj26, Obj27, Obj28, Obj29, Obj30)

Application.ScreenUpdating = False

ActiveSheet.Pictures(ObjectType & 1).Top = ActiveSheet.Range(Obj1).Top
ActiveSheet.Pictures(ObjectType & 1).Left = ActiveSheet.Range(Obj1).Left

ActiveSheet.Pictures(ObjectType & 2).Top = ActiveSheet.Range(Obj2).Top
ActiveSheet.Pictures(ObjectType & 2).Left = ActiveSheet.Range(Obj2).Left

ActiveSheet.Pictures(ObjectType & 3).Top = ActiveSheet.Range(Obj3).Top
ActiveSheet.Pictures(ObjectType & 3).Left = ActiveSheet.Range(Obj3).Left

ActiveSheet.Pictures(ObjectType & 4).Top = ActiveSheet.Range(Obj4).Top
ActiveSheet.Pictures(ObjectType & 4).Left = ActiveSheet.Range(Obj4).Left

ActiveSheet.Pictures(ObjectType & 5).Top = ActiveSheet.Range(Obj5).Top
ActiveSheet.Pictures(ObjectType & 5).Left = ActiveSheet.Range(Obj5).Left

ActiveSheet.Pictures(ObjectType & 6).Top = ActiveSheet.Range(Obj6).Top
ActiveSheet.Pictures(ObjectType & 6).Left = ActiveSheet.Range(Obj6).Left

ActiveSheet.Pictures(ObjectType & 7).Top = ActiveSheet.Range(Obj7).Top
ActiveSheet.Pictures(ObjectType & 7).Left = ActiveSheet.Range(Obj7).Left

ActiveSheet.Pictures(ObjectType & 8).Top = ActiveSheet.Range(Obj8).Top
ActiveSheet.Pictures(ObjectType & 8).Left = ActiveSheet.Range(Obj8).Left

ActiveSheet.Pictures(ObjectType & 9).Top = ActiveSheet.Range(Obj9).Top
ActiveSheet.Pictures(ObjectType & 9).Left = ActiveSheet.Range(Obj9).Left

ActiveSheet.Pictures(ObjectType & 10).Top = ActiveSheet.Range(Obj10).Top
ActiveSheet.Pictures(ObjectType & 10).Left = ActiveSheet.Range(Obj10).Left

ActiveSheet.Pictures(ObjectType & 11).Top = ActiveSheet.Range(Obj11).Top
ActiveSheet.Pictures(ObjectType & 11).Left = ActiveSheet.Range(Obj11).Left

ActiveSheet.Pictures(ObjectType & 12).Top = ActiveSheet.Range(Obj12).Top
ActiveSheet.Pictures(ObjectType & 12).Left = ActiveSheet.Range(Obj12).Left

ActiveSheet.Pictures(ObjectType & 13).Top = ActiveSheet.Range(Obj13).Top
ActiveSheet.Pictures(ObjectType & 13).Left = ActiveSheet.Range(Obj13).Left

ActiveSheet.Pictures(ObjectType & 14).Top = ActiveSheet.Range(Obj14).Top
ActiveSheet.Pictures(ObjectType & 14).Left = ActiveSheet.Range(Obj14).Left

ActiveSheet.Pictures(ObjectType & 15).Top = ActiveSheet.Range(Obj15).Top
ActiveSheet.Pictures(ObjectType & 15).Left = ActiveSheet.Range(Obj15).Left

ActiveSheet.Pictures(ObjectType & 16).Top = ActiveSheet.Range(Obj16).Top
ActiveSheet.Pictures(ObjectType & 16).Left = ActiveSheet.Range(Obj16).Left

ActiveSheet.Pictures(ObjectType & 17).Top = ActiveSheet.Range(Obj17).Top
ActiveSheet.Pictures(ObjectType & 17).Left = ActiveSheet.Range(Obj17).Left

ActiveSheet.Pictures(ObjectType & 18).Top = ActiveSheet.Range(Obj18).Top
ActiveSheet.Pictures(ObjectType & 18).Left = ActiveSheet.Range(Obj18).Left

ActiveSheet.Pictures(ObjectType & 19).Top = ActiveSheet.Range(Obj19).Top
ActiveSheet.Pictures(ObjectType & 19).Left = ActiveSheet.Range(Obj19).Left

ActiveSheet.Pictures(ObjectType & 20).Top = ActiveSheet.Range(Obj20).Top
ActiveSheet.Pictures(ObjectType & 20).Left = ActiveSheet.Range(Obj20).Left

ActiveSheet.Pictures(ObjectType & 21).Top = ActiveSheet.Range(Obj21).Top
ActiveSheet.Pictures(ObjectType & 21).Left = ActiveSheet.Range(Obj21).Left

ActiveSheet.Pictures(ObjectType & 22).Top = ActiveSheet.Range(Obj22).Top
ActiveSheet.Pictures(ObjectType & 22).Left = ActiveSheet.Range(Obj22).Left

ActiveSheet.Pictures(ObjectType & 23).Top = ActiveSheet.Range(Obj23).Top
ActiveSheet.Pictures(ObjectType & 23).Left = ActiveSheet.Range(Obj23).Left

ActiveSheet.Pictures(ObjectType & 24).Top = ActiveSheet.Range(Obj24).Top
ActiveSheet.Pictures(ObjectType & 24).Left = ActiveSheet.Range(Obj24).Left

ActiveSheet.Pictures(ObjectType & 25).Top = ActiveSheet.Range(Obj25).Top
ActiveSheet.Pictures(ObjectType & 25).Left = ActiveSheet.Range(Obj25).Left

ActiveSheet.Pictures(ObjectType & 26).Top = ActiveSheet.Range(Obj26).Top
ActiveSheet.Pictures(ObjectType & 26).Left = ActiveSheet.Range(Obj26).Left

ActiveSheet.Pictures(ObjectType & 27).Top = ActiveSheet.Range(Obj27).Top
ActiveSheet.Pictures(ObjectType & 27).Left = ActiveSheet.Range(Obj27).Left

ActiveSheet.Pictures(ObjectType & 28).Top = ActiveSheet.Range(Obj28).Top
ActiveSheet.Pictures(ObjectType & 28).Left = ActiveSheet.Range(Obj28).Left

ActiveSheet.Pictures(ObjectType & 29).Top = ActiveSheet.Range(Obj29).Top
ActiveSheet.Pictures(ObjectType & 29).Left = ActiveSheet.Range(Obj29).Left

ActiveSheet.Pictures(ObjectType & 30).Top = ActiveSheet.Range(Obj30).Top
ActiveSheet.Pictures(ObjectType & 30).Left = ActiveSheet.Range(Obj30).Left

ActiveSheet.Pictures(ObjectType & 1).Visible = True
ActiveSheet.Pictures(ObjectType & 2).Visible = True
ActiveSheet.Pictures(ObjectType & 3).Visible = True
ActiveSheet.Pictures(ObjectType & 4).Visible = True
ActiveSheet.Pictures(ObjectType & 5).Visible = True
ActiveSheet.Pictures(ObjectType & 6).Visible = True
ActiveSheet.Pictures(ObjectType & 7).Visible = True
ActiveSheet.Pictures(ObjectType & 8).Visible = True
ActiveSheet.Pictures(ObjectType & 9).Visible = True
ActiveSheet.Pictures(ObjectType & 10).Visible = True
ActiveSheet.Pictures(ObjectType & 11).Visible = True
ActiveSheet.Pictures(ObjectType & 12).Visible = True
ActiveSheet.Pictures(ObjectType & 13).Visible = True
ActiveSheet.Pictures(ObjectType & 14).Visible = True
ActiveSheet.Pictures(ObjectType & 15).Visible = True
ActiveSheet.Pictures(ObjectType & 16).Visible = True
ActiveSheet.Pictures(ObjectType & 17).Visible = True
ActiveSheet.Pictures(ObjectType & 18).Visible = True
ActiveSheet.Pictures(ObjectType & 19).Visible = True
ActiveSheet.Pictures(ObjectType & 20).Visible = True
ActiveSheet.Pictures(ObjectType & 21).Visible = True
ActiveSheet.Pictures(ObjectType & 22).Visible = True
ActiveSheet.Pictures(ObjectType & 23).Visible = True
ActiveSheet.Pictures(ObjectType & 24).Visible = True
ActiveSheet.Pictures(ObjectType & 25).Visible = True
ActiveSheet.Pictures(ObjectType & 26).Visible = True
ActiveSheet.Pictures(ObjectType & 27).Visible = True
ActiveSheet.Pictures(ObjectType & 28).Visible = True
ActiveSheet.Pictures(ObjectType & 29).Visible = True
ActiveSheet.Pictures(ObjectType & 30).Visible = True

Application.ScreenUpdating = True

End Sub

Sub getHeartPiece(HeartNum)

Dim myCell

myCell = ActiveSheet.Shapes("HeartPiece").TopLeftCell.Address

ActiveSheet.Range(ActiveSheet.Range(myCell), ActiveSheet.Range(myCell).Offset(3, 5)).Value = ""

ActiveSheet.Shapes("HeartPiece").Visible = False

ActiveSheet.Shapes("LinkWin").Top = LinkSprite.Top
ActiveSheet.Shapes("LinkWin").Left = LinkSprite.Left
LinkSprite.Visible = False

ActiveSheet.Shapes("LinkWin").Visible = True
ActiveSheet.Shapes("HeartPiece").Top = ActiveSheet.Shapes("LinkWin").Top - 45
ActiveSheet.Shapes("HeartPiece").Left = ActiveSheet.Shapes("LinkWin").Left
ActiveSheet.Shapes("HeartPiece").Visible = True

Range("A1").Copy Range("A2")
Sleep 2000

'record the find on the data sheet
heartRange = "AB" & HeartNum
Sheets("Data").Range(heartRange).Value = "Y"
Sheets("Data").Range("AB1").Value = Sheets("Data").Range("AB1").Value + 1

'work out which message to show
If Sheets("Data").Range("AB1").Value = 4 Then
    Sheets("Data").Range("C42").Value = Sheets("Data").Range("R8").Value
    DialogueForm.Show
    Sheets("Data").Range("AB1").Value = 0
    'XXX increment life amount here
Else
    Sheets("Data").Range("C42").Value = Sheets("Data").Range("R7").Value
    DialogueForm.Show
End If

'reset link sprite
ActiveSheet.Shapes("HeartPiece").Visible = False
ActiveSheet.Shapes("LinkWin").Visible = False
Range("A1").Copy Range("A2")


End Sub
