Attribute VB_Name = "AJ_Triggers"
Public TriggerCel

Global CollidedWith
Global RNDBounceback

Global SwordHit

Global RNDEnemyBounceback1
Global RNDEnemyBounceback2
Global RNDEnemyBounceback3
Global RNDEnemyBounceback4







'###############################################################################
'######      ##     ##     ###    ####     ###      #     ###      #############
'########  ####  ##  ### ###   ##  ##   ##  ##  #####  #   #   #################
'########  ####     #### ###  ######   ######     ###     ####    ##############
'########  ####  ##  ### ###   ##   #   ##   #  #####  ##  ######   ############
'########  ####  ##  #     ##     # ##     # #      #  ##  #       #############
'###############################################################################

Sub EnemyTrigger(triggerCode)

Dim triggerLen
triggerLen = Len(triggerCode)

'S1XXXXETSK01DA1
'S1XXXXETSK01DA10
'S1XXXXETSK01DAB1
'S1XXXXETSK01DAB10
'S1XXXXETSK01DA100
'S1XXXXETSK01DAB100


Dim enemyIndicator
enemyIndicator = Mid(triggerCode, 9, 2)

Dim enemyNumber
enemyNumber = Mid(triggerCode, 11, 2)

Dim triggerDirection
triggerDirection = Mid(triggerCode, 13, 1)

Dim linkDirection
linkDirection = Sheets("Data").Range("C21").Value

'account for different range value lengths (e.g. A1 = 2, A10 = 3, A256 = 4, AA256 = 5)
Select Case triggerLen

    Case Is = 15
        TriggerCel = Right(triggerCode, 2)
        
    Case Is = 16
        TriggerCel = Right(triggerCode, 3)
        
    Case Is = 17
        TriggerCel = Right(triggerCode, 4)
        
    Case Is = 18
        TriggerCel = Right(triggerCode, 5)

End Select

Dim mySub
Dim myEnemy
Dim myShowHide

'If InStr(linkDirection, triggerDirection) <> 0 Then
myShowHide = "show"
'Else
'myShowHide = "hide"
'End If

Select Case enemyIndicator
    Case Is = "SK"
        myEnemy = "skeleton"
    Case Is = "SC"
        myEnemy = "sandcrab"
    Case Is = "SD"
        myEnemy = "soldier"
    Case Is = "BD"
        myEnemy = "bird"
    Case Is = "OC"
        myEnemy = "Octorok"
    'Add more cases here...
    Case Is = "GD"
        myEnemy = "gordo"
    Case Is = "MA"
        myEnemy = "Marin"
    Case Is = "TA"
        myEnemy = "Tarin"
            Case Is = "RC"
        myEnemy = "Raccoon"
    Case Else
        Exit Sub
        
End Select


mySub = myShowHide & myEnemy & enemyNumber


Application.Run mySub


End Sub

Sub RNDEnemyMove(enemyNo)

Dim myEnemyName, myEnemyBehaviour

'Work out which of the four enemies to deal with
    Select Case enemyNo

        Case Is = 1
            myEnemyBehaviour = RNDenemyBehaviour1
    
        Case Is = 2
            myEnemyBehaviour = RNDenemyBehaviour2
 
        Case Is = 3
             myEnemyBehaviour = RNDenemyBehaviour3
             
        Case Is = 4
             myEnemyBehaviour = RNDenemyBehaviour4
             
    End Select


'Work out which behaviour to apply
    Select Case myEnemyBehaviour

        Case Is = "Random"
            Call moveRandom(enemyNo)

        Case Is = "Chase"
            'Call moveChase(enemyNo)
        
        Case Is = "Still"
            Call moveStill(enemyNo)

        Case Is = "StillFollow"
            Call moveStillFollow(enemyNo)
        
        'Add more cases here...
        
        
    End Select


End Sub


'Collisions

Sub enemyCollision(LinkImage, myEnemyImage)

Dim overlap, sideOverlap, topOverlap As Boolean
Dim myCollide

myCollide = ""

Select Case myEnemyImage
        
    Case Is = RNDenemyName1
        myCollide = RNDenemyCanCollide1
    Case Is = RNDenemyName2
        myCollide = RNDenemyCanCollide2
    Case Is = RNDenemyName3
        myCollide = RNDenemyCanCollide3
        'MsgBox "myCollide = " & myCollide
    Case Is = RNDenemyName4
        myCollide = RNDenemyCanCollide4
End Select



If myCollide = "Y" Then
 
    Set enemyImage = ActiveSheet.Shapes(myEnemyImage)
'check sides
    If LinkImage.Left < enemyImage.Left And enemyImage.Left <= LinkImage.Left + LinkImage.Width Then
        sideOverlap = True
    ElseIf enemyImage.Left < LinkImage.Left And LinkImage.Left <= enemyImage.Left + enemyImage.Width Then
        sideOverlap = True
    End If

'check tops
    If LinkImage.Top < enemyImage.Top And enemyImage.Top <= LinkImage.Top + LinkImage.Height Then
        topOverlap = True
    ElseIf enemyImage.Top < LinkImage.Top And LinkImage.Top <= enemyImage.Top + enemyImage.Height Then
        topOverlap = True
    End If

    If sideOverlap And topOverlap Then
        overlap = True
    End If

    If overlap = True Then
        If Sheets("Data").Range("C28").Value = "Y" Then
        'Shield up
            Call pushImage(myEnemyImage)
            
        Else
        'MsgBox ("Collided!")
            RNDBounceback = Sheets("Data").Range("C23").Value
            CollidedWith = myEnemyImage
        End If
    End If

End If
End Sub

Sub BounceBack(LinkImage, enemyImage)

Dim myBounceBackSpeed, moveCellValue

myBounceBackSpeed = Sheets("Data").Range("C23").Value

Select Case moveDir

    Case Is = ""
        Select Case LinkSprite.Name
      
            Case Is = "LinkDown1"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(-1, 2).Value
                If moveCellValue = "" Then
                    LinkImage.Top = LinkImage.Top - myBounceBackSpeed
                End If
            Case Is = "LinkDown2"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(-1, 2).Value
                If moveCellValue = "" Then
                    LinkImage.Top = LinkImage.Top - myBounceBackSpeed
                End If
            
            Case Is = "LinkUp1"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(4, 2).Value
                If moveCellValue = "" Then
                    LinkImage.Top = LinkImage.Top + myBounceBackSpeed
                End If

            Case Is = "LinkUp2"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(4, 2).Value
                If moveCellValue = "" Then
                    LinkImage.Top = LinkImage.Top + myBounceBackSpeed
                End If
            
            Case Is = "LinkLeft1"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, 4).Value
                If moveCellValue = "" Then
                    LinkImage.Left = LinkImage.Left + myBounceBackSpeed
                End If
            
            Case Is = "LinkLeft2"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, 4).Value
                If moveCellValue = "" Then
                    LinkImage.Left = LinkImage.Left + myBounceBackSpeed
                End If
            
            Case Is = "LinkRight1"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, -1).Value
                If moveCellValue = "" Then
                    LinkImage.Left = LinkImage.Left - myBounceBackSpeed
                End If
                        
            Case Is = "LinkRight2"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, -1).Value
                If moveCellValue = "" Then
                    LinkImage.Left = LinkImage.Left - myBounceBackSpeed
                End If
        End Select
        
    Case Else
        Select Case LinkSprite.Name
      
            Case Is = "LinkDown1"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(-1, 2).Value
                If moveCellValue = "" Then
                    LinkImage.Top = LinkImage.Top - myBounceBackSpeed
                End If
            Case Is = "LinkDown2"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(-1, 2).Value
                If moveCellValue = "" Then
                    LinkImage.Top = LinkImage.Top - myBounceBackSpeed
                End If
            
            Case Is = "LinkUp1"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(4, 2).Value
                If moveCellValue = "" Then
                    LinkImage.Top = LinkImage.Top + myBounceBackSpeed
                End If

            Case Is = "LinkUp2"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(4, 2).Value
                If moveCellValue = "" Then
                    LinkImage.Top = LinkImage.Top + myBounceBackSpeed
                End If
            
            Case Is = "LinkLeft1"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, 4).Value
                If moveCellValue = "" Then
                    LinkImage.Left = LinkImage.Left + myBounceBackSpeed
                End If
            
            Case Is = "LinkLeft2"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, 4).Value
                If moveCellValue = "" Then
                    LinkImage.Left = LinkImage.Left + myBounceBackSpeed
                End If
            
            Case Is = "LinkRight1"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, -1).Value
                If moveCellValue = "" Then
                    LinkImage.Left = LinkImage.Left - myBounceBackSpeed
                End If
                        
            Case Is = "LinkRight2"
                moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, -1).Value
                If moveCellValue = "" Then
                    LinkImage.Left = LinkImage.Left - myBounceBackSpeed
                End If
        End Select

End Select

RNDBounceback = RNDBounceback - 1

If RNDBounceback <= 0 Then
    RNDBounceback = ""
End If

End Sub


'####################################################################################


'Show/Hide enemies


'####################################################################################

Sub resetEnemy1()


If RNDenemyName1 <> "" Then

Set myFrame1 = ActiveSheet.Shapes(RNDenemyFrame1_1)
Set myFrame2 = ActiveSheet.Shapes(RNDenemyFrame1_2)
myFrame1.Rotation = 0
myFrame2.Rotation = 0
myFrame1.Visible = False
myFrame2.Visible = False

RNDenemyName1 = ""
RNDenemyFrame1_1 = ""
RNDenemyFrame1_2 = ""
RNDenemyInitialCount1 = ""
RNDenemyCount1 = ""

RNDenemySpeed1 = ""
RNDenemyBehaviour1 = ""
RNDenemyChangeRotation1 = ""

RNDenemyCanShoot1 = ""
RNDenemyChargeSpeed1 = ""

RNDenemyCanCollide1 = ""
RNDenemyCollisionDamage1 = ""
RNDenemyShootDamage1 = ""
RNDenemyChargeDamage1 = ""
RNDenemyLife1 = ""
End If

End Sub

Sub resetEnemy2()

If RNDenemyName2 <> "" Then

Set myFrame1 = ActiveSheet.Shapes(RNDenemyFrame2_1)
Set myFrame2 = ActiveSheet.Shapes(RNDenemyFrame2_2)
myFrame1.Rotation = 0
myFrame2.Rotation = 0
myFrame1.Visible = False
myFrame2.Visible = False

RNDenemyName2 = ""
RNDenemyFrame2_1 = ""
RNDenemyFrame2_2 = ""
RNDenemyInitialCount2 = ""
RNDenemyCount2 = ""

RNDenemySpeed2 = ""
RNDenemyBehaviour2 = ""
RNDenemyChangeRotation2 = ""

RNDenemyCanShoot2 = ""
RNDenemyChargeSpeed2 = ""

RNDenemyCanCollide2 = ""
RNDenemyCollisionDamage2 = ""
RNDenemyShootDamage2 = ""
RNDenemyChargeDamage2 = ""
RNDenemyLife2 = ""
End If

End Sub
Sub resetEnemy3()

If RNDenemyName3 <> "" Then

Set myFrame1 = ActiveSheet.Shapes(RNDenemyFrame3_1)
Set myFrame2 = ActiveSheet.Shapes(RNDenemyFrame3_2)
myFrame1.Rotation = 0
myFrame2.Rotation = 0
myFrame1.Visible = False
myFrame2.Visible = False


RNDenemyName3 = ""
RNDenemyFrame3_1 = ""
RNDenemyFrame3_2 = ""
RNDenemyInitialCount3 = ""
RNDenemyCount3 = ""

RNDenemySpeed3 = ""
RNDenemyBehaviour3 = ""
RNDenemyChangeRotation3 = ""

RNDenemyCanShoot3 = ""
RNDenemyChargeSpeed3 = ""

RNDenemyCanCollide3 = ""
RNDenemyCollisionDamage3 = ""
RNDenemyShootDamage3 = ""
RNDenemyChargeDamage3 = ""
RNDenemyLife3 = ""
End If

End Sub

Sub resetEnemy4()

If RNDenemyName4 <> "" Then

Set myFrame1 = ActiveSheet.Shapes(RNDenemyFrame4_1)
Set myFrame2 = ActiveSheet.Shapes(RNDenemyFrame4_2)
myFrame1.Rotation = 0
myFrame2.Rotation = 0
myFrame1.Visible = False
myFrame2.Visible = False

RNDenemyName4 = ""
RNDenemyFrame4_1 = ""
RNDenemyFrame4_2 = ""
RNDenemyInitialCount4 = ""
RNDenemyCount4 = ""

RNDenemySpeed4 = ""
RNDenemyBehaviour4 = ""
RNDenemyChangeRotation4 = ""

RNDenemyCanShoot4 = ""
RNDenemyChargeSpeed4 = ""

RNDenemyCanCollide4 = ""
RNDenemyCollisionDamage4 = ""
RNDenemyShootDamage4 = ""
RNDenemyChargeDamage4 = ""
RNDenemyLife3 = ""
End If

End Sub




'Sword hits

Sub didSwordHit(mySwordImage, myEnemyImage)

Dim overlap, sideOverlap, topOverlap As Boolean

If myEnemyImage = Empty Then
overlap = False
Exit Sub
End If

Set swordImage = mySwordImage
Set enemyImage = ActiveSheet.Shapes(myEnemyImage)

'check sides
If swordImage.Left < enemyImage.Left And enemyImage.Left <= swordImage.Left + swordImage.Width Then
    sideOverlap = True
ElseIf enemyImage.Left < swordImage.Left And swordImage.Left <= enemyImage.Left + enemyImage.Width Then
    sideOverlap = True
End If

'check tops
If swordImage.Top < enemyImage.Top And enemyImage.Top <= swordImage.Top + swordImage.Height Then
    topOverlap = True
ElseIf enemyImage.Top < swordImage.Top And swordImage.Top <= enemyImage.Top + enemyImage.Height Then
    topOverlap = True
End If

If sideOverlap And topOverlap Then
    overlap = True
Else
    overlap = False
End If

If overlap = True Then

    SwordHit = 5
      
    Call enemyIdentify(myEnemyImage)
    
Else
    SwordHit = False
End If


End Sub


Sub enemyIdentify(whichEnemy)

Dim myDir

If moveDir <> "" Then
    myDir = moveDir
Else
    myDir = lastDir
End If

Select Case whichEnemy

    Case Is = RNDenemyFrame1_1
        RNDenemyHit1 = 5
        RNDEnemyBounceback1 = myDir
        RNDenemyLife1 = RNDenemyLife1 - 1
    Case Is = RNDenemyFrame1_2
        RNDenemyHit1 = 5
        RNDEnemyBounceback1 = myDir
        RNDenemyLife1 = RNDenemyLife1 - 1
        
    Case Is = RNDenemyFrame2_1
        RNDenemyHit2 = 5
        RNDEnemyBounceback2 = myDir
        RNDenemyLife2 = RNDenemyLife2 - 1
    Case Is = RNDenemyFrame2_1
        RNDenemyHit2 = 5
        RNDEnemyBounceback2 = myDir
        RNDenemyLife2 = RNDenemyLife2 - 1
        
    Case Is = RNDenemyFrame3_1
        RNDenemyHit3 = 5
        RNDEnemyBounceback3 = myDir
        RNDenemyLife3 = RNDenemyLife3 - 1
    Case Is = RNDenemyFrame3_1
        RNDenemyHit3 = 5
        RNDEnemyBounceback3 = myDir
        RNDenemyLife3 = RNDenemyLife3 - 1
                
    Case Is = RNDenemyFrame4_1
        RNDenemyHit4 = 5
        RNDEnemyBounceback4 = myDir
        RNDenemyLife4 = RNDenemyLife4 - 1
        
    Case Is = RNDenemyFrame4_1
        RNDenemyHit4 = 5
        RNDEnemyBounceback4 = myDir
        RNDenemyLife4 = RNDenemyLife4 - 1
        
End Select


End Sub


Sub enemyBounceBack(whichEnemy)

Select Case whichEnemy

    Case Is = 1

        Select Case RNDEnemyBounceback1

            Case Is = "U"
                ActiveSheet.Pictures(RNDenemyName1).Top = ActiveSheet.Pictures(RNDenemyName1).Top - 15

            Case Is = "LU"
                ActiveSheet.Pictures(RNDenemyName1).Top = ActiveSheet.Pictures(RNDenemyName1).Top - 15

            Case Is = "RU"
                ActiveSheet.Pictures(RNDenemyName1).Top = ActiveSheet.Pictures(RNDenemyName1).Top - 15
            
            Case Is = "D"
                ActiveSheet.Pictures(RNDenemyName1).Top = ActiveSheet.Pictures(RNDenemyName1).Top + 15

            Case Is = "LD"
                ActiveSheet.Pictures(RNDenemyName1).Top = ActiveSheet.Pictures(RNDenemyName1).Top + 15
        
            Case Is = "RD"
                ActiveSheet.Pictures(RNDenemyName1).Top = ActiveSheet.Pictures(RNDenemyName1).Top + 15
                
            Case Is = "L"
                ActiveSheet.Pictures(RNDenemyName1).Left = ActiveSheet.Pictures(RNDenemyName1).Left - 15
        
            Case Is = "R"
            ActiveSheet.Pictures(RNDenemyName1).Left = ActiveSheet.Pictures(RNDenemyName1).Left + 15
        
            Case Else
                MsgBox "unknown RNDenemybounceback1"
                

        End Select
        
        RNDenemyHit1 = RNDenemyHit1 - 1
        
        If RNDenemyHit1 = 0 Then
            If RNDenemyLife1 <= 0 Then
                Call killEnemy(1)
            End If
        End If

    Case Is = 2
    
        Select Case RNDEnemyBounceback2

            Case Is = "U"
                ActiveSheet.Pictures(RNDenemyName2).Top = ActiveSheet.Pictures(RNDenemyName2).Top - 15
                'barrier collision detection here...
            Case Is = "LU"
                ActiveSheet.Pictures(RNDenemyName2).Top = ActiveSheet.Pictures(RNDenemyName2).Top - 15

            Case Is = "RU"
                ActiveSheet.Pictures(RNDenemyName2).Top = ActiveSheet.Pictures(RNDenemyName2).Top - 15
            
            Case Is = "D"
                ActiveSheet.Pictures(RNDenemyName2).Top = ActiveSheet.Pictures(RNDenemyName2).Top + 15

            Case Is = "LD"
                ActiveSheet.Pictures(RNDenemyName2).Top = ActiveSheet.Pictures(RNDenemyName2).Top + 15
        
            Case Is = "RD"
                ActiveSheet.Pictures(RNDenemyName2).Top = ActiveSheet.Pictures(RNDenemyName2).Top + 15
                
            Case Is = "L"
                ActiveSheet.Pictures(RNDenemyName2).Left = ActiveSheet.Pictures(RNDenemyName2).Left - 15
        
            Case Is = "R"
                ActiveSheet.Pictures(RNDenemyName2).Left = ActiveSheet.Pictures(RNDenemyName2).Left + 15
        
            Case Else
                MsgBox "unknown RNDenemybounceback2"

        End Select

    RNDenemyHit2 = RNDenemyHit2 - 1
    
        If RNDenemyHit2 = 0 Then
        
            If RNDenemyLife2 <= 0 Then
                Call killEnemy(2)
            End If
            
        End If
        
    Case Is = 3
    
        Select Case RNDEnemyBounceback3

            Case Is = "U"
                ActiveSheet.Pictures(RNDenemyName3).Top = ActiveSheet.Pictures(RNDenemyName3).Top - 15

            Case Is = "LU"
                ActiveSheet.Pictures(RNDenemyName3).Top = ActiveSheet.Pictures(RNDenemyName3).Top - 15

            Case Is = "RU"
                ActiveSheet.Pictures(RNDenemyName3).Top = ActiveSheet.Pictures(RNDenemyName3).Top - 15
            
            Case Is = "D"
                ActiveSheet.Pictures(RNDenemyName3).Top = ActiveSheet.Pictures(RNDenemyName3).Top + 15

            Case Is = "LD"
                ActiveSheet.Pictures(RNDenemyName3).Top = ActiveSheet.Pictures(RNDenemyName3).Top + 15
        
            Case Is = "RD"
                ActiveSheet.Pictures(RNDenemyName3).Top = ActiveSheet.Pictures(RNDenemyName3).Top + 15
                
            Case Is = "L"
                ActiveSheet.Pictures(RNDenemyName3).Left = ActiveSheet.Pictures(RNDenemyName3).Left - 15
        
            Case Is = "R"
                ActiveSheet.Pictures(RNDenemyName3).Left = ActiveSheet.Pictures(RNDenemyName3).Left + 15
        
            Case Else
                MsgBox "unknown RNDenemybounceback3"

        End Select

    RNDenemyHit3 = RNDenemyHit3 - 1
    
        If RNDenemyHit3 = 0 Then
        
            If RNDenemyLife3 <= 0 Then
                Call killEnemy(3)
            End If
            
        End If
        
     Case Is = 4
    
        Select Case RNDEnemyBounceback4

            Case Is = "U"
                ActiveSheet.Pictures(RNDenemyName4).Top = ActiveSheet.Pictures(RNDenemyName4).Top - 15

            Case Is = "LU"
                ActiveSheet.Pictures(RNDenemyName4).Top = ActiveSheet.Pictures(RNDenemyName4).Top - 15

            Case Is = "RU"
                ActiveSheet.Pictures(RNDenemyName4).Top = ActiveSheet.Pictures(RNDenemyName4).Top - 15
            
            Case Is = "D"
                ActiveSheet.Pictures(RNDenemyName4).Top = ActiveSheet.Pictures(RNDenemyName4).Top + 15

            Case Is = "LD"
                ActiveSheet.Pictures(RNDenemyName4).Top = ActiveSheet.Pictures(RNDenemyName4).Top + 15
        
            Case Is = "RD"
                ActiveSheet.Pictures(RNDenemyName4).Top = ActiveSheet.Pictures(RNDenemyName4).Top + 15
                
            Case Is = "L"
                ActiveSheet.Pictures(RNDenemyName4).Left = ActiveSheet.Pictures(RNDenemyName4).Left - 15
        
            Case Is = "R"
                ActiveSheet.Pictures(RNDenemyName4).Left = ActiveSheet.Pictures(RNDenemyName4).Left + 15
        
            Case Else
                MsgBox "unknown RNDenemybounceback4"

        End Select

    RNDenemyHit4 = RNDenemyHit4 - 1
    
        If RNDenemyHit4 = 0 Then
        
            If RNDenemyLife4 <= 0 Then
                Call killEnemy(4)
            End If
            
        End If
        
End Select



End Sub


Sub killEnemy(enemyNumber)

Dim enemyName, myNumber

Select Case enemyNumber

    Case Is = 1
        enemyName = RNDenemyName1
        myNumber = "1"
    Case Is = 2
        enemyName = RNDenemyName2
        myNumber = "2"
    Case Is = 3
        enemyName = RNDenemyName3
        myNumber = "3"
    Case Is = 4
        enemyName = RNDenemyName4
        myNumber = "4"

End Select

Call explosion(enemyName, myNumber)


End Sub

Sub explosion(picPosition, enemyNumber)


ActiveSheet.Pictures("Explosion1").Top = ActiveSheet.Pictures(picPosition).Top
ActiveSheet.Pictures("Explosion1").Left = ActiveSheet.Pictures(picPosition).Left

ActiveSheet.Pictures("Explosion2").Top = ActiveSheet.Pictures(picPosition).Top
ActiveSheet.Pictures("Explosion2").Left = ActiveSheet.Pictures(picPosition).Left

ActiveSheet.Pictures("Explosion3").Top = ActiveSheet.Pictures(picPosition).Top - 5
ActiveSheet.Pictures("Explosion3").Left = ActiveSheet.Pictures(picPosition).Left - 5

    Select Case enemyNumber

        Case Is = 1
            Call resetEnemy1
        Case Is = 2
            Call resetEnemy2
        Case Is = 3
           Call resetEnemy3
        Case Is = 4
            Call resetEnemy4
        Case Else
        
    End Select

ActiveSheet.Pictures("Explosion1").Visible = True
Range("A1").Copy Range("A2")
Sleep 8

ActiveSheet.Pictures("Explosion1").Visible = False
ActiveSheet.Pictures("Explosion2").Visible = True
Range("A1").Copy Range("A2")
Sleep 8

ActiveSheet.Pictures("Explosion2").Visible = False
ActiveSheet.Pictures("Explosion3").Visible = True
Range("A1").Copy Range("A2")
Sleep 8

ActiveSheet.Pictures("Explosion3").Visible = False
Range("A1").Copy Range("A2")

'call random item drop here

End Sub


Sub pushImage(myEnemyImage)

Dim pushDir

Select Case moveDir

    Case Is = "U"
        ActiveSheet.Pictures(myEnemyImage).Top = ActiveSheet.Pictures(myEnemyImage).Top - 5
    Case Is = "LU"
        ActiveSheet.Pictures(myEnemyImage).Top = ActiveSheet.Pictures(myEnemyImage).Top - 5
    Case Is = "RU"
        ActiveSheet.Pictures(myEnemyImage).Top = ActiveSheet.Pictures(myEnemyImage).Top - 5
    Case Is = "D"
        ActiveSheet.Pictures(myEnemyImage).Top = ActiveSheet.Pictures(myEnemyImage).Top + 5
    Case Is = "LD"
        ActiveSheet.Pictures(myEnemyImage).Top = ActiveSheet.Pictures(myEnemyImage).Top + 5
    Case Is = "RD"
        ActiveSheet.Pictures(myEnemyImage).Top = ActiveSheet.Pictures(myEnemyImage).Top + 5
    Case Is = "L"
        ActiveSheet.Pictures(myEnemyImage).Left = ActiveSheet.Pictures(myEnemyImage).Left - 5
    Case Is = "R"
        ActiveSheet.Pictures(myEnemyImage).Left = ActiveSheet.Pictures(myEnemyImage).Left + 5
End Select

End Sub





