'Attribute VB_Name = "AHa_EnemyAI"
'####################################################################################
'####     ##     #  ##  ###  ###  ##  #  ###  ###  ##  #     ###     ################
'####  ##  #  ####  ##  ##    ##  ##  #  ##    ##  ##  #  ##  #  ####################
'####     ##    ##      #  ##  #  ##  #  #  ##  #  ##  #     ###    #################
'####  ##  #  ####  ##  #      ##    ##  ##    ##  ##  #  ##  #####  ################
'####     ##     #  ##  #  ##  ###  ###  ###  ####    ##  ##  #     #################
'####################################################################################

Sub moveRandom(enemyNumber)

'MsgBox ("moveRandom called")

Dim myEnemy, myFrame1, myFrame2, myCount, myInitialCount, myDir, mySpeed, myFacing, myEnemyAddress

'---------------- Set variables to match enemyNumber ------------------------------

Select Case enemyNumber

    Case Is = 1
        myEnemy = RNDenemyName1
        myFrame1 = RNDenemyFrame1_1
        myFrame2 = RNDenemyFrame1_2
        myCount = RNDenemyCount1
        myInitialCount = RNDenemyInitialCount1
        myDir = RNDenemyDir1
        mySpeed = RNDenemySpeed1
        myChangeRotation = RNDenemyChangeRotation1
        myCanShoot = RNDenemyCanShoot1
        
    Case Is = 2
        myEnemy = RNDenemyName2
        myFrame1 = RNDenemyFrame2_1
        myFrame2 = RNDenemyFrame2_2
        myCount = RNDenemyCount2
        myInitialCount = RNDenemyInitialCount2
        myDir = RNDenemyDir2
        mySpeed = RNDenemySpeed2
        myChangeRotation = RNDenemyChangeRotation2
        myCanShoot = RNDenemyCanShoot2
    Case Is = 3
        myEnemy = RNDenemyName3
        myFrame1 = RNDenemyFrame3_1
        myFrame2 = RNDenemyFrame3_2
        myCount = RNDenemyCount3
        myInitialCount = RNDenemyInitialCount3
        myDir = RNDenemyDir3
        mySpeed = RNDenemySpeed3
        myChangeRotation = RNDenemyChangeRotation3
        myCanShoot = RNDenemyCanShoot3
    Case Is = 4
        myEnemy = RNDenemyName4
        myFrame1 = RNDenemyFrame4_1
        myFrame2 = RNDenemyFrame4_2
        myCount = RNDenemyCount4
        myInitialCount = RNDenemyInitialCount4
        myDir = RNDenemyDir4
        mySpeed = RNDenemySpeed4
        myChangeRotation = RNDenemyChangeRotation4
        myCanShoot = RNDenemyCanShoot4

End Select

myEnemyAddress = ActiveSheet.Shapes(myEnemy).TopLeftCell.Address
myFacing = ActiveSheet.Shapes(myEnemy).Rotation

'----------------------------------------------------------------------------------
'-------------------- Check to see how far through a 'cycle' the enemy is ---------

Select Case myCount

    Case Is = ""
        'Do nothing - should never reach this state
    Case Is = "10"
      
        If myFrame1 <> "" Then
      
            Select Case myEnemy
        
                Case Is = myFrame1
                    ActiveSheet.Shapes(myFrame2).Top = ActiveSheet.Shapes(myFrame1).Top
                    ActiveSheet.Shapes(myFrame2).Left = ActiveSheet.Shapes(myFrame1).Left
                    
                    Select Case enemyNumber

                        Case Is = 1
                            RNDenemyName1 = myFrame2
                            ActiveSheet.Shapes(myFrame2).Visible = True
                            ActiveSheet.Shapes(myFrame1).Visible = False
        
                        Case Is = 2
                            RNDenemyName2 = myFrame2
                            ActiveSheet.Shapes(myFrame2).Visible = True
                            ActiveSheet.Shapes(myFrame1).Visible = False
                        Case Is = 3
                            RNDenemyName3 = myFrame2
                            ActiveSheet.Shapes(myFrame2).Visible = True
                            ActiveSheet.Shapes(myFrame1).Visible = False
                        Case Is = 4
                            RNDenemyName4 = myFrame2
                            ActiveSheet.Shapes(myFrame2).Visible = True
                            ActiveSheet.Shapes(myFrame1).Visible = False

                    End Select
                    
                Case Is = myFrame2
                    ActiveSheet.Shapes(myFrame1).Top = ActiveSheet.Shapes(myFrame2).Top
                    ActiveSheet.Shapes(myFrame1).Left = ActiveSheet.Shapes(myFrame2).Left
                    
                    Select Case enemyNumber

                        Case Is = 1
                            RNDenemyName1 = myFrame1
                            ActiveSheet.Shapes(myFrame1).Visible = True
                            ActiveSheet.Shapes(myFrame2).Visible = False
                        Case Is = 2
                            RNDenemyName2 = myFrame1
                            ActiveSheet.Shapes(myFrame1).Visible = True
                            ActiveSheet.Shapes(myFrame2).Visible = False
                        Case Is = 3
                            RNDenemyName3 = myFrame1
                            ActiveSheet.Shapes(myFrame1).Visible = True
                            ActiveSheet.Shapes(myFrame2).Visible = False
                        Case Is = 4
                            RNDenemyName4 = myFrame1
                            ActiveSheet.Shapes(myFrame1).Visible = True
                            ActiveSheet.Shapes(myFrame2).Visible = False
                    End Select
                    
            End Select
            
        End If
      
        myCount = myCount - 1
    Case Is > 0
        'Part way through the cycle, continue counting down
        myCount = myCount - 1
    Case Is = 0
        'finished the cycle, select a new direction and reset the cycle

        'select a new direction and rotate the enemy accordingly
        Dim RNDNo
        RNDNo = Int((5 - 1 + 1) * RND + 1) 'generate random number between 1 and 5
                
        Select Case RNDNo
        
            Case Is = 1
                myDir = "N"
            Case Is = 2
                myDir = "S"
            Case Is = 3
                myDir = "E"
            Case Is = 4
                myDir = "W"
            Case Is = 5
                If myCanShoot <> "" Then
                    Call shoot(myEnemy)
                End If
        End Select
   
        If myChangeRotation = "Y" Then
    
                Select Case myDir
                    Case Is = "S"
                        myRotate = 0
                    Case Is = "N"
                        myRotate = 180
                    Case Is = "W"
                        myRotate = 90
                    Case Is = "E"
                        myRotate = 270
                End Select
                
            ActiveSheet.Shapes(myEnemy).Rotation = myRotate
            ActiveSheet.Shapes(myFrame1).Rotation = myRotate
            ActiveSheet.Shapes(myFrame2).Rotation = myRotate
    End If
    
    
    'reset the cycle
    myCount = myInitialCount
        
End Select



'----------------------------------------------------------------------------------
' --------------------- move the enemy sprite -------------------------------------

'work out movement direction, and collision detection
    Select Case myDir
        
        Case Is = "N"
            If Range(myEnemyAddress).Offset(-1, 1).Value = "" Or Range(myEnemyAddress).Offset(-1, 1).Value = "_\|/_" Then
                ActiveSheet.Shapes(myEnemy).Top = ActiveSheet.Shapes(myEnemy).Top - mySpeed
            Else
                GoTo endLoop

            End If

        Case Is = "S"
            If Range(myEnemyAddress).Offset(4, 1).Value = "" Or Range(myEnemyAddress).Offset(4, 1).Value = "_\|/_" Then
                ActiveSheet.Shapes(myEnemy).Top = ActiveSheet.Shapes(myEnemy).Top + mySpeed
            Else
                GoTo endLoop
            End If
        Case Is = "E"
            If Range(myEnemyAddress).Offset(2, 4).Value = "" Or Range(myEnemyAddress).Offset(2, 4).Value = "_\|/_" Then
                ActiveSheet.Shapes(myEnemy).Left = ActiveSheet.Shapes(myEnemy).Left + mySpeed
            Else
                GoTo endLoop
            End If
        Case Is = "W"
            If Range(myEnemyAddress).Offset(2, -1).Value = "" Or Range(myEnemyAddress).Offset(2, -1).Value = "_\|/_" Then
                ActiveSheet.Shapes(myEnemy).Left = ActiveSheet.Shapes(myEnemy).Left - mySpeed
            Else
                GoTo endLoop
            End If
    End Select

endLoop:

' set global variables to match local ones, ready for next loop through

Select Case enemyNumber

    Case Is = 1
        RNDenemyCount1 = myCount
        RNDenemyDir1 = myDir
        
    Case Is = 2
        RNDenemyCount2 = myCount
        RNDenemyDir2 = myDir
        
    Case Is = 3
        RNDenemyCount3 = myCount
        RNDenemyDir3 = myDir
        
    Case Is = 4
        RNDenemyCount4 = myCount
        RNDenemyDir4 = myDir

End Select


End Sub

Sub moveChase()



End Sub

Sub moveStill(enemyNumber)

'MsgBox "Called Still"

Dim myEnemy, myFrame1, myFrame2, myCount, myInitialCount, myDir, mySpeed, myFacing, myEnemyAddress

Select Case enemyNumber

    Case Is = 1
        myEnemy = RNDenemyName1
        myFrame1 = RNDenemyFrame1_1
        myFrame2 = RNDenemyFrame1_2
        myCount = RNDenemyCount1
        myInitialCount = RNDenemyInitialCount1
        myDir = RNDenemyDir1
        mySpeed = RNDenemySpeed1
        myChangeRotation = RNDenemyChangeRotation1
        myCanShoot = RNDenemyCanShoot1
        
    Case Is = 2
        myEnemy = RNDenemyName2
        myFrame1 = RNDenemyFrame2_1
        myFrame2 = RNDenemyFrame2_2
        myCount = RNDenemyCount2
        myInitialCount = RNDenemyInitialCount2
        myDir = RNDenemyDir2
        mySpeed = RNDenemySpeed2
        myChangeRotation = RNDenemyChangeRotation2
        myCanShoot = RNDenemyCanShoot2
    Case Is = 3
        myEnemy = RNDenemyName3
        myFrame1 = RNDenemyFrame3_1
        myFrame2 = RNDenemyFrame3_2
        myCount = RNDenemyCount3
        myInitialCount = RNDenemyInitialCount3
        myDir = RNDenemyDir3
        mySpeed = RNDenemySpeed3
        myChangeRotation = RNDenemyChangeRotation3
        myCanShoot = RNDenemyCanShoot3
    Case Is = 4
        myEnemy = RNDenemyName4
        myFrame1 = RNDenemyFrame4_1
        myFrame2 = RNDenemyFrame4_2
        myCount = RNDenemyCount4
        myInitialCount = RNDenemyInitialCount4
        myDir = RNDenemyDir4
        mySpeed = RNDenemySpeed4
        myChangeRotation = RNDenemyChangeRotation4
        myCanShoot = RNDenemyCanShoot4
        
End Select

myEnemyAddress = ActiveSheet.Shapes(myEnemy).TopLeftCell.Address
'myFacing = ActiveSheet.Shapes(myEnemy).Rotation

'MsgBox (myCount)

Select Case myCount
    
    Case Is = ""
    
        'Do nothing - should never reach this state
    Case Is = "10"
        'MsgBox "change sprite"
        If myFrame1 <> "" Then
      
            Select Case myEnemy
        
                Case Is = myFrame1
                    ActiveSheet.Shapes(myFrame2).Top = ActiveSheet.Shapes(myFrame1).Top
                    ActiveSheet.Shapes(myFrame2).Left = ActiveSheet.Shapes(myFrame1).Left
                    
                    Select Case enemyNumber

                        Case Is = 1
                            RNDenemyName1 = myFrame2
                            ActiveSheet.Shapes(myFrame2).Visible = True
                            ActiveSheet.Shapes(myFrame1).Visible = False
        
                        Case Is = 2
                            RNDenemyName2 = myFrame2
                            ActiveSheet.Shapes(myFrame2).Visible = True
                            ActiveSheet.Shapes(myFrame1).Visible = False
                            
                        Case Is = 3
                            RNDenemyName3 = myFrame2
                            ActiveSheet.Shapes(myFrame2).Visible = True
                            ActiveSheet.Shapes(myFrame1).Visible = False
                            
                        Case Is = 4
                            RNDenemyName4 = myFrame2
                            ActiveSheet.Shapes(myFrame2).Visible = True
                            ActiveSheet.Shapes(myFrame1).Visible = False
                            
                    End Select
                    
                Case Is = myFrame2
                    ActiveSheet.Shapes(myFrame1).Top = ActiveSheet.Shapes(myFrame2).Top
                    ActiveSheet.Shapes(myFrame1).Left = ActiveSheet.Shapes(myFrame2).Left
                    
                    Select Case enemyNumber

                        Case Is = 1
                            RNDenemyName1 = myFrame1
                            ActiveSheet.Shapes(myFrame1).Visible = True
                            ActiveSheet.Shapes(myFrame2).Visible = False
                        Case Is = 2
                            RNDenemyName2 = myFrame1
                            ActiveSheet.Shapes(myFrame1).Visible = True
                            ActiveSheet.Shapes(myFrame2).Visible = False
                        Case Is = 3
                            RNDenemyName3 = myFrame1
                            ActiveSheet.Shapes(myFrame1).Visible = True
                            ActiveSheet.Shapes(myFrame2).Visible = False
                        Case Is = 4
                            RNDenemyName4 = myFrame1
                            ActiveSheet.Shapes(myFrame1).Visible = True
                            ActiveSheet.Shapes(myFrame2).Visible = False
                    End Select
                    
            End Select
            
        End If
    
        myCount = myCount - 1
    Case Is > 0
        'Part way through the cycle, continue counting down
        myCount = myCount - 1
    Case Is = 0

    myCount = myInitialCount
    
End Select

Select Case enemyNumber

    Case Is = 1
        RNDenemyCount1 = myCount
        RNDenemyDir1 = myDir
        
    Case Is = 2
        RNDenemyCount2 = myCount
        RNDenemyDir2 = myDir
        
    Case Is = 3
        RNDenemyCount3 = myCount
        RNDenemyDir3 = myDir
        
    Case Is = 4
        RNDenemyCount4 = myCount
        RNDenemyDir4 = myDir
End Select

End Sub

Sub shoot(enemyName)

Dim myProjectile

Select Case enemyName

    Case Is = "Octorok1F1"
        showCannonball1
        
    Case Is = "Octorok2F1"
    


End Select

End Sub

Sub showCannonball1()

projectileName1 = Sheets("Data").Range("B34").Value
projectileSpeed1 = Sheets("Data").Range("G34").Value
projectileBehaviour1 = Sheets("Data").Range("J34").Value
projectileDir1 = Sheets("Data").Range("F34").Value

ActiveSheet.Shapes("Cannonball1").Visible = True

End Sub

Sub hideCannonball1()

projectileName1 = ""
projectileSpeed1 = ""
projectileBehaviour1 = ""
projectileDir1 = ""

ActiveSheet.Shapes("Cannonball1").Visible = False

End Sub


Sub projectileMove1()

'msgBox ("EnemyMove1 called")

If projectileName1 <> "" Then

    Select Case projectileBehaviour1

        Case Is = "Straightline"

            Call moveStraight(1)
        
    End Select

End If

End Sub
