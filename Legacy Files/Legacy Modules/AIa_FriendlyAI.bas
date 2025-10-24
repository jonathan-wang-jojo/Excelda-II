Attribute VB_Name = "AIa_FriendlyAI"
Sub moveStillFollow(enemyNumber)

'work out which character
Dim enemyName, enemyFrame

Select Case enemyNumber

    Case Is = 1
        enemyName = RNDenemyName1
    Case Is = 2
        enemyName = RNDenemyName2
    Case Is = 3
        enemyName = RNDenemyName3
    Case Is = 4
        enemyName = RNDenemyName4
End Select

enemyFrame = enemyName ' the original
enemydir = Right(enemyName, 1)
enemyName = Left(enemyName, Len(enemyName) - 1) 'the root


'work out if Link is higher or lower on the screen

If LinkSprite.Top < ActiveSheet.Shapes(enemyFrame).Top Then
    'look up
    If enemydir <> "U" Then
    
        ActiveSheet.Pictures(enemyFrame).Visible = False
    
        Select Case enemyNumber
        
            Case Is = 1
                RNDenemyName1 = enemyName & "U"
                enemyName = RNDenemyName1
            Case Is = 2
                RNDenemyName2 = enemyName & "U"
                enemyName = RNDenemyName2
            Case Is = 3
                RNDenemyName3 = enemyName & "U"
                enemyName = RNDenemyName3
            Case Is = 4
                RNDenemyName4 = enemyName & "U"
                enemyName = RNDenemyName4
            
        End Select
    
        ActiveSheet.Pictures(enemyName).Visible = True
        Range("A1").Copy Range("A2")
    End If

Else
    
    If LinkSprite.Top > ActiveSheet.Shapes(enemyFrame).Top Then
    
        'look down
        If LinkSprite.Top > ActiveSheet.Shapes(enemyFrame).Top + 60 Then

            If enemydir <> "D" Then
    
                ActiveSheet.Pictures(enemyFrame).Visible = False
    
                Select Case enemyNumber
        
                    Case Is = 1
                        RNDenemyName1 = enemyName & "D"
                        enemyName = RNDenemyName1
                    Case Is = 2
                        RNDenemyName2 = enemyName & "D"
                        enemyName = RNDenemyName2
                    Case Is = 3
                        RNDenemyName3 = enemyName & "D"
                        enemyName = RNDenemyName3
                    Case Is = 4
                        RNDenemyName4 = enemyName & "D"
                        enemyName = RNDenemyName4
                End Select
    
                ActiveSheet.Pictures(enemyName).Visible = True
                Range("A1").Copy Range("A2")
            End If
        Else
        
        'look left or right
            If LinkSprite.Left < ActiveSheet.Shapes(enemyFrame).Left Then
                If enemydir <> "L" Then
                
                    ActiveSheet.Pictures(enemyFrame).Visible = False
                    
                    Select Case enemyNumber
        
                        Case Is = 1
                            RNDenemyName1 = enemyName & "L"
                            enemyName = RNDenemyName1
                        Case Is = 2
                            RNDenemyName2 = enemyName & "L"
                            enemyName = RNDenemyName2
                        Case Is = 3
                            RNDenemyName3 = enemyName & "L"
                            enemyName = RNDenemyName3
                        Case Is = 4
                            RNDenemyName4 = enemyName & "L"
                            enemyName = RNDenemyName4
                    End Select
                    
                    ActiveSheet.Pictures(enemyName).Visible = True
                    Range("A1").Copy Range("A2")
                End If
            
            ElseIf LinkSprite.Left > ActiveSheet.Shapes(enemyFrame).Left + 30 Then
                'Look right
                If enemydir <> "R" Then
                
                    ActiveSheet.Pictures(enemyFrame).Visible = False
                    
                    Select Case enemyNumber
        
                        Case Is = 1
                            RNDenemyName1 = enemyName & "R"
                            enemyName = RNDenemyName1
                        Case Is = 2
                            RNDenemyName2 = enemyName & "R"
                            enemyName = RNDenemyName2
                        Case Is = 3
                            RNDenemyName3 = enemyName & "R"
                            enemyName = RNDenemyName3
                        Case Is = 4
                            RNDenemyName4 = enemyName & "R"
                            enemyName = RNDenemyName4
                    End Select
                    
                ActiveSheet.Pictures(enemyName).Visible = True
                Range("A1").Copy Range("A2")
                End If
                
            End If
            
        End If
        
    End If
    
End If

End Sub
