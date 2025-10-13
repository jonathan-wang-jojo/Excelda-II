'Attribute VB_Name = "AHa_EnemyAI"
Option Explicit

'###################################################################################
'                              ENEMY AI SYSTEM
'###################################################################################
' All AI behaviors are now handled in EnemyManager class
' This module contains only the projectile system
'###################################################################################

' Projectile variables (will be refactored into Projectile class later)
Global projectileName1 As String
Global projectileSpeed1 As Long
Global projectileBehaviour1 As String
Global projectileDir1 As String

'###################################################################################
'                              PROJECTILE SYSTEM
'###################################################################################
' TODO: Refactor into Projectile class

Sub shoot(enemyName As String)
    ' Enemy shoots projectile
    On Error Resume Next
    
    Select Case enemyName
        Case "Octorok1F1"
            Call showCannonball1
        Case "Octorok2F1"
            ' Add other projectiles here
    End Select
End Sub

Sub showCannonball1()
    ' Show projectile 1
    On Error Resume Next
    
    projectileName1 = Sheets(SHEET_DATA).Range("B34").Value
    projectileSpeed1 = Sheets(SHEET_DATA).Range("G34").Value
    projectileBehaviour1 = Sheets(SHEET_DATA).Range("J34").Value
    projectileDir1 = Sheets(SHEET_DATA).Range("F34").Value
    
    ActiveSheet.Shapes("Cannonball1").Visible = True
End Sub

Sub hideCannonball1()
    ' Hide projectile 1
    On Error Resume Next
    
    projectileName1 = ""
    projectileSpeed1 = ""
    projectileBehaviour1 = ""
    projectileDir1 = ""
    
    ActiveSheet.Shapes("Cannonball1").Visible = False
End Sub

Sub projectileMove1()
    ' Update projectile 1
    On Error Resume Next
    
    If projectileName1 <> "" Then
        Select Case projectileBehaviour1
            Case "Straightline"
                Call moveStraight(1)
        End Select
    End If
End Sub

Sub moveStraight(projectileNumber As Long)
    ' Move projectile in straight line
    ' TODO: Implement projectile movement logic
End Sub
