'@Folder("Engine.Core")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' TRIGGER SYSTEM
'═══════════════════════════════════════════════════════════════════════════════
'   XX      = Padding
'   ET      = Action (ET=Enemy Trigger, SE=Special Event, RL=Relocate)
'   SK      = Enemy type (SK=Skeleton, SC=Sandcrab, OC=Octorok, etc.)
'   01      = Enemy slot number (01-04)
'   D       = Trigger direction
'   A1      = Cell address for enemy spawn
'###################################################################################

'###################################################################################
'                              ENEMY TRIGGERS
'###################################################################################

Sub EnemyTrigger(triggerCode As String)
    ' Process enemy spawn triggers
    ' Format from DevConsole: S1XXXXETOC02DR484
    '   Position 7-8: Action (ET)
    '   Position 9-10: Enemy type (OC)
    '   Position 11-12: Enemy slot (02)
    '   Position 13: Direction (D) - legacy, not used
    '   Position 14+: Cell address (R484)
    On Error Resume Next
    
    triggerCode = Trim$(triggerCode)
    If Len(triggerCode) < 14 Then Exit Sub
    
    ' Parse: S1XXXXETOC02DR484
    Dim enemyType As String
    Dim slotNumber As Long
    Dim cellAddress As String
    
    enemyType = Mid$(triggerCode, 9, 2)          ' OC
    slotNumber = CLng(Mid$(triggerCode, 11, 2))  ' 02
    cellAddress = Mid$(triggerCode, 14)          ' R484 (skip direction at pos 13)
    
    Dim gs As GameState
    Set gs = GameStateInstance()
    If Not gs Is Nothing Then
        gs.TriggerCellAddress = cellAddress
    End If
    
    ' Map enemy code to name
    Dim enemyName As String
    Select Case enemyType
        Case "SK": enemyName = "skeleton"
        Case "SC": enemyName = "sandcrab"
        Case "OC": enemyName = "Octorok"
        Case "SS": enemyName = "sandspinner"
        Case "GD": enemyName = "gordo"
        Case "MB": enemyName = "Moblin"
        Case "MA": enemyName = "Marin"
        Case "TA": enemyName = "Tarin"
        Case "RC": enemyName = "Raccoon"
        Case Else: Exit Sub
    End Select
    
    ' TODO: Re-implement enemy spawning using EntityManager
    ' Old EnemyManager system has been removed - migrate to Entity/EntityManager pattern
    Debug.Print "SpawnEnemyTrigger: Legacy enemy system removed - needs migration to EntityManager"
End Sub

Private Function FindEnemyDataRow(enemyType As String) As Long
    ' Map enemy type names to their data rows in the Data sheet
    ' Based on the enemy data structure starting around row 41+
    On Error Resume Next
    
    Select Case LCase(enemyType)
        ' NPCs/Friendlies (rows ~41-44)
        Case "marin": FindEnemyDataRow = 41
        Case "tarin": FindEnemyDataRow = 42
        Case "broomlady": FindEnemyDataRow = 43
        Case "raccoon": FindEnemyDataRow = 44
        
        ' Enemies (rows ~46+)
        Case "skeleton": FindEnemyDataRow = 46
        Case "sandcrab": FindEnemyDataRow = 48
        Case "octorok": FindEnemyDataRow = 50
        Case "sandspinner": FindEnemyDataRow = 52
        Case "gordo": FindEnemyDataRow = 54
        Case "moblin": FindEnemyDataRow = 59
        
        ' Add more as needed
        Case "soldier": FindEnemyDataRow = 47
        Case "bird": FindEnemyDataRow = 49
        
        Case Else: FindEnemyDataRow = 0
    End Select
End Function

'###################################################################################
'                              ENEMY RESET FUNCTIONS
'###################################################################################
' TODO: Re-implement using EntityManager - Legacy EnemyManager removed

Sub ResetAllEnemies()
    ' TODO: Reset enemies through EntityManager when implemented
    Debug.Print "ResetAllEnemies: Legacy enemy system removed - needs migration to EntityManager"
End Sub
