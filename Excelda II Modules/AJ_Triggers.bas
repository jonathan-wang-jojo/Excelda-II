'Attribute VB_Name = "AJ_Triggers"
Option Explicit

'###################################################################################
'                              TRIGGER SYSTEM
'###################################################################################
' Trigger Code Format: S[Dir][Fall][XX][Action][EnemyType][EnemyNum][Dir][CellAddr]
' Example: S1XXXXETSK01DA1
'   S       = Scroll indicator
'   1       = Scroll direction (1=Right, 2=Left, 3=Down, 4=Up)
'   XX      = Fall indicator (FL=Fall, JD=Jump Down)
'   XX      = Padding
'   ET      = Action (ET=Enemy Trigger, SE=Special Event, RL=Relocate)
'   SK      = Enemy type (SK=Skeleton, SC=Sandcrab, OC=Octorok, etc.)
'   01      = Enemy slot number (01-04)
'   D       = Trigger direction
'   A1      = Cell address for enemy spawn
'###################################################################################

' Module-level variables
Public TriggerCel As String

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
    
    TriggerCel = cellAddress
    
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
    
    ' Get data row and spawn
    Dim dataRow As Long
    dataRow = FindEnemyDataRow(enemyName)
    If dataRow = 0 Or slotNumber < 1 Or slotNumber > 4 Then Exit Sub
    
    Dim manager As EnemyManager
    Set manager = EnemyManagerInstance()
    manager.SpawnEnemy enemyName, slotNumber, dataRow, cellAddress
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

Sub BounceBack(LinkImage, enemyImage)
    ' Bounce Link back when hit by enemy
    On Error Resume Next
    
    Dim myBounceBackSpeed As Long
    Dim moveCellValue As String
    Dim currentDir As String
    
    ' Get bounce speed and direction from game state
    myBounceBackSpeed = Sheets(SHEET_DATA).Range("C23").Value
    currentDir = Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value
    
    ' Use last direction if no current movement
    If currentDir = "" Then
        Dim gs As GameState
        Set gs = GameStateInstance()
        currentDir = gs.LastDir
    End If
    
    ' Bounce back based on Link's sprite direction
    Select Case LinkSprite.Name
        Case "LinkDown1", "LinkDown2"
            moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(-1, 2).Value
            If moveCellValue = "" Then
                LinkImage.Top = LinkImage.Top - myBounceBackSpeed
            End If
        
        Case "LinkUp1", "LinkUp2"
            moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(4, 2).Value
            If moveCellValue = "" Then
                LinkImage.Top = LinkImage.Top + myBounceBackSpeed
            End If
        
        Case "LinkLeft1", "LinkLeft2"
            moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, 4).Value
            If moveCellValue = "" Then
                LinkImage.Left = LinkImage.Left + myBounceBackSpeed
            End If
        
        Case "LinkRight1", "LinkRight2"
            moveCellValue = Range(LinkImage.TopLeftCell.Address).Offset(2, -1).Value
            If moveCellValue = "" Then
                LinkImage.Left = LinkImage.Left - myBounceBackSpeed
            End If
    End Select
End Sub


'###################################################################################
'                              ENEMY RESET FUNCTIONS
'###################################################################################
' Simplified - just delegates to EnemyManager

Sub ResetAllEnemies()
    ' Reset all enemies through the EnemyManager
    Dim manager As EnemyManager
    Set manager = EnemyManagerInstance()
    
    If Not manager Is Nothing Then
        manager.Reset
    End If
End Sub


'###################################################################################
'                              SWORD HIT DETECTION
'###################################################################################

Sub didSwordHit(mySwordImage As Shape, myEnemyImage As String)
    ' Check if sword hit enemy
    On Error Resume Next

    If myEnemyImage = "" Then Exit Sub

    Dim swordImage As Shape
    Dim enemyImage As Shape
    Dim overlap As Boolean, sideOverlap As Boolean, topOverlap As Boolean

    Set swordImage = mySwordImage
    Set enemyImage = ActiveSheet.Shapes(myEnemyImage)

    ' Check side overlap
    If swordImage.Left < enemyImage.Left And enemyImage.Left <= swordImage.Left + swordImage.Width Then
        sideOverlap = True
    ElseIf enemyImage.Left < swordImage.Left And swordImage.Left <= enemyImage.Left + enemyImage.Width Then
        sideOverlap = True
    End If

    ' Check top overlap
    If swordImage.Top < enemyImage.Top And enemyImage.Top <= swordImage.Top + swordImage.Height Then
        topOverlap = True
    ElseIf enemyImage.Top < swordImage.Top And swordImage.Top <= enemyImage.Top + enemyImage.Height Then
        topOverlap = True
    End If

    overlap = (sideOverlap And topOverlap)

    If overlap Then
        Call HitEnemyByName(myEnemyImage)
    End If
End Sub

Private Sub HitEnemyByName(enemyName As String)
    ' Apply sword hit to enemy by finding which slot it's in
    On Error Resume Next
    
    ' Get direction from game state
    Dim gs As GameState
    Set gs = GameStateInstance()
    
    Dim myDir As String
    myDir = Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value
    If myDir = "" Then myDir = gs.LastDir
    
    Dim manager As EnemyManager
    Set manager = EnemyManagerInstance()
    
    ' Find which enemy slot this is
    Dim i As Long
    For i = 1 To 4
        If manager.IsActive(i) Then
            Dim enemy As enemy
            Set enemy = manager.enemy(i)
            If enemy.Frame1 = enemyName Or enemy.Frame2 = enemyName Then
                manager.HitEnemy i, myDir
                Exit For
            End If
        End If
    Next i
End Sub