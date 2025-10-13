Option Explicit

'###################################################################################
'                              ENEMY MANAGER CLASS
'###################################################################################
' Manages all enemies in the game - up to 4 concurrent enemies
'###################################################################################

Private m_Enemies(1 To 4) As Enemy
Private m_BouncebackDir(1 To 4) As String

'###################################################################################
'                              LIFECYCLE
'###################################################################################

Private Sub Class_Initialize()
    Dim i As Long
    For i = 1 To 4
        Set m_Enemies(i) = New Enemy
        m_BouncebackDir(i) = ""
    Next i
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Reset()
    ' Reset all enemies
    Dim i As Long
    For i = 1 To 4
        If Not m_Enemies(i) Is Nothing Then
            m_Enemies(i).Hide
        End If
        m_BouncebackDir(i) = ""
    Next i
End Sub

Public Sub Destroy()
    ' Clean up resources
    Dim i As Long
    For i = 1 To 4
        If Not m_Enemies(i) Is Nothing Then
            m_Enemies(i).Hide
            Set m_Enemies(i) = Nothing
        End If
    Next i
End Sub

'###################################################################################
'                              CORE PROCESSING
'###################################################################################

Public Sub ProcessEnemy(ByVal index As Long, ByVal linkSprite As Object)
    ' Main enemy update - called every frame for each active enemy
    On Error Resume Next
    
    If index < 1 Or index > 4 Then Exit Sub
    If Not m_Enemies(index).IsActive Then Exit Sub
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    ' Check if enemy is being hit
    If enemy.Hit > 0 Then
        Call ApplyBounceBack(index)
        enemy.Hit = enemy.Hit - 1
        
        ' Check if enemy is dead
        If enemy.Hit = 0 And enemy.Life <= 0 Then
            Call KillEnemy(index)
        End If
    Else
        ' Normal behavior
        Call UpdateEnemy(index)
        Call CheckCollision(index, linkSprite)
    End If
End Sub

Private Sub UpdateEnemy(ByVal index As Long)
    ' Update enemy based on behavior
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    Select Case enemy.Behavior
        Case "Random"
            Call UpdateRandomMovement(index)
        Case "Still"
            Call UpdateStillAnimation(index)
        Case "StillFollow"
            Call UpdateStillFollow(index)
        Case "Chase"
            ' Call UpdateChase(index)
    End Select
End Sub

'###################################################################################
'                              ENEMY BEHAVIORS
'###################################################################################

Private Sub UpdateRandomMovement(ByVal index As Long)
    ' Random movement AI
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    Dim enemyAddress As String
    enemyAddress = enemy.GetAddress()
    
    ' Update animation counter
    If enemy.Count = 10 Then
        enemy.SwapFrame
        enemy.Count = enemy.Count - 1
    ElseIf enemy.Count > 0 Then
        enemy.Count = enemy.Count - 1
    ElseIf enemy.Count = 0 Then
        ' Pick new direction
        Dim rndNo As Long
        rndNo = Int((5 - 1 + 1) * Rnd + 1)
        
        Select Case rndNo
            Case 1: enemy.Direction = "N"
            Case 2: enemy.Direction = "S"
            Case 3: enemy.Direction = "E"
            Case 4: enemy.Direction = "W"
            Case 5
                If enemy.CanShoot <> "" Then
                    Call ShootProjectile(index)
                End If
        End Select
        
        ' Update rotation if needed
        If enemy.ChangeRotation = "Y" Then
            Dim rotation As Long
            Select Case enemy.Direction
                Case "S": rotation = 0
                Case "N": rotation = 180
                Case "W": rotation = 90
                Case "E": rotation = 270
            End Select
            
            Dim shp As Shape
            Set shp = enemy.GetShape()
            If Not shp Is Nothing Then
                shp.Rotation = rotation
                If enemy.Frame1 <> "" Then ActiveSheet.Shapes(enemy.Frame1).Rotation = rotation
                If enemy.Frame2 <> "" Then ActiveSheet.Shapes(enemy.Frame2).Rotation = rotation
            End If
        End If
        
        enemy.Count = enemy.InitialCount
    End If
    
    ' Move enemy
    Call MoveEnemy(index, enemy.Direction, enemy.Speed, enemyAddress)
End Sub

Private Sub UpdateStillAnimation(ByVal index As Long)
    ' Still enemy with animation
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    If enemy.Count = 10 Then
        enemy.SwapFrame
        enemy.Count = enemy.Count - 1
    ElseIf enemy.Count > 0 Then
        enemy.Count = enemy.Count - 1
    ElseIf enemy.Count = 0 Then
        enemy.Count = enemy.InitialCount
    End If
End Sub

Private Sub UpdateStillFollow(ByVal index As Long)
    ' Still enemy that follows/faces player
    On Error Resume Next
    Call UpdateStillAnimation(index)
    ' TODO: Add facing logic
End Sub

Private Sub MoveEnemy(ByVal index As Long, ByVal Direction As String, ByVal Speed As Long, ByVal address As String)
    ' Move enemy in specified direction with collision detection
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    Dim shp As Shape
    Set shp = enemy.GetShape()
    If shp Is Nothing Then Exit Sub
    
    Dim canMove As Boolean
    canMove = False
    
    Select Case Direction
        Case "N"
            If Range(address).Offset(-1, 1).Value = "" Or _
               Range(address).Offset(-1, 1).Value = "_\|/_" Then
                canMove = True
                shp.Top = shp.Top - Speed
            End If
        Case "S"
            If Range(address).Offset(4, 1).Value = "" Or _
               Range(address).Offset(4, 1).Value = "_\|/_" Then
                canMove = True
                shp.Top = shp.Top + Speed
            End If
        Case "E"
            If Range(address).Offset(2, 4).Value = "" Or _
               Range(address).Offset(2, 4).Value = "_\|/_" Then
                canMove = True
                shp.Left = shp.Left + Speed
            End If
        Case "W"
            If Range(address).Offset(2, -1).Value = "" Or _
               Range(address).Offset(2, -1).Value = "_\|/_" Then
                canMove = True
                shp.Left = shp.Left - Speed
            End If
    End Select
End Sub

'###################################################################################
'                              COLLISION & COMBAT
'###################################################################################

Private Sub CheckCollision(ByVal index As Long, ByVal linkSprite As Object)
    ' Check collision with Link
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    If enemy.CanCollide <> "Y" Then Exit Sub
    
    Dim shp As Shape
    Set shp = enemy.GetShape()
    If shp Is Nothing Then Exit Sub
    
    ' Check overlap
    Dim sideOverlap As Boolean, topOverlap As Boolean
    
    If linkSprite.Left < shp.Left And shp.Left <= linkSprite.Left + linkSprite.Width Then
        sideOverlap = True
    ElseIf shp.Left < linkSprite.Left And linkSprite.Left <= shp.Left + shp.Width Then
        sideOverlap = True
    End If
    
    If linkSprite.Top < shp.Top And shp.Top <= linkSprite.Top + linkSprite.Height Then
        topOverlap = True
    ElseIf shp.Top < linkSprite.Top And linkSprite.Top <= shp.Top + shp.Height Then
        topOverlap = True
    End If
    
    If sideOverlap And topOverlap Then
        If Sheets(SHEET_DATA).Range(RANGE_SHIELD_STATE).Value = "Y" Then
            ' Shield up - push enemy
            Call PushEnemy(index)
        Else
            ' Hit Link
            RNDBounceback = Sheets(SHEET_DATA).Range("C23").Value
            CollidedWith = enemy.Name
        End If
    End If
End Sub

Public Sub HitEnemy(ByVal index As Long, ByVal Direction As String)
    ' Apply hit to enemy
    On Error Resume Next
    
    If index < 1 Or index > 4 Then Exit Sub
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    If Not enemy.IsActive Then Exit Sub
    
    enemy.Hit = 5
    enemy.Life = enemy.Life - 1
    m_BouncebackDir(index) = Direction
End Sub

Private Sub ApplyBounceBack(ByVal index As Long)
    ' Bounce enemy back when hit
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    Dim shp As Shape
    Set shp = enemy.GetShape()
    If shp Is Nothing Then Exit Sub
    
    Dim bounceDir As String
    bounceDir = m_BouncebackDir(index)
    
    Select Case bounceDir
        Case "U", "LU", "RU"
            shp.Top = shp.Top - 15
        Case "D", "LD", "RD"
            shp.Top = shp.Top + 15
        Case "L"
            shp.Left = shp.Left - 15
        Case "R"
            shp.Left = shp.Left + 15
    End Select
End Sub

Private Sub PushEnemy(ByVal index As Long)
    ' Push enemy when Link has shield
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    Dim shp As Shape
    Set shp = enemy.GetShape()
    If shp Is Nothing Then Exit Sub
    
    Dim pushDir As String
    pushDir = Sheets(SHEET_DATA).Range(RANGE_MOVE_DIR).Value
    
    Select Case pushDir
        Case "U", "LU", "RU"
            shp.Top = shp.Top - 5
        Case "D", "LD", "RD"
            shp.Top = shp.Top + 5
        Case "L"
            shp.Left = shp.Left - 5
        Case "R"
            shp.Left = shp.Left + 5
    End Select
End Sub

Private Sub KillEnemy(ByVal index As Long)
    ' Kill enemy and show explosion
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(index)
    
    Dim shp As Shape
    Set shp = enemy.GetShape()
    If shp Is Nothing Then Exit Sub
    
    ' Show explosion animation
    With ActiveSheet
        .Shapes("Explosion1").Top = shp.Top
        .Shapes("Explosion1").Left = shp.Left
        .Shapes("Explosion2").Top = shp.Top
        .Shapes("Explosion2").Left = shp.Left
        .Shapes("Explosion3").Top = shp.Top - 5
        .Shapes("Explosion3").Left = shp.Left - 5
        
        .Shapes("Explosion1").Visible = True
        Range("A1").Copy Range("A2")
        Sleep 8
        
        .Shapes("Explosion1").Visible = False
        .Shapes("Explosion2").Visible = True
        Range("A1").Copy Range("A2")
        Sleep 8
        
        .Shapes("Explosion2").Visible = False
        .Shapes("Explosion3").Visible = True
        Range("A1").Copy Range("A2")
        Sleep 8
        
        .Shapes("Explosion3").Visible = False
    End With
    
    ' Hide enemy
    enemy.Hide
End Sub

Private Sub ShootProjectile(ByVal index As Long)
    ' Enemy shoots projectile
    On Error Resume Next
    ' TODO: Implement projectile system
End Sub

'###################################################################################
'                              ENEMY SPAWNING
'###################################################################################

Public Function SpawnEnemy(ByVal enemyTypeName As String, ByVal slotNum As Long, _
                          ByVal dataRow As Long, ByVal cellAddress As String) As Boolean
    ' Spawn enemy at location
    On Error Resume Next
    
    If slotNum < 1 Or slotNum > 4 Then Exit Function
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(slotNum)
    
    ' Load enemy data
    enemy.LoadFromData enemyTypeName, slotNum, dataRow
    
    ' Show at location
    enemy.Show cellAddress
    
    ' Mark as active in Data sheet
    Sheets(SHEET_DATA).Range("C" & dataRow).Value = "Y"
    
    SpawnEnemy = True
End Function

Public Sub DespawnEnemy(ByVal slotNum As Long, ByVal dataRow As Long)
    ' Remove enemy from game
    On Error Resume Next
    
    If slotNum < 1 Or slotNum > 4 Then Exit Sub
    
    m_Enemies(slotNum).Hide
    m_BouncebackDir(slotNum) = ""
    
    ' Mark as inactive in Data sheet
    If dataRow > 0 Then
        Sheets(SHEET_DATA).Range("C" & dataRow).Value = "N"
    End If
End Sub

'###################################################################################
'                              ACCESSORS
'###################################################################################

Public Property Get Enemy(ByVal index As Long) As Enemy
    If index >= 1 And index <= 4 Then
        Set Enemy = m_Enemies(index)
    End If
End Property

Public Function IsActive(ByVal index As Long) As Boolean
    If index >= 1 And index <= 4 Then
        IsActive = m_Enemies(index).IsActive
    End If
End Function