Option Explicit

'###################################################################################
'                              ENEMY MANAGER
'###################################################################################

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const MAX_ENEMY_SLOTS As Long = 4

Private m_Enemies(1 To MAX_ENEMY_SLOTS) As Enemy
Private m_BouncebackDir(1 To MAX_ENEMY_SLOTS) As String

Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Reset()
    EnsureSlots
    ResetSlots
End Sub

Public Sub Initialize()
    EnsureSlots
    ResetSlots
End Sub

Public Sub Destroy()
    Dim slotIndex As Long
    For slotIndex = 1 To MAX_ENEMY_SLOTS
        If Not m_Enemies(slotIndex) Is Nothing Then
            m_Enemies(slotIndex).Hide
            Set m_Enemies(slotIndex) = Nothing
        End If
    Next slotIndex
End Sub

Private Sub EnsureSlots()
    Dim slotIndex As Long
    For slotIndex = LBound(m_Enemies) To UBound(m_Enemies)
        If m_Enemies(slotIndex) Is Nothing Then
            Set m_Enemies(slotIndex) = New Enemy
        End If
    Next slotIndex
End Sub

Private Sub ResetSlots()
    Dim slotIndex As Long
    For slotIndex = LBound(m_Enemies) To UBound(m_Enemies)
        If Not m_Enemies(slotIndex) Is Nothing Then
            m_Enemies(slotIndex).Hide
        End If
        m_BouncebackDir(slotIndex) = ""
    Next slotIndex
End Sub

Public Sub ProcessEnemy(ByVal slotIndex As Long, ByVal linkSprite As Object)
    On Error Resume Next

    If slotIndex < 1 Or slotIndex > MAX_ENEMY_SLOTS Then Exit Sub
    If Not m_Enemies(slotIndex).IsActive Then Exit Sub

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    If enemy.Hit > 0 Then
        ApplyBounceBack slotIndex
        enemy.Hit = enemy.Hit - 1

        If enemy.Hit = 0 And enemy.Life <= 0 Then
            KillEnemy slotIndex
        End If
    Else
        UpdateEnemy slotIndex
        CheckCollision slotIndex, linkSprite
    End If
End Sub

Public Sub HandleSwordHit(ByVal swordShape As Shape)
    On Error Resume Next
    If swordShape Is Nothing Then Exit Sub

    Dim attackDirection As String
    Dim gs As GameState
    Set gs = GameStateInstance()
    attackDirection = Trim$(gs.MoveDir)
    If attackDirection = "" Then attackDirection = Trim$(gs.LastDir)
    If attackDirection = "" Then attackDirection = "D"

    Dim slotIndex As Long
    For slotIndex = LBound(m_Enemies) To UBound(m_Enemies)
        If m_Enemies(slotIndex).IsActive Then
            Dim enemyShape As Shape
            Set enemyShape = m_Enemies(slotIndex).GetShape()
            If Not enemyShape Is Nothing Then
                If ShapesOverlap(swordShape, enemyShape) Then
                    HitEnemy slotIndex, attackDirection
                End If
            End If
        End If
    Next slotIndex
    On Error GoTo 0
End Sub

Private Sub UpdateEnemy(ByVal slotIndex As Long)
    On Error Resume Next

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    Select Case enemy.Behavior
        Case "Random"
            UpdateRandomMovement slotIndex
        Case "Still"
            UpdateStillAnimation slotIndex
        Case "StillFollow"
            UpdateStillFollow slotIndex
        Case "Chase"
            ' UpdateChase slotIndex
    End Select
End Sub

Private Sub UpdateRandomMovement(ByVal slotIndex As Long)
    ' Random movement AI
    On Error Resume Next

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    Dim enemyAddress As String
    enemyAddress = enemy.GetAddress()

    If enemy.Count = 10 Then
        enemy.SwapFrame
        enemy.Count = enemy.Count - 1
    ElseIf enemy.Count > 0 Then
        enemy.Count = enemy.Count - 1
    ElseIf enemy.Count = 0 Then
        Dim roll As Long
        roll = Int((5 - 1 + 1) * Rnd + 1)

        Select Case roll
            Case 1: enemy.Direction = "N"
            Case 2: enemy.Direction = "S"
            Case 3: enemy.Direction = "E"
            Case 4: enemy.Direction = "W"
            Case 5
                If enemy.CanShoot <> "" Then
                    ShootProjectile slotIndex
                End If
        End Select

        If enemy.ChangeRotation = "Y" Then
            Dim rotation As Long
            Select Case enemy.Direction
                Case "S": rotation = 0
                Case "N": rotation = 180
                Case "W": rotation = 90
                Case "E": rotation = 270
            End Select

            Dim sprite As Shape
            Set sprite = enemy.GetShape()
            If Not sprite Is Nothing Then
                sprite.Rotation = rotation
                If enemy.Frame1 <> "" Then ActiveSheet.Shapes(enemy.Frame1).Rotation = rotation
                If enemy.Frame2 <> "" Then ActiveSheet.Shapes(enemy.Frame2).Rotation = rotation
            End If
        End If

        enemy.Count = enemy.InitialCount
    End If

    MoveEnemy slotIndex, enemy.Direction, enemy.Speed, enemyAddress
End Sub

Private Sub UpdateStillAnimation(ByVal slotIndex As Long)
    ' Still enemy with animation
    On Error Resume Next

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    If enemy.Count = 10 Then
        enemy.SwapFrame
        enemy.Count = enemy.Count - 1
    ElseIf enemy.Count > 0 Then
        enemy.Count = enemy.Count - 1
    ElseIf enemy.Count = 0 Then
        enemy.Count = enemy.InitialCount
    End If
End Sub

Private Sub UpdateStillFollow(ByVal slotIndex As Long)
    ' Still enemy that follows/faces player
    On Error Resume Next
    UpdateStillAnimation slotIndex
    ' TODO: Add facing logic
End Sub

Private Sub MoveEnemy(ByVal slotIndex As Long, ByVal moveDirection As String, ByVal moveSpeed As Long, ByVal cellAddress As String)
    ' Move enemy in specified direction with collision detection
    On Error Resume Next

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    Dim sprite As Shape
    Set sprite = enemy.GetShape()
    If sprite Is Nothing Then Exit Sub

    Dim newTop As Double
    Dim newLeft As Double
    newTop = sprite.Top
    newLeft = sprite.Left

    Select Case moveDirection
        Case "N"
            If Range(cellAddress).Offset(-1, 1).Value = "" Or _
               Range(cellAddress).Offset(-1, 1).Value = "_\|/_" Then
                newTop = newTop - moveSpeed
            End If
        Case "S"
            If Range(cellAddress).Offset(4, 1).Value = "" Or _
               Range(cellAddress).Offset(4, 1).Value = "_\|/_" Then
                newTop = newTop + moveSpeed
            End If
        Case "E"
            If Range(cellAddress).Offset(2, 4).Value = "" Or _
               Range(cellAddress).Offset(2, 4).Value = "_\|/_" Then
                newLeft = newLeft + moveSpeed
            End If
        Case "W"
            If Range(cellAddress).Offset(2, -1).Value = "" Or _
               Range(cellAddress).Offset(2, -1).Value = "_\|/_" Then
                newLeft = newLeft - moveSpeed
            End If
    End Select
    
    BatchRendererInstance.QueuePosition sprite.Parent, sprite.Name, newTop, newLeft
End Sub

'===================================================================================
'                              COLLISION & COMBAT
'===================================================================================

Private Sub CheckCollision(ByVal slotIndex As Long, ByVal linkSprite As Object)
    ' Check collision with Link
    On Error Resume Next

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    If enemy.CanCollide <> "Y" Then Exit Sub

    Dim sprite As Shape
    Set sprite = enemy.GetShape()
    If sprite Is Nothing Then Exit Sub

    Dim sideOverlap As Boolean
    Dim verticalOverlap As Boolean

    If linkSprite.Left < sprite.Left And sprite.Left <= linkSprite.Left + linkSprite.Width Then
        sideOverlap = True
    ElseIf sprite.Left < linkSprite.Left And linkSprite.Left <= sprite.Left + sprite.Width Then
        sideOverlap = True
    End If

    If linkSprite.Top < sprite.Top And sprite.Top <= linkSprite.Top + linkSprite.Height Then
        verticalOverlap = True
    ElseIf sprite.Top < linkSprite.Top And linkSprite.Top <= sprite.Top + sprite.Height Then
        verticalOverlap = True
    End If
    
    If sideOverlap And verticalOverlap Then
        If DataCacheInstance.GetValue(RANGE_SHIELD_STATE) = "Y" Then
            PushEnemy slotIndex
        Else
            ' TODO: Implement Link damage/bounceback handling in new engine layer
        End If
    End If
End Sub

Private Function ShapesOverlap(ByVal firstShape As Shape, ByVal secondShape As Shape) As Boolean
    If firstShape Is Nothing Then Exit Function
    If secondShape Is Nothing Then Exit Function

    Dim horizontalOverlap As Boolean
    Dim verticalOverlap As Boolean

    horizontalOverlap = (firstShape.Left < secondShape.Left + secondShape.Width) And _
                        (secondShape.Left < firstShape.Left + firstShape.Width)

    verticalOverlap = (firstShape.Top < secondShape.Top + secondShape.Height) And _
                      (secondShape.Top < firstShape.Top + firstShape.Height)

    ShapesOverlap = horizontalOverlap And verticalOverlap
End Function

Public Sub HitEnemy(ByVal slotIndex As Long, ByVal hitDirection As String)
    On Error Resume Next

    If slotIndex < 1 Or slotIndex > MAX_ENEMY_SLOTS Then Exit Sub

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    If Not enemy.IsActive Then Exit Sub

    enemy.Hit = 5
    enemy.Life = enemy.Life - 1
    m_BouncebackDir(slotIndex) = hitDirection
End Sub

Private Sub ApplyBounceBack(ByVal slotIndex As Long)
    On Error Resume Next

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    Dim sprite As Shape
    Set sprite = enemy.GetShape()
    If sprite Is Nothing Then Exit Sub

    Dim bounceDirection As String
    bounceDirection = m_BouncebackDir(slotIndex)
    
    Dim newTop As Double
    Dim newLeft As Double
    newTop = sprite.Top
    newLeft = sprite.Left

    Select Case bounceDirection
        Case "U", "LU", "RU"
            newTop = newTop - 15
        Case "D", "LD", "RD"
            newTop = newTop + 15
        Case "L"
            newLeft = newLeft - 15
        Case "R"
            newLeft = newLeft + 15
    End Select
    
    BatchRendererInstance.QueuePosition sprite.Parent, sprite.Name, newTop, newLeft
End Sub

Private Sub PushEnemy(ByVal slotIndex As Long)
    ' Push enemy when Link has shield
    On Error Resume Next

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)

    Dim sprite As Shape
    Set sprite = enemy.GetShape()
    If sprite Is Nothing Then Exit Sub

    Dim dataSheet As Worksheet
    Set dataSheet = GameRegistryInstance().GetGameDataSheet()
    
    Dim pushDirection As String
    pushDirection = DataCacheInstance.GetValue(RANGE_MOVE_DIR)
    
    Dim newTop As Double
    Dim newLeft As Double
    newTop = sprite.Top
    newLeft = sprite.Left

    Select Case pushDirection
        Case "U", "LU", "RU"
            newTop = newTop - 5
        Case "D", "LD", "RD"
            newTop = newTop + 5
        Case "L"
            newLeft = newLeft - 5
        Case "R"
            newLeft = newLeft + 5
    End Select
    
    BatchRendererInstance.QueuePosition sprite.Parent, sprite.Name, newTop, newLeft
End Sub

Private Sub KillEnemy(ByVal slotIndex As Long)
    ' Kill enemy and show explosion
    On Error Resume Next
    
    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)
    
    Dim shp As Shape
    Set shp = enemy.GetShape()
    If shp Is Nothing Then Exit Sub
    
    Dim explosionTop As Double
    Dim explosionLeft As Double
    explosionTop = shp.Top
    explosionLeft = shp.Left
    
    ' Show explosion animation (batch queue)
    With ActiveSheet
        BatchRendererInstance.QueuePosition ActiveSheet, "Explosion1", explosionTop, explosionLeft
        BatchRendererInstance.QueuePosition ActiveSheet, "Explosion2", explosionTop, explosionLeft
        BatchRendererInstance.QueuePosition ActiveSheet, "Explosion3", explosionTop - 5, explosionLeft - 5
        
        BatchRendererInstance.QueueVisibility ActiveSheet, "Explosion1", True
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

Private Sub ShootProjectile(ByVal slotIndex As Long)
    ' Enemy shoots projectile
    On Error Resume Next
    ' TODO: Implement projectile system
End Sub

'===================================================================================
'                              ENEMY SPAWNING
'===================================================================================

Public Function SpawnEnemy(ByVal enemyTypeName As String, ByVal slotIndex As Long, _
                          ByVal dataRow As Long, ByVal cellAddress As String) As Boolean
    ' Spawn enemy at location
    On Error Resume Next

    If slotIndex < 1 Or slotIndex > MAX_ENEMY_SLOTS Then Exit Function

    Dim enemy As Enemy
    Set enemy = m_Enemies(slotIndex)
    
    ' Load enemy data
    enemy.LoadFromData enemyTypeName, slotIndex, dataRow
    
    ' Show at location
    enemy.Show cellAddress
    
    ' Mark as active in Data sheet
    Dim dataSheet As Worksheet
    Set dataSheet = GameRegistryInstance().GetGameDataSheet()
    dataSheet.Range("C" & dataRow).Value = "Y"
    
    SpawnEnemy = True
End Function

Public Sub DespawnEnemy(ByVal slotIndex As Long, ByVal dataRow As Long)
    ' Remove enemy from game
    On Error Resume Next
    
    If slotIndex < 1 Or slotIndex > MAX_ENEMY_SLOTS Then Exit Sub

    m_Enemies(slotIndex).Hide
    m_BouncebackDir(slotIndex) = ""
    
    ' Mark as inactive in Data sheet
    If dataRow > 0 Then
        Dim dataSheet As Worksheet
        Set dataSheet = GameRegistryInstance().GetGameDataSheet()
        dataSheet.Range("C" & dataRow).Value = "N"
    End If
End Sub

'===================================================================================
'                              ACCESSORS
'===================================================================================

Public Property Get Enemy(ByVal slotIndex As Long) As Enemy
    If slotIndex >= 1 And slotIndex <= MAX_ENEMY_SLOTS Then
        Set Enemy = m_Enemies(slotIndex)
    End If
End Property

Public Function IsActive(ByVal slotIndex As Long) As Boolean
    If slotIndex >= 1 And slotIndex <= MAX_ENEMY_SLOTS Then
        IsActive = m_Enemies(slotIndex).IsActive
    End If
End Function