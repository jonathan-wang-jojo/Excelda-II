'@Folder("Entity.Behaviors")
Option Explicit
Implements IBehavior

'===================================================================================
'                              OCTOROK BEHAVIOR
'===================================================================================
' Classic Zelda Octorok AI:
' - Random direction changes every 2 seconds
' - Shoots projectile at Link every 3 seconds
' - Simple wandering movement
'===================================================================================

Private m_MoveTimer As Double
Private m_ShootTimer As Double

Private Const MOVE_CHANGE_INTERVAL As Double = 2#  ' Change direction every 2 seconds
Private Const SHOOT_INTERVAL As Double = 3#        ' Shoot every 3 seconds
Private Const MOVE_SPEED As Double = 2#            ' Pixels per frame

'===================================================================================
'                              IBEHAVIOR IMPLEMENTATION
'===================================================================================

Private Sub IBehavior_Update(entity As Entity, deltaTime As Double)
    ' Update timers
    m_MoveTimer = m_MoveTimer + deltaTime
    m_ShootTimer = m_ShootTimer + deltaTime
    
    ' Random direction changes (classic Octorok wandering)
    If m_MoveTimer >= MOVE_CHANGE_INTERVAL Then
        m_MoveTimer = 0
        ChangeDirection entity
    End If
    
    ' Move in current direction
    entity.Transform.MoveInDirection deltaTime
    
    ' TODO: Shoot at Link (Phase 2 - requires projectile entities)
    ' If m_ShootTimer >= SHOOT_INTERVAL Then
    '     m_ShootTimer = 0
    '     ShootAtLink entity
    ' End If
End Sub

'===================================================================================
'                              AI LOGIC
'===================================================================================

Private Sub ChangeDirection(entity As Entity)
    ' Random direction (Up/Down/Left/Right)
    Dim directions As Variant
    directions = Array("Up", "Down", "Left", "Right")
    
    Dim randomIndex As Long
    randomIndex = Int(Rnd * 4)
    
    entity.Transform.Direction = directions(randomIndex)
    entity.Transform.Speed = MOVE_SPEED
End Sub

Private Sub ShootAtLink(entity As Entity)
    ' TODO: Implement in Phase 2 (requires projectile entity system)
    ' For now, just a placeholder
    
    ' Dim projectile As Entity
    ' Set projectile = Singletons.EntityManager.CreateEntity("Projectile", _
    '     entity.Transform.x, entity.Transform.y)
    ' projectile.Sprite.LoadSprite "octorok_rock"
    ' projectile.WithBehavior(New ProjectileBehavior)
    ' ... aim at Link's position
End Sub
