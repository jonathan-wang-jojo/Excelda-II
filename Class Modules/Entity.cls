Option Explicit

'===================================================================================
'                              ENTITY CLASS
'===================================================================================
' Foundation class for all game entities (enemies, NPCs, objects, pickups)
' Uses composition: Transform + Sprite are REQUIRED, other components optional
'===================================================================================

' Required components (always present)
Private m_Transform As TransformComponent
Private m_Sprite As SpriteComponent

' Optional components (can be Nothing)
Private m_Health As HealthComponent
Private m_Collision As CollisionComponent
Private m_Behavior As IBehavior

' Entity metadata
Private m_EntityID As String
Private m_EntityType As String
Private m_IsActive As Boolean

'===================================================================================
'                              INITIALIZATION
'===================================================================================

' Factory method - creates entity with required components
Public Function Create(entityID As String, entityType As String, x As Double, y As Double) As Entity
    Set Create = Me
    
    m_EntityID = entityID
    m_EntityType = entityType
    m_IsActive = True
    
    ' Every entity MUST have Transform + Sprite
    Set m_Transform = New TransformComponent
    m_Transform.Init x, y
    
    Set m_Sprite = New SpriteComponent
    m_Sprite.Init m_Transform
End Function

'===================================================================================
'                              UPDATE LOOP
'===================================================================================

Public Sub Update(deltaTime As Double)
    If Not m_IsActive Then Exit Sub
    
    ' Update optional components if they exist
    If Not m_Behavior Is Nothing Then
        m_Behavior.Update Me, deltaTime
    End If
    
    If Not m_Health Is Nothing Then
        m_Health.Update deltaTime
    End If
    
    ' Sprite updates last (after position/animation changes)
    m_Sprite.Update
End Sub

'===================================================================================
'                              FLUENT API (Builder Pattern)
'===================================================================================
' Allows chaining: entity.WithHealth(3).WithCollision(8).WithBehavior(ai)
'===================================================================================

Public Function WithHealth(maxHealth As Long) As Entity
    Set m_Health = New HealthComponent
    m_Health.Init maxHealth
    Set WithHealth = Me
End Function

Public Function WithCollision(radius As Double) As Entity
    Set m_Collision = New CollisionComponent
    m_Collision.Init m_Transform, radius
    Set WithCollision = Me
End Function

Public Function WithBehavior(behavior As IBehavior) As Entity
    Set m_Behavior = behavior
    Set WithBehavior = Me
End Function

'===================================================================================
'                              COMPONENT ACCESS
'===================================================================================

Public Property Get Transform() As TransformComponent
    Set Transform = m_Transform
End Property

Public Property Get Sprite() As SpriteComponent
    Set Sprite = m_Sprite
End Property

Public Property Get Health() As HealthComponent
    Set Health = m_Health
End Property

Public Property Get Collision() As CollisionComponent
    Set Collision = m_Collision
End Property

Public Property Get Behavior() As IBehavior
    Set Behavior = m_Behavior
End Property

'===================================================================================
'                              ENTITY METADATA
'===================================================================================

Public Property Get EntityID() As String
    EntityID = m_EntityID
End Property

Public Property Get EntityType() As String
    EntityType = m_EntityType
End Property

Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property

'===================================================================================
'                              LIFECYCLE
'===================================================================================

Public Sub Destroy()
    m_IsActive = False
    m_Sprite.Hide
End Sub

' Check if entity has specific component
Public Function HasHealth() As Boolean
    HasHealth = Not (m_Health Is Nothing)
End Function

Public Function HasCollision() As Boolean
    HasCollision = Not (m_Collision Is Nothing)
End Function

Public Function HasBehavior() As Boolean
    HasBehavior = Not (m_Behavior Is Nothing)
End Function
