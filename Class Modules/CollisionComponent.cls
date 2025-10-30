Option Explicit

'===================================================================================
'                              COLLISION COMPONENT
'===================================================================================
' Handles circle-based collision detection for entities
' Simple radius-based collisions (like original Zelda)
'===================================================================================

Private m_Transform As TransformComponent
Private m_Radius As Double
Private m_CollisionType As String ' "Enemy", "Friendly", "Projectile", "Pickup", "Solid"

'===================================================================================
'                              INITIALIZATION
'===================================================================================

Public Sub Init(transform As TransformComponent, radius As Double)
    Set m_Transform = transform
    m_Radius = radius
    m_CollisionType = "Enemy" ' Default collision type
End Sub

'===================================================================================
'                              COLLISION DETECTION
'===================================================================================

' Check collision with another entity
Public Function CheckCollisionWith(other As Entity) As Boolean
    CheckCollisionWith = False
    
    ' Safety checks
    If other Is Nothing Then Exit Function
    If Not other.HasCollision Then Exit Function
    
    ' Calculate distance between entities
    Dim dx As Double: dx = m_Transform.x - other.Transform.x
    Dim dy As Double: dy = m_Transform.y - other.Transform.y
    Dim distance As Double: distance = Sqr(dx * dx + dy * dy)
    
    ' Collision if distance less than combined radii
    CheckCollisionWith = (distance < m_Radius + other.Collision.Radius)
End Function

' Check collision with a specific point (for Link sword hits, etc.)
Public Function CheckCollisionWithPoint(x As Double, y As Double, pointRadius As Double) As Boolean
    Dim dx As Double: dx = m_Transform.x - x
    Dim dy As Double: dy = m_Transform.y - y
    Dim distance As Double: distance = Sqr(dx * dx + dy * dy)
    
    CheckCollisionWithPoint = (distance < m_Radius + pointRadius)
End Function

' Check collision with Link (convenience method)
Public Function CheckCollisionWithLink() As Boolean
    Dim linkX As Double: linkX = Singletons.DataCache.GetValue("LinkX")
    Dim linkY As Double: linkY = Singletons.DataCache.GetValue("LinkY")
    
    ' Link collision radius (standard 8 pixels)
    CheckCollisionWithLink = CheckCollisionWithPoint(linkX, linkY, 8)
End Function

'===================================================================================
'                              PROPERTIES
'===================================================================================

Public Property Get Radius() As Double
    Radius = m_Radius
End Property

Public Property Let Radius(value As Double)
    m_Radius = value
End Property

Public Property Get CollisionType() As String
    CollisionType = m_CollisionType
End Property

Public Property Let CollisionType(value As String)
    m_CollisionType = value
End Property
