'@Folder("Entity.Components")
Option Explicit

'===================================================================================
'                              TRANSFORM COMPONENT
'===================================================================================
' Handles position, movement, and direction for entities
' Think: Where is it? Where is it moving? How fast?
'===================================================================================

Private m_X As Double
Private m_Y As Double
Private m_Direction As String ' "Up", "Down", "Left", "Right"
Private m_Speed As Double

'===================================================================================
'                              INITIALIZATION
'===================================================================================

Public Sub Init(x As Double, y As Double)
    m_X = x
    m_Y = y
    m_Direction = "Down" ' Default facing down (Zelda standard)
    m_Speed = 0
End Sub

'===================================================================================
'                              MOVEMENT
'===================================================================================

' Move by delta amounts (precise control)
Public Sub Move(dx As Double, dy As Double)
    m_X = m_X + dx
    m_Y = m_Y + dy
    
    ' Auto-update direction based on movement
    If Abs(dx) > Abs(dy) Then
        m_Direction = IIf(dx > 0, "Right", "Left")
    ElseIf Abs(dy) > 0 Then
        m_Direction = IIf(dy > 0, "Down", "Up")
    End If
End Sub

' Move in current direction at current speed (AI-friendly)
Public Sub MoveInDirection(deltaTime As Double)
    Dim moveAmount As Double
    moveAmount = m_Speed * deltaTime
    
    Select Case m_Direction
        Case "Up":    Move 0, -moveAmount
        Case "Down":  Move 0, moveAmount
        Case "Left":  Move -moveAmount, 0
        Case "Right": Move moveAmount, 0
    End Select
End Sub

' Set position directly (teleport)
Public Sub SetPosition(x As Double, y As Double)
    m_X = x
    m_Y = y
End Sub

'===================================================================================
'                              PROPERTIES
'===================================================================================

Public Property Get x() As Double
    x = m_X
End Property

Public Property Let x(value As Double)
    m_X = value
End Property

Public Property Get y() As Double
    y = m_Y
End Property

Public Property Let y(value As Double)
    m_Y = value
End Property

Public Property Get Direction() As String
    Direction = m_Direction
End Property

Public Property Let Direction(value As String)
    m_Direction = value
End Property

Public Property Get Speed() As Double
    Speed = m_Speed
End Property

Public Property Let Speed(value As Double)
    m_Speed = value
End Property

'===================================================================================
'                              UTILITY
'===================================================================================

' Calculate distance to another position
Public Function DistanceTo(targetX As Double, targetY As Double) As Double
    Dim dx As Double: dx = m_X - targetX
    Dim dy As Double: dy = m_Y - targetY
    DistanceTo = Sqr(dx * dx + dy * dy)
End Function

' Calculate distance to another transform
Public Function DistanceToTransform(other As TransformComponent) As Double
    DistanceToTransform = DistanceTo(other.x, other.y)
End Function

' Get direction toward a target position (for AI aiming)
Public Function DirectionToward(targetX As Double, targetY As Double) As String
    Dim dx As Double: dx = targetX - m_X
    Dim dy As Double: dy = targetY - m_Y
    
    If Abs(dx) > Abs(dy) Then
        DirectionToward = IIf(dx > 0, "Right", "Left")
    Else
        DirectionToward = IIf(dy > 0, "Down", "Up")
    End If
End Function
