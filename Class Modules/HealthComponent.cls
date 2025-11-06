'@Folder("Entity.Components")
Option Explicit

'===================================================================================
'                              HEALTH COMPONENT
'===================================================================================
' Handles damage, death, and invulnerability for entities
' Classic Zelda pattern: brief invulnerability after taking damage
'===================================================================================

Private m_MaxHealth As Long
Private m_CurrentHealth As Long
Private m_IsInvulnerable As Boolean
Private m_InvulnerabilityTimer As Double

Private Const INVULNERABILITY_DURATION As Double = 0.5 ' 500ms (classic Zelda timing)

'===================================================================================
'                              INITIALIZATION
'===================================================================================

Public Sub Init(maxHealth As Long)
    m_MaxHealth = maxHealth
    m_CurrentHealth = maxHealth
    m_IsInvulnerable = False
    m_InvulnerabilityTimer = 0
End Sub

'===================================================================================
'                              DAMAGE SYSTEM
'===================================================================================

Public Sub TakeDamage(amount As Long)
    ' Cannot take damage if invulnerable or already dead
    If m_IsInvulnerable Then Exit Sub
    If m_CurrentHealth <= 0 Then Exit Sub
    
    ' Apply damage
    m_CurrentHealth = m_CurrentHealth - amount
    If m_CurrentHealth < 0 Then m_CurrentHealth = 0
    
    ' Brief invulnerability after hit (prevents multi-hit spam)
    If m_CurrentHealth > 0 Then ' Don't apply iframes if dead
        m_IsInvulnerable = True
        m_InvulnerabilityTimer = INVULNERABILITY_DURATION
    End If
End Sub

Public Sub Heal(amount As Long)
    m_CurrentHealth = m_CurrentHealth + amount
    If m_CurrentHealth > m_MaxHealth Then m_CurrentHealth = m_MaxHealth
End Sub

' Instant death (for special cases like bottomless pits)
Public Sub Kill()
    m_CurrentHealth = 0
End Sub

'===================================================================================
'                              UPDATE
'===================================================================================

Public Sub Update(deltaTime As Double)
    ' Update invulnerability timer
    If m_IsInvulnerable Then
        m_InvulnerabilityTimer = m_InvulnerabilityTimer - deltaTime
        If m_InvulnerabilityTimer <= 0 Then
            m_IsInvulnerable = False
            m_InvulnerabilityTimer = 0
        End If
    End If
End Sub

'===================================================================================
'                              PROPERTIES
'===================================================================================

Public Property Get IsAlive() As Boolean
    IsAlive = (m_CurrentHealth > 0)
End Property

Public Property Get IsDead() As Boolean
    IsDead = (m_CurrentHealth <= 0)
End Property

Public Property Get CurrentHealth() As Long
    CurrentHealth = m_CurrentHealth
End Property

Public Property Get MaxHealth() As Long
    MaxHealth = m_MaxHealth
End Property

Public Property Get IsInvulnerable() As Boolean
    IsInvulnerable = m_IsInvulnerable
End Property

' Get health as percentage (for health bars)
Public Property Get HealthPercent() As Double
    If m_MaxHealth = 0 Then
        HealthPercent = 0
    Else
        HealthPercent = CDbl(m_CurrentHealth) / CDbl(m_MaxHealth)
    End If
End Property
