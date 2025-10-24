Option Explicit

'===================================================================================
'                           SINGLETON ACCESSORS
'===================================================================================
' Provides central access to all game managers:
'   - GameState
'   - ActionManager
'   - EnemyManager
'   - SpriteManager
'===================================================================================

Private m_GameState As GameState
Private m_ActionManager As ActionManager
Private m_EnemyManager As EnemyManager
Private m_SpriteManager As SpriteManager
Private m_FriendlyManager As FriendlyManager
Private m_ObjectManager As ObjectManager
Private m_SpecialEventManager As SpecialEventManager
Private m_SceneManager As SceneManager
Private m_ViewportManager As ViewportManager

'------------------------------- GameState -------------------------------
Public Function GameStateInstance() As GameState
    If m_GameState Is Nothing Then
        Set m_GameState = New GameState
        m_GameState.Initialize
    End If
    Set GameStateInstance = m_GameState
End Function

'------------------------------- ActionManager -------------------------------
Public Function ActionManagerInstance() As ActionManager
    If m_ActionManager Is Nothing Then
        Set m_ActionManager = New ActionManager
        m_ActionManager.Initialize
    End If
    Set ActionManagerInstance = m_ActionManager
End Function

'------------------------------- EnemyManager -------------------------------
Public Function EnemyManagerInstance() As EnemyManager
    If m_EnemyManager Is Nothing Then
        Set m_EnemyManager = New EnemyManager
        m_EnemyManager.Initialize
    End If
    Set EnemyManagerInstance = m_EnemyManager
End Function

'------------------------------- SpriteManager -------------------------------
Public Function SpriteManagerInstance() As SpriteManager
    If m_SpriteManager Is Nothing Then
        Set m_SpriteManager = New SpriteManager
        m_SpriteManager.Initialize
    End If
    Set SpriteManagerInstance = m_SpriteManager
End Function

'------------------------------- FriendlyManager -------------------------------
Public Function FriendlyManagerInstance() As FriendlyManager
    If m_FriendlyManager Is Nothing Then
        Set m_FriendlyManager = New FriendlyManager
        m_FriendlyManager.Initialize
    End If
    Set FriendlyManagerInstance = m_FriendlyManager
End Function

'------------------------------- ObjectManager -------------------------------
Public Function ObjectManagerInstance() As ObjectManager
    If m_ObjectManager Is Nothing Then
        Set m_ObjectManager = New ObjectManager
        m_ObjectManager.Initialize
    End If
    Set ObjectManagerInstance = m_ObjectManager
End Function

'------------------------------- SpecialEventManager -------------------------------
Public Function SpecialEventManagerInstance() As SpecialEventManager
    If m_SpecialEventManager Is Nothing Then
        Set m_SpecialEventManager = New SpecialEventManager
        m_SpecialEventManager.Initialize
    End If
    Set SpecialEventManagerInstance = m_SpecialEventManager
End Function

'------------------------------- SceneManager -------------------------------
Public Function SceneManagerInstance() As SceneManager
    If m_SceneManager Is Nothing Then
        Set m_SceneManager = New SceneManager
        m_SceneManager.Initialize
    End If
    Set SceneManagerInstance = m_SceneManager
End Function

'------------------------------- ViewportManager -------------------------------
Public Function ViewportManagerInstance() As ViewportManager
    If m_ViewportManager Is Nothing Then
        Set m_ViewportManager = New ViewportManager
        m_ViewportManager.Initialize
    End If
    Set ViewportManagerInstance = m_ViewportManager
End Function

'------------------------------- Manager Lifecycle -------------------------------
Public Sub ResetAllManagers()
    If Not m_GameState Is Nothing Then m_GameState.Reset
    If Not m_ActionManager Is Nothing Then m_ActionManager.Reset
    If Not m_EnemyManager Is Nothing Then m_EnemyManager.Reset
    If Not m_SpriteManager Is Nothing Then m_SpriteManager.Reset
    If Not m_FriendlyManager Is Nothing Then m_FriendlyManager.Reset
    If Not m_ObjectManager Is Nothing Then
        m_ObjectManager.Reset
        m_ObjectManager.ResetObjects "All"
    End If
    If Not m_SpecialEventManager Is Nothing Then
        m_SpecialEventManager.Reset
    End If
    If Not m_SceneManager Is Nothing Then
        m_SceneManager.Reset
    End If
    If Not m_ViewportManager Is Nothing Then
        m_ViewportManager.Reset
    End If
End Sub

Public Sub DestroyAllManagers()
    If Not m_GameState Is Nothing Then m_GameState.Destroy: Set m_GameState = Nothing
    If Not m_ActionManager Is Nothing Then m_ActionManager.Destroy: Set m_ActionManager = Nothing
    If Not m_EnemyManager Is Nothing Then m_EnemyManager.Destroy: Set m_EnemyManager = Nothing
    If Not m_SpriteManager Is Nothing Then m_SpriteManager.Destroy: Set m_SpriteManager = Nothing
    If Not m_FriendlyManager Is Nothing Then m_FriendlyManager.Destroy: Set m_FriendlyManager = Nothing
    If Not m_ObjectManager Is Nothing Then m_ObjectManager.Destroy: Set m_ObjectManager = Nothing
    If Not m_SpecialEventManager Is Nothing Then m_SpecialEventManager.Destroy: Set m_SpecialEventManager = Nothing
    If Not m_SceneManager Is Nothing Then m_SceneManager.Destroy: Set m_SceneManager = Nothing
    If Not m_ViewportManager Is Nothing Then m_ViewportManager.Destroy: Set m_ViewportManager = Nothing
End Sub
'===================================================================================