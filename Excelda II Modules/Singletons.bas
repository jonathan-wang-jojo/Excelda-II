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

'------------------------------- GameState -------------------------------
Public Function GameStateInstance() As GameState
    If m_GameState Is Nothing Then
        Set m_GameState = New GameState
    End If
    Set GameStateInstance = m_GameState
End Function

'------------------------------- ActionManager -------------------------------
Public Function ActionManagerInstance() As ActionManager
    If m_ActionManager Is Nothing Then
        Set m_ActionManager = New ActionManager
    End If
    Set ActionManagerInstance = m_ActionManager
End Function

'------------------------------- EnemyManager -------------------------------
Public Function EnemyManagerInstance() As EnemyManager
    If m_EnemyManager Is Nothing Then
        Set m_EnemyManager = New EnemyManager
    End If
    Set EnemyManagerInstance = m_EnemyManager
End Function

'------------------------------- SpriteManager -------------------------------
Public Function SpriteManagerInstance() As SpriteManager
    If m_SpriteManager Is Nothing Then
        Set m_SpriteManager = New SpriteManager
    End If
    Set SpriteManagerInstance = m_SpriteManager
End Function

'------------------------------- Manager Lifecycle -------------------------------
Public Sub ResetAllManagers()
    If Not m_GameState Is Nothing Then m_GameState.Reset
    If Not m_ActionManager Is Nothing Then m_ActionManager.Reset
    If Not m_EnemyManager Is Nothing Then m_EnemyManager.Reset
    If Not m_SpriteManager Is Nothing Then m_SpriteManager.Reset
End Sub

Public Sub DestroyAllManagers()
    If Not m_GameState Is Nothing Then m_GameState.Destroy: Set m_GameState = Nothing
    If Not m_ActionManager Is Nothing Then m_ActionManager.Destroy: Set m_ActionManager = Nothing
    If Not m_EnemyManager Is Nothing Then m_EnemyManager.Destroy: Set m_EnemyManager = Nothing
    If Not m_SpriteManager Is Nothing Then m_SpriteManager.Destroy: Set m_SpriteManager = Nothing
End Sub
'===================================================================================