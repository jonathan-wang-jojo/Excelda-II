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

'------------------------------- GameState -------------------------------
Private m_GameState As GameState
Public Function GameStateInstance() As GameState
    If m_GameState Is Nothing Then
        Set m_GameState = New GameState
    End If
    Set GameStateInstance = m_GameState
End Function

'------------------------------- ActionManager -------------------------------
Private m_ActionManager As ActionManager
Public Function ActionManagerInstance() As ActionManager
    If m_ActionManager Is Nothing Then
        Set m_ActionManager = New ActionManager
    End If
    Set ActionManagerInstance = m_ActionManager
End Function

'------------------------------- EnemyManager -------------------------------
Private m_EnemyManager As EnemyManager
Public Function EnemyManagerInstance() As EnemyManager
    If m_EnemyManager Is Nothing Then
        Set m_EnemyManager = New EnemyManager
    End If
    Set EnemyManagerInstance = m_EnemyManager
End Function

'------------------------------- SpriteManager -------------------------------
Private m_SpriteManager As SpriteManager
Public Function SpriteManagerInstance() As SpriteManager
    If m_SpriteManager Is Nothing Then
        Set m_SpriteManager = New SpriteManager
    End If
    Set SpriteManagerInstance = m_SpriteManager
End Function

'------------------------------- Reset All Managers -------------------------------
Public Sub ResetAllManagers()
    If Not m_GameState Is Nothing Then m_GameState.InitializeState
    If Not m_ActionManager Is Nothing Then m_ActionManager.Reset
    If Not m_EnemyManager Is Nothing Then m_EnemyManager.InitializeEnemies
    If Not m_SpriteManager Is Nothing Then m_SpriteManager.Reset
End Sub

'------------------------------- Destroy All Managers -------------------------------
Public Sub DestroyAllManagers()
    If Not m_GameState Is Nothing Then Set m_GameState = Nothing
    If Not m_ActionManager Is Nothing Then Set m_ActionManager = Nothing
    If Not m_EnemyManager Is Nothing Then Set m_EnemyManager = Nothing
    If Not m_SpriteManager Is Nothing Then Set m_SpriteManager = Nothing
End Sub
