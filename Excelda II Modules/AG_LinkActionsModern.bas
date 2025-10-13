Option Explicit

'###################################################################################
'                              LINK ACTIONS (MODERN)
'###################################################################################
' Bridges legacy Link action routines (falling, jumping, sword, shield) to the
' modernized manager/state system without direct sheet manipulation.
'###################################################################################

Private Const LINK_FALL_FRAME_COUNT As Long = 3
Private Const LINK_FALL_DELAY_MS As Long = 300
Private Const LINK_JUMP_STEPS As Long = 30
Private Const LINK_JUMP_INCREMENT As Double = 2
Private Const LINK_JUMP_PHASES As Long = 3
Private Const LINK_SWORD_ANIM_DELAY As Long = 25

Private m_SpriteManager As SpriteManager
Private m_GameState As GameState

Public Sub InitializeLinkActions()
    Set m_SpriteManager = SpriteManagerInstance()
    Set m_GameState = GameStateInstance()
End Sub

Public Sub ResetLinkActions()
    Set m_SpriteManager = Nothing
    Set m_GameState = Nothing
End Sub
