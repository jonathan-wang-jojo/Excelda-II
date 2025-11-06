'@Folder("Game.Configs")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' GAME CONFIGURATION TEMPLATE
'═══════════════════════════════════════════════════════════════════════════════
' Copy this template to create your own game configuration!
'
' Steps to create a new game:
'   1. Copy this file and rename (e.g., GameConfigMyGame.cls)
'   2. Change class name in Properties window
'   3. Fill in all the Property Get methods below with your game's settings
'   4. Optionally override lifecycle hooks (OnGameInitialize, OnGameStart, OnGameEnd)
'   5. Register in GameRegistry.RegisterDefaultGames() or manually:
'      GameRegistryInstance().RegisterGame New GameConfigMyGame
'   6. Done! Engine will automatically load your config when that sheet is active
'
' That's it! No need to modify engine code. Your game is a plugin.
'═══════════════════════════════════════════════════════════════════════════════

Implements IGameConfig

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Basic Identity
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_GameName() As String
    ' TODO: Set your game's display name
    IGameConfig_GameName = "My Awesome Game"
End Property

Private Property Get IGameConfig_GameVersion() As String
    ' TODO: Set your game's version
    IGameConfig_GameVersion = "1.0.0"
End Property

Private Property Get IGameConfig_SheetName() As String
    ' TODO: Set the Excel sheet name where your game runs
    IGameConfig_SheetName = "MyGameSheet"
End Property

Private Property Get IGameConfig_DataSheetName() As String
    ' TODO: Set the Excel sheet name for game data storage
    ' Can be empty string "" to use engine default data sheet
    ' Recommended: Create dedicated data sheet for your game (e.g., "MyGameData")
    IGameConfig_DataSheetName = "MyGameData"
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Player Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_PlayerSpriteName() As String
    ' TODO: Set your player sprite shape name (must exist on sheet)
    IGameConfig_PlayerSpriteName = "Player"
End Property

Private Property Get IGameConfig_PlayerSpriteBaseName() As String
    ' Optional: Override base name used for sprite discovery (e.g., "Link")
    ' Return "" to reuse PlayerSpriteName
    IGameConfig_PlayerSpriteBaseName = ""
End Property

Private Property Get IGameConfig_PlayerSpeed() As Double
    ' TODO: Set player movement speed (pixels per tick at 60Hz)
    ' Examples: 24 = fast (Zelda), 2.5 = slow (boss battle)
    IGameConfig_PlayerSpeed = 12#
End Property

Private Property Get IGameConfig_StartCell() As String
    ' TODO: Set starting cell address for player spawn
    IGameConfig_StartCell = "A1"
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Camera Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_CameraMode() As Long
    ' TODO: Choose camera mode:
    '   0 = CameraInstant: Always centered on player (Unity/modern style)
    '   1 = CameraMargins: Only scrolls when player near edges (platformer style)
    '   2 = CameraGrid: Snaps to screen quadrants (Zelda/retro style)
    IGameConfig_CameraMode = 0
End Property

Private Property Get IGameConfig_CameraBoundsMinRow() As Long
    ' TODO: Set minimum row for camera (0 = no restriction)
    IGameConfig_CameraBoundsMinRow = 0
End Property

Private Property Get IGameConfig_CameraBoundsMaxRow() As Long
    ' TODO: Set maximum row for camera (0 = no restriction)
    IGameConfig_CameraBoundsMaxRow = 0
End Property

Private Property Get IGameConfig_CameraBoundsMinColumn() As Long
    ' TODO: Set minimum column for camera (0 = no restriction)
    IGameConfig_CameraBoundsMinColumn = 0
End Property

Private Property Get IGameConfig_CameraBoundsMaxColumn() As Long
    ' TODO: Set maximum column for camera (0 = no restriction)
    IGameConfig_CameraBoundsMaxColumn = 0
End Property

Private Property Get IGameConfig_CameraRowMargin() As Long
    ' TODO: Row margin for CameraMargins mode (-1 = use engine default of 6)
    IGameConfig_CameraRowMargin = -1
End Property

Private Property Get IGameConfig_CameraColumnMargin() As Long
    ' TODO: Column margin for CameraMargins mode (-1 = use engine default of 10)
    IGameConfig_CameraColumnMargin = -1
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Visual Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_ZoomLevel() As Long
    ' TODO: Set zoom percentage (0 = default 100%, 200 = 2x zoom, etc.)
    IGameConfig_ZoomLevel = 0
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Lifecycle Hooks (Optional)
'──────────────────────────────────────────────────────────────────────────────
Private Sub IGameConfig_OnGameInitialize()
    ' TODO: (Optional) Add initialization code here
    ' Called once when game config is first loaded
    ' Example: Set up custom managers, load save data, etc.
End Sub

Private Sub IGameConfig_OnGameStart()
    ' TODO: (Optional) Add start code here
    ' Called when game actually starts (after reset)
    ' Example: Spawn enemies, play music, show intro text, etc.
End Sub

Private Sub IGameConfig_OnGameEnd()
    ' TODO: (Optional) Add cleanup code here
    ' Called when game ends or player switches games
    ' Example: Save progress, stop music, cleanup temporary data, etc.
End Sub

'──────────────────────────────────────────────────────────────────────────────
' Custom Game Logic (Optional)
'──────────────────────────────────────────────────────────────────────────────
' Add any game-specific methods/properties below this line.
' These won't be called by the engine but can be used by your game code.
'
' Examples:
'   Public Function GetBossHealth() As Long
'   Public Sub ShowInventory()
'   Public Property Get CurrentLevel() As Long
'──────────────────────────────────────────────────────────────────────────────
