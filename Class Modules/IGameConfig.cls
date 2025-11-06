'@Folder("Game.Configs")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' GAME CONFIGURATION INTERFACE
'═══════════════════════════════════════════════════════════════════════════════
' Interface defining the contract for game-specific configurations.
' Each game (Link, Minotaur, custom games) implements this interface to provide
' their settings to the engine. This allows complete separation of game logic
' from engine code.
'
' Usage Pattern:
'   1. Create a class implementing IGameConfig (e.g., GameConfigLink.cls)
'   2. Register it with GameRegistry
'   3. Engine loads config dynamically based on sheet/game name
'═══════════════════════════════════════════════════════════════════════════════

'──────────────────────────────────────────────────────────────────────────────
' Basic Game Identity
'──────────────────────────────────────────────────────────────────────────────
Public Property Get GameName() As String
    ' Human-readable name (e.g., "Link's Adventure", "Minotaur Battle")
End Property

Public Property Get GameVersion() As String
    ' Version string (e.g., "1.0.0", "0.2.0-demo")
End Property

Public Property Get SheetName() As String
    ' Excel worksheet name where this game runs (e.g., "Game1", "Sheet9")
End Property

Public Property Get DataSheetName() As String
    ' Excel worksheet for game data storage (e.g., "LinkData", "MinotaurData")
    ' Stores persistent data like player position, enemies, triggers, etc.
    ' Empty string = no dedicated data sheet (use engine default)
End Property

'──────────────────────────────────────────────────────────────────────────────
' Player Configuration
'──────────────────────────────────────────────────────────────────────────────
Public Property Get PlayerSpriteName() As String
    ' Name of player sprite shape (e.g., "Link", "Player", "Minotaur")
End Property

Public Property Get PlayerSpriteBaseName() As String
    ' Base name for sprite discovery (e.g., "Link", "Player").
    ' Return empty string to fall back to PlayerSpriteName.
End Property

Public Property Get PlayerSpeed() As Double
    ' Default player movement speed in pixels per tick
End Property

Public Property Get StartCell() As String
    ' Starting cell address for player spawn (e.g., "CO598", "DZ47")
End Property

'──────────────────────────────────────────────────────────────────────────────
' Camera Configuration
'──────────────────────────────────────────────────────────────────────────────
Public Property Get CameraMode() As Long
    ' Camera behavior: 0=Instant, 1=Margins, 2=Grid
End Property

Public Property Get CameraBoundsMinRow() As Long
    ' Minimum row for camera bounds (0 = no restriction)
End Property

Public Property Get CameraBoundsMaxRow() As Long
    ' Maximum row for camera bounds (0 = no restriction)
End Property

Public Property Get CameraBoundsMinColumn() As Long
    ' Minimum column for camera bounds (0 = no restriction)
End Property

Public Property Get CameraBoundsMaxColumn() As Long
    ' Maximum column for camera bounds (0 = no restriction)
End Property

Public Property Get CameraRowMargin() As Long
    ' Row margin for margin-based camera (-1 = use default)
End Property

Public Property Get CameraColumnMargin() As Long
    ' Column margin for margin-based camera (-1 = use default)
End Property

'──────────────────────────────────────────────────────────────────────────────
' Visual Configuration
'──────────────────────────────────────────────────────────────────────────────
Public Property Get ZoomLevel() As Long
    ' Worksheet zoom percentage (0 = use default, e.g., 100, 200)
End Property

'──────────────────────────────────────────────────────────────────────────────
' Lifecycle Hooks (Optional - override in implementation)
'──────────────────────────────────────────────────────────────────────────────
Public Sub OnGameInitialize()
    ' Called when game is first loaded
    ' Override to set up game-specific managers, data, etc.
End Sub

Public Sub OnGameStart()
    ' Called when game actually starts (after reset/load)
    ' Override for game-specific start logic
End Sub

Public Sub OnGameEnd()
    ' Called when game ends or switches
    ' Override for cleanup, save data, etc.
End Sub
