Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' MINOTAUR BATTLE - GAME CONFIGURATION
'═══════════════════════════════════════════════════════════════════════════════
' Configuration for the Minotaur boss battle demo.
' Features tight arena with instant-follow camera and high zoom level.
'═══════════════════════════════════════════════════════════════════════════════

Implements IGameConfig

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Basic Identity
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_GameName() As String
    IGameConfig_GameName = "Minotaur Battle"
End Property

Private Property Get IGameConfig_GameVersion() As String
    IGameConfig_GameVersion = "0.2.0-demo"
End Property

Private Property Get IGameConfig_SheetName() As String
    IGameConfig_SheetName = "Maze"  ' Excel visible sheet name, not VBA object name (Sheet9)
End Property

Private Property Get IGameConfig_DataSheetName() As String
    IGameConfig_DataSheetName = "Data"  ' Shared legacy data sheet
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Player Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_PlayerSpriteName() As String
    IGameConfig_PlayerSpriteName = "Minotaur"
End Property

Private Property Get IGameConfig_PlayerSpriteBaseName() As String
    IGameConfig_PlayerSpriteBaseName = "Minotaur"
End Property

Private Property Get IGameConfig_PlayerSpeed() As Double
    IGameConfig_PlayerSpeed = 2.5#
End Property

Private Property Get IGameConfig_StartCell() As String
    IGameConfig_StartCell = "DZ47"
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Camera Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_CameraMode() As Long
    ' Instant follow camera for tight boss arena
    IGameConfig_CameraMode = 0  ' CameraInstant
End Property

Private Property Get IGameConfig_CameraBoundsMinRow() As Long
    IGameConfig_CameraBoundsMinRow = 17
End Property

Private Property Get IGameConfig_CameraBoundsMaxRow() As Long
    IGameConfig_CameraBoundsMaxRow = 82
End Property

Private Property Get IGameConfig_CameraBoundsMinColumn() As Long
    IGameConfig_CameraBoundsMinColumn = 101  ' Column CW
End Property

Private Property Get IGameConfig_CameraBoundsMaxColumn() As Long
    IGameConfig_CameraBoundsMaxColumn = 287  ' Column KA
End Property

Private Property Get IGameConfig_CameraRowMargin() As Long
    IGameConfig_CameraRowMargin = -1  ' Not used with instant camera
End Property

Private Property Get IGameConfig_CameraColumnMargin() As Long
    IGameConfig_CameraColumnMargin = -1  ' Not used with instant camera
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Visual Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_ZoomLevel() As Long
    IGameConfig_ZoomLevel = 200  ' 200% zoom for close-up boss battle
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Lifecycle Hooks
'──────────────────────────────────────────────────────────────────────────────
Private Sub IGameConfig_OnGameInitialize()
    ' Minotaur-specific initialization can be added here if needed
    ' NOTE: Called during ResetGame/ApplyConfig - managers may not be fully initialized yet
    ' Keep this minimal or move complex logic to OnGameStart
End Sub

Private Sub IGameConfig_OnGameStart()
    ' Called when Minotaur Battle starts (after all managers initialized)
End Sub

Private Sub IGameConfig_OnGameEnd()
    ' Called when Minotaur Battle ends
    ' e.g., Victory/defeat sequence, cleanup
End Sub
