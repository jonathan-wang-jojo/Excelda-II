'@Folder("Game.Configs")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' LINK'S ADVENTURE - GAME CONFIGURATION
'═══════════════════════════════════════════════════════════════════════════════
' Configuration for the original Zelda-style Link demo.
' Features quadrant-based camera system with screen grid navigation.
'═══════════════════════════════════════════════════════════════════════════════

Implements IGameConfig

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Basic Identity
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_GameName() As String
    IGameConfig_GameName = "Excelda XL"
End Property

Private Property Get IGameConfig_GameVersion() As String
    IGameConfig_GameVersion = "2.0.0"
End Property

Private Property Get IGameConfig_SheetName() As String
    IGameConfig_SheetName = "Game1"
End Property

Private Property Get IGameConfig_DataSheetName() As String
    IGameConfig_DataSheetName = "Data"  ' Shared legacy data sheet
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Player Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_PlayerSpriteName() As String
    IGameConfig_PlayerSpriteName = "Link"
End Property

Private Property Get IGameConfig_PlayerSpriteBaseName() As String
    IGameConfig_PlayerSpriteBaseName = "Link"
End Property

Private Property Get IGameConfig_PlayerSpeed() As Double
    IGameConfig_PlayerSpeed = 12#
End Property

Private Property Get IGameConfig_StartCell() As String
    IGameConfig_StartCell = "CO598"
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Camera Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_CameraMode() As Long
    ' Zelda-style grid quadrant camera (snaps to predefined screens)
    IGameConfig_CameraMode = 2  ' CameraGrid
End Property

Private Property Get IGameConfig_CameraBoundsMinRow() As Long
    IGameConfig_CameraBoundsMinRow = 0  ' No bounds restriction
End Property

Private Property Get IGameConfig_CameraBoundsMaxRow() As Long
    IGameConfig_CameraBoundsMaxRow = 0  ' No bounds restriction
End Property

Private Property Get IGameConfig_CameraBoundsMinColumn() As Long
    IGameConfig_CameraBoundsMinColumn = 0  ' No bounds restriction
End Property

Private Property Get IGameConfig_CameraBoundsMaxColumn() As Long
    IGameConfig_CameraBoundsMaxColumn = 0  ' No bounds restriction
End Property

Private Property Get IGameConfig_CameraRowMargin() As Long
    IGameConfig_CameraRowMargin = -1  ' Use engine default (6)
End Property

Private Property Get IGameConfig_CameraColumnMargin() As Long
    IGameConfig_CameraColumnMargin = -1  ' Use engine default (10)
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Visual Configuration
'──────────────────────────────────────────────────────────────────────────────
Private Property Get IGameConfig_ZoomLevel() As Long
    IGameConfig_ZoomLevel = 0  ' Use default zoom (typically 100%)
End Property

'──────────────────────────────────────────────────────────────────────────────
' IGameConfig Implementation - Lifecycle Hooks
'──────────────────────────────────────────────────────────────────────────────
Private Sub IGameConfig_OnGameInitialize()
    ' Link-specific initialization (if needed)
    ' e.g., Set up inventory system, hearts, rupees, etc.
End Sub

Private Sub IGameConfig_OnGameStart()
    ' Called when Link's Adventure starts
End Sub

Private Sub IGameConfig_OnGameEnd()
    ' Called when Link's Adventure ends
End Sub
