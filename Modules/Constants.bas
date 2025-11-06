Attribute VB_Name = "Constants"
'@Folder("Shared")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' CORE ENGINE CONSTANTS
'═══════════════════════════════════════════════════════════════════════════════

'──────────────────────────────────────────────────────────────────────────────
' Engine Settings
'──────────────────────────────────────────────────────────────────────────────
Public Const GAME_VERSION As String = "0.2.0-minotaur-demo"

'──────────────────────────────────────────────────────────────────────────────
' Timing & Frame Control
'──────────────────────────────────────────────────────────────────────────────
Public Const TICK_RATE As Long = 60
Public Const FIXED_FRAME_SECONDS As Double = 1# / TICK_RATE
Public Const MAX_FRAME_SKIP As Long = 3            ' Spiral of death prevention
Public Const DEFAULT_GAME_SPEED As Long = 16
Public Const DEFAULT_FRAME_SECONDS As Double = DEFAULT_GAME_SPEED / 1000#
Public Const MIN_GAME_SPEED As Long = 8
Public Const MAX_GAME_SPEED As Long = 200

'──────────────────────────────────────────────────────────────────────────────
' Movement & Animation
'──────────────────────────────────────────────────────────────────────────────
Public Const DEFAULT_PLAYER_SPEED As Double = 24#
Public Const MIN_PLAYER_SPEED As Double = 2#
Public Const MAX_PLAYER_SPEED As Double = 60#
Public Const SPEED_MULTIPLIER As Double = 0.75
Public Const MIN_PIXELS_PER_TICK As Double = 1#
Public Const ANIMATION_TICKS_PER_FRAME As Long = 5

'──────────────────────────────────────────────────────────────────────────────
' Input Handling
'──────────────────────────────────────────────────────────────────────────────
Public Const INPUT_BUFFER_SECONDS As Double = 0.03

'──────────────────────────────────────────────────────────────────────────────
' Core Action Keys (C & D buttons - extensible for additional keys)
'──────────────────────────────────────────────────────────────────────────────
Public Const KEY_ACTION_PRIMARY As Integer = 67    ' C key
Public Const KEY_ACTION_SECONDARY As Integer = 68  ' D key
Public Const KEY_QUIT As Integer = 81              ' Q key

'──────────────────────────────────────────────────────────────────────────────
' Directional Keys
'──────────────────────────────────────────────────────────────────────────────
Public Const KEY_UP As Integer = 38
Public Const KEY_DOWN As Integer = 40
Public Const KEY_LEFT As Integer = 37
Public Const KEY_RIGHT As Integer = 39

'──────────────────────────────────────────────────────────────────────────────
' Legacy Key Aliases (for backwards compatibility)
'──────────────────────────────────────────────────────────────────────────────
Public Const KEY_C As Integer = KEY_ACTION_PRIMARY
Public Const KEY_D As Integer = KEY_ACTION_SECONDARY
Public Const KEY_Q As Integer = KEY_QUIT

'──────────────────────────────────────────────────────────────────────────────
' Enumerations
'──────────────────────────────────────────────────────────────────────────────
Public Enum Direction
    Up = 0
    Right = 1
    Down = 2
    Left = 3
End Enum

Public Enum EntityType
    Player = 0
    Enemy = 1
    NPC = 2
    Object = 3
End Enum

Public Enum GameStateType
    MainMenu = 0
    Playing = 1
    Paused = 2
    Dialog = 3
    GameOver = 4
End Enum

'═══════════════════════════════════════════════════════════════════════════════
' GAME-SPECIFIC CONFIGURATION
'═══════════════════════════════════════════════════════════════════════════════

'──────────────────────────────────────────────────────────────────────────────
' Sheet Names
'──────────────────────────────────────────────────────────────────────────────
Public Const SHEET_GAME As String = "Game1"
Public Const SHEET_DATA As String = "Data"
Public Const SHEET_TITLE As String = "Title"  ' Link game title screen

'──────────────────────────────────────────────────────────────────────────────
' Data Sheet Cell Ranges (Game-Specific State Storage)
'──────────────────────────────────────────────────────────────────────────────
Public Const RANGE_MOVE_DIR As String = "C21"
Public Const RANGE_GAME_SPEED As String = "C4"
Public Const RANGE_PLAYER_MOVE As String = "C19"
Public Const RANGE_ANIM_COUNTER As String = "C20"
Public Const RANGE_ACTION_C As String = "C24"
Public Const RANGE_ACTION_D As String = "C25"
Public Const RANGE_C_ITEM As String = "C26"
Public Const RANGE_D_ITEM As String = "C27"
Public Const RANGE_SHIELD_STATE As String = "C28"
Public Const RANGE_FALLING As String = "C9"
Public Const RANGE_FALL_SEQUENCE As String = "C10"
Public Const RANGE_SCROLL_COOLDOWN As String = "C6"
Public Const RANGE_CURRENT_CELL As String = "C8"
Public Const RANGE_SCROLL_DIRECTION As String = "C7"
Public Const RANGE_PREVIOUS_SCROLL As String = "D7"
Public Const RANGE_PREVIOUS_CELL As String = "D8"
Public Const RANGE_SCREEN_ROW As String = "C7"
Public Const RANGE_SCREEN_COLUMN As String = "C8"

'──────────────────────────────────────────────────────────────────────────────
' Scroll System Configuration
'──────────────────────────────────────────────────────────────────────────────
Public Const SCROLL_CODE_VERTICAL As String = "1"
Public Const SCROLL_CODE_HORIZONTAL As String = "2"
Public Const SCROLL_CODE_DIRECT_DOWN As String = "3"
Public Const SCROLL_CODE_DIRECT_UP As String = "4"
Public Const SCROLL_AMOUNT_VERTICAL As Long = 32
Public Const SCROLL_AMOUNT_HORIZONTAL As Long = 60
