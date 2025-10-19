Option Explicit

' Game constants
Public Const GAME_VERSION As String = "2.0.0"
Public Const TICK_RATE As Long = 60
Public Const DEFAULT_GAME_SPEED As Long = 16
Public Const MIN_GAME_SPEED As Long = 8
Public Const MAX_GAME_SPEED As Long = 200
Public Const DEFAULT_LINK_SPEED As Long = 8
Public Const MIN_LINK_SPEED As Long = 2
Public Const MAX_LINK_SPEED As Long = 20
Public Const DEFAULT_FRAME_SECONDS As Double = DEFAULT_GAME_SPEED / 1000#
Public Const MAX_FRAME_DELTA_SECONDS As Double = 0.25
Public Const MIN_FRAME_DELTA_SECONDS As Double = 0.001

' Direction enums
Public Enum Direction
    Up = 0
    Right = 1
    Down = 2
    Left = 3
End Enum

' Entity types
Public Enum EntityType
    Player = 0
    Enemy = 1
    NPC = 2
    Object = 3
End Enum

' Game states
Public Enum GameStateType
    MainMenu = 0
    Playing = 1
    Paused = 2
    Dialog = 3
    GameOver = 4
End Enum

' Sheet names
Public Const SHEET_GAME As String = "Game1"
Public Const SHEET_DATA As String = "Data"
Public Const SHEET_TITLE As String = "Title"

' Key Codes
Public Const KEY_LEFT As Integer = 37
Public Const KEY_RIGHT As Integer = 39
Public Const KEY_UP As Integer = 38
Public Const KEY_DOWN As Integer = 40
Public Const KEY_C As Integer = 67
Public Const KEY_D As Integer = 68
Public Const KEY_Q As Integer = 81

' Data Sheet Ranges
Public Const RANGE_MOVE_DIR As String = "C21"
Public Const RANGE_GAME_SPEED As String = "C4"
Public Const RANGE_LINK_MOVE As String = "C19"
Public Const RANGE_ANIM_COUNTER As String = "C20"
Public Const RANGE_C_ITEM As String = "C26"
Public Const RANGE_D_ITEM As String = "C27"
Public Const RANGE_FALLING As String = "C9"
Public Const RANGE_FALL_SEQUENCE As String = "C10"
Public Const RANGE_ACTION_C As String = "C24"
Public Const RANGE_ACTION_D As String = "C25"
Public Const RANGE_SHIELD_STATE As String = "C28"
Public Const RANGE_SCROLL_COOLDOWN As String = "C6"

' Scroll and Screen Management Ranges
Public Const RANGE_CURRENT_CELL As String = "C8"
Public Const RANGE_SCROLL_DIRECTION As String = "C7"
Public Const RANGE_PREVIOUS_SCROLL As String = "D7"
Public Const RANGE_PREVIOUS_CELL As String = "D8"
Public Const RANGE_SCREEN_ROW As String = "C7"
Public Const RANGE_SCREEN_COLUMN As String = "C8"

' Scroll Constants
Public Const SCROLL_VERTICAL As String = "1"
Public Const SCROLL_HORIZONTAL As String = "2"
Public Const SCROLL_AMOUNT_VERTICAL As Long = 32
Public Const SCROLL_AMOUNT_HORIZONTAL As Long = 60