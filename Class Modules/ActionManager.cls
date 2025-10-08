Option Explicit

'###################################################################################
'                              ACTION MANAGEMENT
'###################################################################################
' Handles all action items and action sprites (sword, shield, etc.)
'###################################################################################

' Action state
Private Type ActionState
    CItem As String
    DItem As String
    CPress As Long
    DPress As Long
End Type

Private m_Action As ActionState

' Action sprites
Private Type ActionSprites
    SwordFrame1 As Object
    SwordFrame2 As Object
    SwordFrame3 As Object
    ShieldSprite As Object
End Type

Private m_Sprites As ActionSprites

' Class Init
Private Sub Class_Initialize()
    Dim wsData As Worksheet
    Dim wsGame As Worksheet

    Set wsData = Sheets(SHEET_DATA)
    Set wsGame = Sheets(GameStateInstance.CurrentScreen)

    ' Initialize action items from data sheet
    m_Action.CItem = wsData.Range(RANGE_C_ITEM).Value
    m_Action.DItem = wsData.Range(RANGE_D_ITEM).Value
    m_Action.CPress = 0
    m_Action.DPress = 0
    
    ' Initialize action sprites
    Set m_Sprites.SwordFrame1 = wsGame.Shapes("SwordLeft")
    Set m_Sprites.SwordFrame2 = wsGame.Shapes("SwordSwipeDownLeft")
    Set m_Sprites.SwordFrame3 = wsGame.Shapes("SwordDown")
    Set m_Sprites.ShieldSprite = wsGame.Shapes("LinkShieldDown")
End Sub

' CALL WHEN game restarts or change levels
' uncalled currently
Public Sub Reset()
    m_Action.CPress = 0
    m_Action.DPress = 0
    m_Sprites.SwordFrame1.Visible = False
    m_Sprites.SwordFrame2.Visible = False
    m_Sprites.SwordFrame3.Visible = False
    m_Sprites.ShieldSprite.Visible = False
End Sub

Public Sub Destroy()
    Set m_Sprites.SwordFrame1 = Nothing
    Set m_Sprites.SwordFrame2 = Nothing
    Set m_Sprites.SwordFrame3 = Nothing
    Set m_Sprites.ShieldSprite = Nothing
End Sub

' Action item properties
Public Property Get CItem() As String
    CItem = m_Action.CItem
End Property

Public Property Let CItem(ByVal Value As String)
    m_Action.CItem = Value
End Property

Public Property Get DItem() As String
    DItem = m_Action.DItem
End Property

Public Property Let DItem(ByVal Value As String)
    m_Action.DItem = Value
End Property

Public Property Get CPress() As Long
    CPress = m_Action.CPress
End Property

Public Property Let CPress(ByVal Value As Long)
    m_Action.CPress = Value
End Property

Public Property Get DPress() As Long
    DPress = m_Action.DPress
End Property

Public Property Let DPress(ByVal Value As Long)
    m_Action.DPress = Value
End Property

' Action sprite properties
Public Property Get SwordFrame1() As Object
    Set SwordFrame1 = m_Sprites.SwordFrame1
End Property

Public Property Set SwordFrame1(ByVal Value As Object)
    Set m_Sprites.SwordFrame1 = Value
End Property

Public Property Get SwordFrame2() As Object
    Set SwordFrame2 = m_Sprites.SwordFrame2
End Property

Public Property Set SwordFrame2(ByVal Value As Object)
    Set m_Sprites.SwordFrame2 = Value
End Property

Public Property Get SwordFrame3() As Object
    Set SwordFrame3 = m_Sprites.SwordFrame3
End Property

Public Property Set SwordFrame3(ByVal Value As Object)
    Set m_Sprites.SwordFrame3 = Value
End Property

Public Property Get ShieldSprite() As Object
    Set ShieldSprite = m_Sprites.ShieldSprite
End Property

Public Property Set ShieldSprite(ByVal Value As Object)
    Set m_Sprites.ShieldSprite = Value
End Property

' Action methods
Public Sub HandleActionKey(ByVal keyCode As Integer, ByVal item As String, _
                           ByRef pressCounter As Long, ByVal flagCell As String)
    Dim keyPressed As Boolean
    keyPressed = (GetAsyncKeyState(keyCode) <> 0)
    UpdateActionState keyPressed, item, pressCounter, flagCell
End Sub

Private Sub UpdateActionState(ByVal keyPressed As Boolean, ByVal item As String, _
                              ByRef pressCounter As Long, ByVal flagCell As String)
    Dim ws As Worksheet
    Set ws = Sheets(SHEET_DATA)

    If keyPressed Then
        ws.Range(flagCell).Value = "Y"
        pressCounter = pressCounter + 1
        Select Case item
            Case "Sword": SwordSwipe IIf(keyCode = KEY_C, 1, 2)
            Case "Shield": ShowShield
        End Select
    Else
        ResetAction item, pressCounter
        ws.Range(flagCell).Value = ""
        pressCounter = 0
    End If
End Sub

Private Sub ResetAction(ByVal item As String, ByVal pressCounter As Long)
    Select Case item
        Case "Sword"
            If pressCounter >= 20 Then SwordSpin
            m_Sprites.SwordFrame1.Visible = False
            m_Sprites.SwordFrame2.Visible = False
            m_Sprites.SwordFrame3.Visible = False
        Case "Shield"
            m_Sprites.ShieldSprite.Visible = False
            Sheets(SHEET_DATA).Range(RANGE_SHIELD_STATE).Value = ""
    End Select
End Sub