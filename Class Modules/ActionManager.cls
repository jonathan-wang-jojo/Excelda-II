Option Explicit

' Win32 API Declaration
Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long

'###################################################################################
'                              ACTION MANAGEMENT
'###################################################################################
' Handles all action items and action sprites (sword, shield, etc.)
'###################################################################################

Private Type ActionData
    CItem As String
    DItem As String
    CPress As Long
    DPress As Long
    SwordFrame1 As Shape
    SwordFrame2 As Shape
    SwordFrame3 As Shape
    ShieldSprite As Shape
End Type

Private m_Action As ActionData

' Initialization
Private Sub Class_Initialize()
    On Error Resume Next
    m_Action.CItem = Sheets(SHEET_DATA).Range(RANGE_C_ITEM).Value
    m_Action.DItem = Sheets(SHEET_DATA).Range(RANGE_D_ITEM).Value
    m_Action.CPress = 0
    m_Action.DPress = 0
    Set m_Action.SwordFrame1 = ActiveSheet.Shapes("SwordLeft")
    Set m_Action.SwordFrame2 = ActiveSheet.Shapes("SwordSwipeDownLeft")
    Set m_Action.SwordFrame3 = ActiveSheet.Shapes("SwordDown")
    Set m_Action.ShieldSprite = ActiveSheet.Shapes("LinkShieldDown")
    On Error GoTo 0
End Sub

' Core lifecycle
Public Sub Reset()
    m_Action.CPress = 0
    m_Action.DPress = 0
End Sub

Public Sub Destroy()
    Reset
    Set m_Action.SwordFrame1 = Nothing
    Set m_Action.SwordFrame2 = Nothing
    Set m_Action.SwordFrame3 = Nothing
    Set m_Action.ShieldSprite = Nothing
End Sub

' Simple accessors
Public Property Get CItem() As String
    CItem = m_Action.CItem
End Property

Public Property Get DItem() As String
    DItem = m_Action.DItem
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

Public Property Get SwordFrame1() As Shape
    Set SwordFrame1 = m_Action.SwordFrame1
End Property

Public Property Get SwordFrame2() As Shape
    Set SwordFrame2 = m_Action.SwordFrame2
End Property

Public Property Get SwordFrame3() As Shape
    Set SwordFrame3 = m_Action.SwordFrame3
End Property

Public Property Get ShieldSprite() As Shape
    Set ShieldSprite = m_Action.ShieldSprite
End Property

' Core action processing - delegates to existing functions
Public Sub ProcessAction(ByVal keyCode As Integer, ByVal item As String, ByRef pressCounter As Long)
    If GetAsyncKeyState(keyCode) <> 0 Then
        pressCounter = pressCounter + 1
        ' Delegate to existing action functions
        Select Case item
            Case "Sword": Call swordSwipe(IIf(keyCode = KEY_C, 1, 2))
            Case "Shield": Call showShield
        End Select
    Else
        If item = "Sword" And pressCounter >= 20 Then Call swordSpin
        pressCounter = 0
    End If
End Sub