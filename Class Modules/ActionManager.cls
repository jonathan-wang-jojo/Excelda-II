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
    ScreenName As String
End Type

Private m_Action As ActionData

' Initialization
Private Sub Class_Initialize()
    LoadItemsFromSheet
End Sub

' Core lifecycle
Public Sub Reset()
    m_Action.CPress = 0
    m_Action.DPress = 0
    LoadItemsFromSheet
End Sub

Public Sub Destroy()
    Reset
    Set m_Action.SwordFrame1 = Nothing
    Set m_Action.SwordFrame2 = Nothing
    Set m_Action.SwordFrame3 = Nothing
    Set m_Action.ShieldSprite = Nothing
End Sub

Public Sub Initialize(ByVal screenName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Sheets(screenName)
    m_Action.ScreenName = screenName
    Set m_Action.SwordFrame1 = ws.Shapes("SwordLeft")
    Set m_Action.SwordFrame2 = ws.Shapes("SwordSwipeDownLeft")
    Set m_Action.SwordFrame3 = ws.Shapes("SwordDown")
    Set m_Action.ShieldSprite = ws.Shapes("LinkShieldDown")
    On Error GoTo 0
    LoadItemsFromSheet
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
Public Sub ProcessAction(ByVal keyCode As Integer)
    LoadItemsFromSheet
    Dim item As String
    Dim pressCounter As Long
    
    Select Case keyCode
        Case KEY_C
            item = m_Action.CItem
            pressCounter = m_Action.CPress
        Case KEY_D
            item = m_Action.DItem
            pressCounter = m_Action.DPress
        Case Else
            Exit Sub
    End Select
    
    If GetAsyncKeyState(keyCode) <> 0 Then
        pressCounter = pressCounter + 1
        Select Case item
            Case "Sword"
                Call swordSwipe(IIf(keyCode = KEY_C, 1, 2))
            Case "Shield"
                Call showShield
        End Select
    Else
        If item = "Sword" And pressCounter >= 20 Then Call swordSpin
        pressCounter = 0
    End If
    
    Select Case keyCode
        Case KEY_C: m_Action.CPress = pressCounter
        Case KEY_D: m_Action.DPress = pressCounter
    End Select
End Sub

Private Sub LoadItemsFromSheet()
    On Error Resume Next
    m_Action.CItem = Sheets(SHEET_DATA).Range(RANGE_C_ITEM).Value
    m_Action.DItem = Sheets(SHEET_DATA).Range(RANGE_D_ITEM).Value
    On Error GoTo 0
End Sub