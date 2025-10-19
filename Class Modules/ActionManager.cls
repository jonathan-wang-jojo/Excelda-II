Option Explicit

' Win32 API Declaration
Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long

'###################################################################################
'                              ACTION MANAGEMENT
'###################################################################################
' Handles all action items and action sprites (sword, shield, etc.)
'###################################################################################

Private Const SWORD_SPIN_THRESHOLD As Long = 20

Private Type ActionData
    CItem As String
    DItem As String
    CPress As Long
    DPress As Long
End Type

Private m_Action As ActionData

' Initialization
Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

' Core lifecycle
Public Sub Reset()
    ResetState
End Sub

Public Sub Destroy()
    ResetState
End Sub

Public Sub Initialize(Optional ByVal screenName As String = "")
    ResetState
End Sub

Private Sub ResetState()
    m_Action.CItem = ""
    m_Action.DItem = ""
    m_Action.CPress = 0
    m_Action.DPress = 0
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

' Core action processing - delegates to existing functions
Public Sub ProcessAction(ByVal keyCode As Integer)
    LoadItemsFromSheet
    Dim item As String
    Dim pressCounter As Long
    Dim actionSlot As Long
    
    Select Case keyCode
        Case KEY_C
            item = m_Action.CItem
            pressCounter = m_Action.CPress
            actionSlot = 1
        Case KEY_D
            item = m_Action.DItem
            pressCounter = m_Action.DPress
            actionSlot = 2
        Case Else
            Exit Sub
    End Select
    
    If GetAsyncKeyState(keyCode) <> 0 Then
        pressCounter = pressCounter + 1
        Select Case item
            Case "Sword"
                swordSwipe actionSlot, pressCounter
            Case "Shield"
                Call showShield
        End Select
    Else
        If item = "Sword" And pressCounter >= SWORD_SPIN_THRESHOLD Then swordSpin
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