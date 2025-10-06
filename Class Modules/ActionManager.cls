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

' Singleton pattern
Private m_Instance As ActionManager
Private m_Initialized As Boolean

Private Sub Class_Initialize()
    If Not m_Initialized Then
        InitializeActions
        m_Initialized = True
    End If
End Sub

Private Sub InitializeActions()
    ' Initialize action items from data sheet
    m_Action.CItem = Sheets(SHEET_DATA).Range(RANGE_C_ITEM).Value
    m_Action.DItem = Sheets(SHEET_DATA).Range(RANGE_D_ITEM).Value
    m_Action.CPress = 0
    m_Action.DPress = 0
    
    ' Initialize action sprites
    Set m_Sprites.SwordFrame1 = ActiveSheet.Shapes("SwordLeft")
    Set m_Sprites.SwordFrame2 = ActiveSheet.Shapes("SwordSwipeDownLeft")
    Set m_Sprites.SwordFrame3 = ActiveSheet.Shapes("SwordDown")
    Set m_Sprites.ShieldSprite = ActiveSheet.Shapes("LinkShieldDown")
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
Public Sub HandleActionKey(ByVal keyCode As Integer, ByVal item As String, ByRef pressCounter As Long, ByVal flagCell As String)
    If GetAsyncKeyState(keyCode) <> 0 Then
        Sheets(SHEET_DATA).Range(flagCell).Value = "Y"
        pressCounter = pressCounter + 1
        
        Select Case item
            Case "Sword"
                Call SwordSwipe(IIf(keyCode = KEY_C, 1, 2))
            Case "Shield"
                Call ShowShield
        End Select
    Else
        If item = "Sword" Then
            If pressCounter >= 20 Then Call SwordSpin
            m_Sprites.SwordFrame1.Visible = False
            m_Sprites.SwordFrame2.Visible = False
            m_Sprites.SwordFrame3.Visible = False
        ElseIf item = "Shield" Then
            m_Sprites.ShieldSprite.Visible = False
            Sheets(SHEET_DATA).Range(RANGE_SHIELD_STATE).Value = ""
        End If
        
        Sheets(SHEET_DATA).Range(flagCell).Value = ""
        pressCounter = 0
    End If
End Sub

' Singleton accessor
Public Property Get Instance() As ActionManager
    If m_Instance Is Nothing Then
        Set m_Instance = New ActionManager
    End If
    Set Instance = m_Instance
End Property
