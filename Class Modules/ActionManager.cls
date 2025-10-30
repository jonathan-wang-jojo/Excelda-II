Option Explicit

Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long

'###################################################################################
'                       GENERIC ACTION KEY MANAGER (C & D)
'###################################################################################
' Generic manager for two action keys (C and D).
' Tracks key state (pressed/released, hold duration) without game-specific logic.
' Game-specific actions should be implemented in the active scene or game module.
'###################################################################################

Private Type ActionKeyState
    IsPressed As Boolean
    HoldFrames As Long
    WasJustPressed As Boolean
    WasJustReleased As Boolean
End Type

Private Type ActionData
    CKey As ActionKeyState
    DKey As ActionKeyState
    CItem As String
    DItem As String
End Type

Private m_Action As ActionData

Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Reset()
    ResetState
End Sub

Public Sub Destroy()
    ResetState
End Sub

Public Sub Initialize()
    ResetState
End Sub

Private Sub ResetState()
    ' Reset key states
    m_Action.CKey.IsPressed = False
    m_Action.CKey.HoldFrames = 0
    m_Action.CKey.WasJustPressed = False
    m_Action.CKey.WasJustReleased = False
    
    m_Action.DKey.IsPressed = False
    m_Action.DKey.HoldFrames = 0
    m_Action.DKey.WasJustPressed = False
    m_Action.DKey.WasJustReleased = False
    
    ' Reset item assignments
    m_Action.CItem = ""
    m_Action.DItem = ""
    
    ' Load items once during initialization
    SyncItemsFromCache
End Sub

'===================================================================================
'                              PUBLIC PROPERTIES
'===================================================================================

' C Key State
Public Property Get CKeyPressed() As Boolean
    CKeyPressed = m_Action.CKey.IsPressed
End Property

Public Property Get CKeyHoldFrames() As Long
    CKeyHoldFrames = m_Action.CKey.HoldFrames
End Property

Public Property Get CKeyJustPressed() As Boolean
    CKeyJustPressed = m_Action.CKey.WasJustPressed
End Property

Public Property Get CKeyJustReleased() As Boolean
    CKeyJustReleased = m_Action.CKey.WasJustReleased
End Property

' D Key State
Public Property Get DKeyPressed() As Boolean
    DKeyPressed = m_Action.DKey.IsPressed
End Property

Public Property Get DKeyHoldFrames() As Long
    DKeyHoldFrames = m_Action.DKey.HoldFrames
End Property

Public Property Get DKeyJustPressed() As Boolean
    DKeyJustPressed = m_Action.DKey.WasJustPressed
End Property

Public Property Get DKeyJustReleased() As Boolean
    DKeyJustReleased = m_Action.DKey.WasJustReleased
End Property

' Item Assignments (for game-specific logic)
Public Property Get CItem() As String
    CItem = m_Action.CItem
End Property

Public Property Get DItem() As String
    DItem = m_Action.DItem
End Property

'===================================================================================
'                              CORE UPDATE
'===================================================================================
' Updates key states by polling hardware. Call once per frame.
Public Sub UpdateKeys()
    UpdateKeyState KEY_C, m_Action.CKey
    UpdateKeyState KEY_D, m_Action.DKey
End Sub

Private Sub UpdateKeyState(ByVal keyCode As Integer, ByRef keyState As ActionKeyState)
    Dim currentlyPressed As Boolean
    currentlyPressed = (GetAsyncKeyState(keyCode) <> 0)
    
    ' Detect edge transitions
    keyState.WasJustPressed = currentlyPressed And Not keyState.IsPressed
    keyState.WasJustReleased = Not currentlyPressed And keyState.IsPressed
    
    ' Update pressed state
    keyState.IsPressed = currentlyPressed
    
    ' Update hold counter
    If keyState.IsPressed Then
        keyState.HoldFrames = keyState.HoldFrames + 1
    Else
        keyState.HoldFrames = 0
    End If
End Sub

'===================================================================================
'                              ITEM SYNCHRONIZATION
'===================================================================================
' Syncs item assignments from DataCache to local memory.
' Call this only when items change (e.g., after pickup), not every frame.
Public Sub SyncItemsFromCache()
    On Error Resume Next
    m_Action.CItem = DataCacheInstance.GetValue(RANGE_C_ITEM)
    m_Action.DItem = DataCacheInstance.GetValue(RANGE_D_ITEM)
    On Error GoTo 0
End Sub