'@Folder("Engine.Systems")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' DATA CACHE
'═══════════════════════════════════════════════════════════════════════════════
' Memory-resident cache for Data sheet state variables
' Eliminates 3-9ms per frame overhead by caching frequently accessed game state
' Drop-in replacement for dataSheet.Range(RANGE_X).Value with dirty tracking
'═══════════════════════════════════════════════════════════════════════════════
'   - Fallback mode if cache fails
'
' USAGE:
'   ' Instead of:
'   dataSheet.Range(RANGE_MOVE_DIR).Value = "U"
'   direction = dataSheet.Range(RANGE_MOVE_DIR).Value
'
'   ' Use:
'   DataCacheInstance().SetValue RANGE_MOVE_DIR, "U"
'   direction = DataCacheInstance().GetValue(RANGE_MOVE_DIR)
'
' PERFORMANCE:
'   Each Range().Value call costs 0.1-0.3ms
'   30 calls per frame = 3-9ms overhead
'   Cache lookup costs <0.001ms (direct memory access)
'   Expected savings: 3-9ms per frame (18-54% of 16.67ms frame budget)
'
'═══════════════════════════════════════════════════════════════════════════════

Private Type CacheData
    ' Cell values stored by address (to handle aliases correctly)
    ' C4 - RANGE_GAME_SPEED
    C4_GameSpeed As String
    
    ' C6 - RANGE_SCROLL_COOLDOWN
    C6_ScrollCooldown As String
    
    ' C7 - RANGE_SCROLL_DIRECTION / RANGE_SCREEN_ROW (ALIASED!)
    C7_Value As String
    
    ' C8 - RANGE_CURRENT_CELL / RANGE_SCREEN_COLUMN (ALIASED!)
    C8_Value As String
    
    ' C9 - RANGE_FALLING
    C9_Falling As String
    
    ' C10 - RANGE_FALL_SEQUENCE
    C10_FallSequence As String
    
    ' C19 - RANGE_PLAYER_MOVE
    C19_PlayerMove As String
    
    ' C20 - RANGE_ANIM_COUNTER (legacy, not used)
    C20_AnimCounter As String
    
    ' C21 - RANGE_MOVE_DIR
    C21_MoveDir As String
    
    ' C24 - RANGE_ACTION_C (legacy, not used)
    C24_ActionC As String
    
    ' C25 - RANGE_ACTION_D (legacy, not used)
    C25_ActionD As String
    
    ' C26 - RANGE_C_ITEM (legacy, not used)
    C26_CItem As String
    
    ' C27 - RANGE_D_ITEM (legacy, not used)
    C27_DItem As String
    
    ' C28 - RANGE_SHIELD_STATE
    C28_ShieldState As String
    
    ' D7 - RANGE_PREVIOUS_SCROLL
    D7_PreviousScroll As String
    
    ' D8 - RANGE_PREVIOUS_CELL
    D8_PreviousCell As String
    
    ' State tracking
    IsInitialized As Boolean
    IsDirty As Boolean
    FallbackMode As Boolean  ' If True, bypass cache and use direct sheet access
End Type

Private m_Data As CacheData

'═══════════════════════════════════════════════════════════════════════════════
' PUBLIC API
'═══════════════════════════════════════════════════════════════════════════════

Public Sub Initialize(ByVal dataSheet As Worksheet)
    '───────────────────────────────────────────────────────────────────────────
    ' Load all cached values from Data sheet into memory
    '───────────────────────────────────────────────────────────────────────────
    On Error GoTo ErrorHandler
    
    If dataSheet Is Nothing Then
        Debug.Print "DataCache.Initialize: ERROR - dataSheet is Nothing"
        m_Data.FallbackMode = True
        Exit Sub
    End If
    
    ' Load all cell values into cache
    m_Data.C4_GameSpeed = CStr(dataSheet.Range("C4").Value)
    m_Data.C6_ScrollCooldown = CStr(dataSheet.Range("C6").Value)
    m_Data.C7_Value = CStr(dataSheet.Range("C7").Value)
    m_Data.C8_Value = CStr(dataSheet.Range("C8").Value)
    m_Data.C9_Falling = CStr(dataSheet.Range("C9").Value)
    m_Data.C10_FallSequence = CStr(dataSheet.Range("C10").Value)
    m_Data.C19_PlayerMove = CStr(dataSheet.Range("C19").Value)
    m_Data.C20_AnimCounter = CStr(dataSheet.Range("C20").Value)
    m_Data.C21_MoveDir = CStr(dataSheet.Range("C21").Value)
    m_Data.C24_ActionC = CStr(dataSheet.Range("C24").Value)
    m_Data.C25_ActionD = CStr(dataSheet.Range("C25").Value)
    m_Data.C26_CItem = CStr(dataSheet.Range("C26").Value)
    m_Data.C27_DItem = CStr(dataSheet.Range("C27").Value)
    m_Data.C28_ShieldState = CStr(dataSheet.Range("C28").Value)
    m_Data.D7_PreviousScroll = CStr(dataSheet.Range("D7").Value)
    m_Data.D8_PreviousCell = CStr(dataSheet.Range("D8").Value)
    
    m_Data.IsInitialized = True
    m_Data.IsDirty = False
    m_Data.FallbackMode = False
    
    Debug.Print "DataCache.Initialize: SUCCESS - Loaded " & GetCellCount() & " cells into cache"
    Exit Sub
    
ErrorHandler:
    Debug.Print "DataCache.Initialize: ERROR - " & Err.Description
    m_Data.FallbackMode = True
    m_Data.IsInitialized = False
End Sub

Public Function GetValue(ByVal rangeName As String) As Variant
    '───────────────────────────────────────────────────────────────────────────
    ' Retrieve cached value for given RANGE_* constant
    ' Matches interface of: dataSheet.Range(RANGE_X).Value
    '───────────────────────────────────────────────────────────────────────────
    On Error GoTo ErrorHandler
    
    ' Fallback mode - direct sheet access
    If m_Data.FallbackMode Or Not m_Data.IsInitialized Then
        GetValue = GetFallbackValue(rangeName)
        Exit Function
    End If
    
    ' Map RANGE_* constant to cached cell value
    Select Case UCase$(Trim$(rangeName))
        Case "C4"
            GetValue = m_Data.C4_GameSpeed
        Case "C6"
            GetValue = m_Data.C6_ScrollCooldown
        Case "C7"
            GetValue = m_Data.C7_Value
        Case "C8"
            GetValue = m_Data.C8_Value
        Case "C9"
            GetValue = m_Data.C9_Falling
        Case "C10"
            GetValue = m_Data.C10_FallSequence
        Case "C19"
            GetValue = m_Data.C19_PlayerMove
        Case "C20"
            GetValue = m_Data.C20_AnimCounter
        Case "C21"
            GetValue = m_Data.C21_MoveDir
        Case "C24"
            GetValue = m_Data.C24_ActionC
        Case "C25"
            GetValue = m_Data.C25_ActionD
        Case "C26"
            GetValue = m_Data.C26_CItem
        Case "C27"
            GetValue = m_Data.C27_DItem
        Case "C28"
            GetValue = m_Data.C28_ShieldState
        Case "D7"
            GetValue = m_Data.D7_PreviousScroll
        Case "D8"
            GetValue = m_Data.D8_PreviousCell
        Case Else
            ' Unknown range name - fallback to sheet
            Debug.Print "DataCache.GetValue: WARNING - Unknown range '" & rangeName & "', using fallback"
            GetValue = GetFallbackValue(rangeName)
    End Select
    
    Exit Function
    
ErrorHandler:
    Debug.Print "DataCache.GetValue: ERROR - " & Err.Description & " for range '" & rangeName & "'"
    m_Data.FallbackMode = True
    GetValue = GetFallbackValue(rangeName)
End Function

Public Sub SetValue(ByVal rangeName As String, ByVal value As Variant)
    '───────────────────────────────────────────────────────────────────────────
    ' Update cached value for given RANGE_* constant
    ' Matches interface of: dataSheet.Range(RANGE_X).Value = value
    '───────────────────────────────────────────────────────────────────────────
    On Error GoTo ErrorHandler
    
    ' Fallback mode - direct sheet write
    If m_Data.FallbackMode Or Not m_Data.IsInitialized Then
        SetFallbackValue rangeName, value
        Exit Sub
    End If
    
    ' Convert value to string for consistent storage
    Dim stringValue As String
    stringValue = CStr(value)
    
    ' Map RANGE_* constant to cached cell value
    Select Case UCase$(Trim$(rangeName))
        Case "C4"
            If m_Data.C4_GameSpeed <> stringValue Then
                m_Data.C4_GameSpeed = stringValue
                m_Data.IsDirty = True
            End If
        Case "C6"
            If m_Data.C6_ScrollCooldown <> stringValue Then
                m_Data.C6_ScrollCooldown = stringValue
                m_Data.IsDirty = True
            End If
        Case "C7"
            If m_Data.C7_Value <> stringValue Then
                m_Data.C7_Value = stringValue
                m_Data.IsDirty = True
            End If
        Case "C8"
            If m_Data.C8_Value <> stringValue Then
                m_Data.C8_Value = stringValue
                m_Data.IsDirty = True
            End If
        Case "C9"
            If m_Data.C9_Falling <> stringValue Then
                m_Data.C9_Falling = stringValue
                m_Data.IsDirty = True
            End If
        Case "C10"
            If m_Data.C10_FallSequence <> stringValue Then
                m_Data.C10_FallSequence = stringValue
                m_Data.IsDirty = True
            End If
        Case "C19"
            If m_Data.C19_PlayerMove <> stringValue Then
                m_Data.C19_PlayerMove = stringValue
                m_Data.IsDirty = True
            End If
        Case "C20"
            If m_Data.C20_AnimCounter <> stringValue Then
                m_Data.C20_AnimCounter = stringValue
                m_Data.IsDirty = True
            End If
        Case "C21"
            If m_Data.C21_MoveDir <> stringValue Then
                m_Data.C21_MoveDir = stringValue
                m_Data.IsDirty = True
            End If
        Case "C24"
            If m_Data.C24_ActionC <> stringValue Then
                m_Data.C24_ActionC = stringValue
                m_Data.IsDirty = True
            End If
        Case "C25"
            If m_Data.C25_ActionD <> stringValue Then
                m_Data.C25_ActionD = stringValue
                m_Data.IsDirty = True
            End If
        Case "C26"
            If m_Data.C26_CItem <> stringValue Then
                m_Data.C26_CItem = stringValue
                m_Data.IsDirty = True
            End If
        Case "C27"
            If m_Data.C27_DItem <> stringValue Then
                m_Data.C27_DItem = stringValue
                m_Data.IsDirty = True
            End If
        Case "C28"
            If m_Data.C28_ShieldState <> stringValue Then
                m_Data.C28_ShieldState = stringValue
                m_Data.IsDirty = True
            End If
        Case "D7"
            If m_Data.D7_PreviousScroll <> stringValue Then
                m_Data.D7_PreviousScroll = stringValue
                m_Data.IsDirty = True
            End If
        Case "D8"
            If m_Data.D8_PreviousCell <> stringValue Then
                m_Data.D8_PreviousCell = stringValue
                m_Data.IsDirty = True
            End If
        Case Else
            ' Unknown range name - fallback to sheet
            Debug.Print "DataCache.SetValue: WARNING - Unknown range '" & rangeName & "', using fallback"
            SetFallbackValue rangeName, value
    End Select
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "DataCache.SetValue: ERROR - " & Err.Description & " for range '" & rangeName & "'"
    m_Data.FallbackMode = True
    SetFallbackValue rangeName, value
End Sub

Public Sub Flush(ByVal dataSheet As Worksheet)
    '───────────────────────────────────────────────────────────────────────────
    ' Write all dirty cached values back to Data sheet
    ' Call this when game state changes (screen transition, save, etc)
    '───────────────────────────────────────────────────────────────────────────
    On Error GoTo ErrorHandler
    
    If Not m_Data.IsDirty Then Exit Sub
    If dataSheet Is Nothing Then
        Debug.Print "DataCache.Flush: ERROR - dataSheet is Nothing"
        Exit Sub
    End If
    
    ' Write all cached values to sheet
    ' Using Application.ScreenUpdating = False for speed
    Dim wasScreenUpdating As Boolean
    wasScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    dataSheet.Range("C4").Value = m_Data.C4_GameSpeed
    dataSheet.Range("C6").Value = m_Data.C6_ScrollCooldown
    dataSheet.Range("C7").Value = m_Data.C7_Value
    dataSheet.Range("C8").Value = m_Data.C8_Value
    dataSheet.Range("C9").Value = m_Data.C9_Falling
    dataSheet.Range("C10").Value = m_Data.C10_FallSequence
    dataSheet.Range("C19").Value = m_Data.C19_PlayerMove
    dataSheet.Range("C20").Value = m_Data.C20_AnimCounter
    dataSheet.Range("C21").Value = m_Data.C21_MoveDir
    dataSheet.Range("C24").Value = m_Data.C24_ActionC
    dataSheet.Range("C25").Value = m_Data.C25_ActionD
    dataSheet.Range("C26").Value = m_Data.C26_CItem
    dataSheet.Range("C27").Value = m_Data.C27_DItem
    dataSheet.Range("C28").Value = m_Data.C28_ShieldState
    dataSheet.Range("D7").Value = m_Data.D7_PreviousScroll
    dataSheet.Range("D8").Value = m_Data.D8_PreviousCell
    
    Application.ScreenUpdating = wasScreenUpdating
    
    m_Data.IsDirty = False
    Debug.Print "DataCache.Flush: SUCCESS - Wrote " & GetCellCount() & " cells to sheet"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = wasScreenUpdating
    Debug.Print "DataCache.Flush: ERROR - " & Err.Description
End Sub

Public Sub Reset()
    '───────────────────────────────────────────────────────────────────────────
    ' Clear all cached values (called on game restart)
    '───────────────────────────────────────────────────────────────────────────
    m_Data.C4_GameSpeed = ""
    m_Data.C6_ScrollCooldown = ""
    m_Data.C7_Value = ""
    m_Data.C8_Value = ""
    m_Data.C9_Falling = ""
    m_Data.C10_FallSequence = ""
    m_Data.C19_PlayerMove = ""
    m_Data.C20_AnimCounter = ""
    m_Data.C21_MoveDir = ""
    m_Data.C24_ActionC = ""
    m_Data.C25_ActionD = ""
    m_Data.C26_CItem = ""
    m_Data.C27_DItem = ""
    m_Data.C28_ShieldState = ""
    m_Data.D7_PreviousScroll = ""
    m_Data.D8_PreviousCell = ""
    
    m_Data.IsInitialized = False
    m_Data.IsDirty = False
    m_Data.FallbackMode = False
    
    Debug.Print "DataCache.Reset: Cache cleared"
End Sub

Public Property Get IsDirty() As Boolean
    '───────────────────────────────────────────────────────────────────────────
    ' Check if cache has unflushed changes
    '───────────────────────────────────────────────────────────────────────────
    IsDirty = m_Data.IsDirty
End Property

Public Property Get IsInitialized() As Boolean
    '───────────────────────────────────────────────────────────────────────────
    ' Check if cache has been loaded from sheet
    '───────────────────────────────────────────────────────────────────────────
    IsInitialized = m_Data.IsInitialized
End Property

'═══════════════════════════════════════════════════════════════════════════════
' PRIVATE HELPERS
'═══════════════════════════════════════════════════════════════════════════════

Private Function GetFallbackValue(ByVal rangeName As String) As Variant
    '───────────────────────────────────────────────────────────────────────────
    ' Direct sheet access when cache is unavailable
    '───────────────────────────────────────────────────────────────────────────
    On Error Resume Next
    Dim dataSheet As Worksheet
    Set dataSheet = GameRegistryInstance().GetGameDataSheet()
    If Not dataSheet Is Nothing Then
        GetFallbackValue = dataSheet.Range(rangeName).Value
    Else
        GetFallbackValue = ""
    End If
End Function

Private Sub SetFallbackValue(ByVal rangeName As String, ByVal value As Variant)
    '───────────────────────────────────────────────────────────────────────────
    ' Direct sheet write when cache is unavailable
    '───────────────────────────────────────────────────────────────────────────
    On Error Resume Next
    Dim dataSheet As Worksheet
    Set dataSheet = GameRegistryInstance().GetGameDataSheet()
    If Not dataSheet Is Nothing Then
        dataSheet.Range(rangeName).Value = value
    End If
End Sub

Private Function GetCellCount() As Long
    '───────────────────────────────────────────────────────────────────────────
    ' Return number of cached cells
    '───────────────────────────────────────────────────────────────────────────
    GetCellCount = 16  ' C4,C6,C7,C8,C9,C10,C19,C20,C21,C24,C25,C26,C27,C28,D7,D8
End Function
