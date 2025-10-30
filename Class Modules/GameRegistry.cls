Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' GAME REGISTRY
'═══════════════════════════════════════════════════════════════════════════════
' Central registry for game configurations. Acts as a factory for game instances.
' Allows engine to discover and load games dynamically without hardcoded references.
'
' Usage:
'   Dim registry As GameRegistry
'   Set registry = New GameRegistry
'   registry.RegisterGame New GameConfigLink
'   registry.RegisterGame New GameConfigMinotaur
'   
'   Dim config As IGameConfig
'   Set config = registry.GetConfigBySheet("Sheet9")  ' Returns Minotaur config
'═══════════════════════════════════════════════════════════════════════════════

Private m_Configs As Collection  ' Collection of IGameConfig implementations

'===================================================================================
'                              LIFECYCLE
'===================================================================================
Private Sub Class_Initialize()
    Set m_Configs = New Collection
    RegisterDefaultGames
End Sub

Private Sub Class_Terminate()
    Set m_Configs = Nothing
End Sub

'===================================================================================
'                              PUBLIC API - Registration
'===================================================================================
''' <summary>
''' Registers a game configuration with the registry.
''' </summary>
Public Sub RegisterGame(ByVal config As IGameConfig)
    If config Is Nothing Then Exit Sub
    
    ' Use sheet name as key for fast lookup
    Dim sheetKey As String
    sheetKey = UCase$(Trim$(config.SheetName))
    
    ' Remove existing if already registered
    On Error Resume Next
    m_Configs.Remove sheetKey
    On Error GoTo 0
    
    ' Add to collection
    m_Configs.Add config, sheetKey
End Sub

''' <summary>
''' Unregisters a game configuration by sheet name.
''' </summary>
Public Sub UnregisterGame(ByVal sheetName As String)
    Dim sheetKey As String
    sheetKey = UCase$(Trim$(sheetName))
    
    On Error Resume Next
    m_Configs.Remove sheetKey
    On Error GoTo 0
End Sub

''' <summary>
''' Clears all registered games.
''' </summary>
Public Sub ClearAll()
    Set m_Configs = New Collection
End Sub

'===================================================================================
'                              PUBLIC API - Retrieval
'===================================================================================
''' <summary>
''' Gets game configuration by sheet name.
''' Returns Nothing if no config found for that sheet.
''' </summary>
Public Function GetConfigBySheet(ByVal sheetName As String) As IGameConfig
    Dim sheetKey As String
    sheetKey = UCase$(Trim$(sheetName))
    
    On Error Resume Next
    Set GetConfigBySheet = m_Configs(sheetKey)
    On Error GoTo 0
End Function

''' <summary>
''' Gets game configuration by game name.
''' Returns first match (case-insensitive). Returns Nothing if not found.
''' </summary>
Public Function GetConfigByName(ByVal gameName As String) As IGameConfig
    Dim targetName As String
    targetName = UCase$(Trim$(gameName))
    
    Dim config As IGameConfig
    Dim i As Long
    For i = 1 To m_Configs.Count
        Set config = m_Configs(i)
        If UCase$(Trim$(config.GameName)) = targetName Then
            Set GetConfigByName = config
            Exit Function
        End If
    Next i
End Function

''' <summary>
''' Returns count of registered games.
''' </summary>
Public Function Count() As Long
    Count = m_Configs.Count
End Function

''' <summary>
''' Returns array of all registered sheet names.
''' </summary>
Public Function GetRegisteredSheets() As Variant
    If m_Configs.Count = 0 Then
        GetRegisteredSheets = Array()
        Exit Function
    End If
    
    ReDim sheets(1 To m_Configs.Count) As String
    Dim config As IGameConfig
    Dim i As Long
    For i = 1 To m_Configs.Count
        Set config = m_Configs(i)
        sheets(i) = config.SheetName
    Next i
    
    GetRegisteredSheets = sheets
End Function

''' <summary>
''' Gets all registered game configurations.
''' Returns array of IGameConfig objects.
''' </summary>
Public Function GetAllConfigs() As Collection
    Set GetAllConfigs = m_Configs
End Function

'===================================================================================
'                              PUBLIC API - Application
'===================================================================================
''' <summary>
''' Applies game configuration to engine managers.
''' Configures viewport, player speed, zoom, etc. based on config.
''' </summary>
Public Sub ApplyConfig(ByVal config As IGameConfig)
    If config Is Nothing Then Exit Sub
    
    ' Call initialization hook first
    On Error Resume Next
    config.OnGameInitialize
    On Error GoTo 0
    
    ' Configure Viewport
    Dim viewport As ViewportManager
    Set viewport = ViewportManagerInstance()
    If Not viewport Is Nothing Then
        viewport.SetCameraMode config.CameraMode
        
        If config.CameraBoundsMinRow > 0 Or config.CameraBoundsMaxRow > 0 Or _
           config.CameraBoundsMinColumn > 0 Or config.CameraBoundsMaxColumn > 0 Then
            viewport.SetCameraBounds config.CameraBoundsMinRow, config.CameraBoundsMaxRow, _
                                     config.CameraBoundsMinColumn, config.CameraBoundsMaxColumn
        Else
            viewport.ClearCameraBounds
        End If
    End If
    
    ' Set Zoom Level
    If config.ZoomLevel > 0 Then
        Dim win As Window
        On Error Resume Next
        Set win = ActiveWindow
        On Error GoTo 0
        If Not win Is Nothing Then
            If win.Zoom <> config.ZoomLevel Then
                On Error Resume Next
                win.Zoom = config.ZoomLevel
                On Error GoTo 0
            End If
        End If
    End If
    
    ' Call start hook
    On Error Resume Next
    config.OnGameStart
    On Error GoTo 0
End Sub

''' <summary>
''' Gets the data sheet for the currently active game.
''' Returns Nothing if no game is configured for the active sheet.
''' </summary>
Public Function GetGameDataSheet() As Worksheet
    Dim config As IGameConfig
    Set config = GetConfigBySheet(ActiveSheet.Name)
    
    If Not config Is Nothing And config.DataSheetName <> "" Then
        On Error Resume Next
        Set GetGameDataSheet = Sheets(config.DataSheetName)
        On Error GoTo 0
    End If
End Function

'===================================================================================
'                              PRIVATE HELPERS
'===================================================================================
Private Sub RegisterDefaultGames()
    ' Register built-in games
    ' New games can be registered externally via RegisterGame()
    
    RegisterGame New GameConfigLink
    RegisterGame New GameConfigMinotaur
End Sub
