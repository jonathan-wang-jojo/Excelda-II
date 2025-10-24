'Attribute VB_Name = "SceneManager"
Option Explicit

Private m_Scenes As Collection
Private m_CurrentScene As IGameScene
Private m_CurrentSceneId As String
Private m_IsInitialized As Boolean

Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Initialize()
    If m_IsInitialized Then Exit Sub

    Set m_Scenes = New Collection
    RegisterScene New SceneGame1

    m_IsInitialized = True
End Sub

Public Sub RegisterScene(ByVal scene As IGameScene)
    If scene Is Nothing Then Exit Sub
    EnsureScenesCollection

    On Error Resume Next
    m_Scenes.Remove scene.SceneId
    On Error GoTo 0

    m_Scenes.Add scene, scene.SceneId
End Sub

Public Sub ActivateSceneById(ByVal sceneId As String)
    EnsureInitialized

    Dim scene As IGameScene
    Set scene = GetScene(sceneId)
    If scene Is Nothing Then
        Debug.Print "SceneManager: Unknown scene id " & sceneId
        Exit Sub
    End If

    SwitchScene scene
End Sub

Public Sub ActivateSceneBySheet(ByVal sheetName As String)
    EnsureInitialized

    Dim scene As IGameScene
    Set scene = FindSceneBySheet(sheetName)
    If scene Is Nothing Then
        Debug.Print "SceneManager: No scene registered for sheet " & sheetName
        Exit Sub
    End If

    SwitchScene scene
End Sub

Public Sub ApplyScreen(ByVal screenCode As String)
    EnsureInitialized

    Dim code As String
    code = NormalizeScreenCode(screenCode)
    If code = "" Then Exit Sub

    Dim scene As IGameScene
    Set scene = ResolveSceneForScreen(code)
    If scene Is Nothing Then
        Debug.Print "SceneManager: No scene registered for screen " & code
        Exit Sub
    End If

    If m_CurrentScene Is Nothing Or m_CurrentScene.SceneId <> scene.SceneId Then
        SwitchScene scene
    End If

    On Error Resume Next
    ViewportManagerInstance().FocusOnScreen code
    On Error GoTo 0

    scene.ApplyScreenLayout code
End Sub

Public Sub ApplyScreenForScene(ByVal sceneId As String, ByVal screenCode As String)
    EnsureInitialized

    Dim scene As IGameScene
    Set scene = GetScene(sceneId)
    If scene Is Nothing Then
        Debug.Print "SceneManager: Unknown scene id " & sceneId
        Exit Sub
    End If

    If m_CurrentScene Is Nothing Or m_CurrentScene.SceneId <> scene.SceneId Then
        SwitchScene scene
    End If

    Dim normalizedCode As String
    normalizedCode = NormalizeScreenCode(screenCode)

    On Error Resume Next
    ViewportManagerInstance().FocusOnScreen normalizedCode
    On Error GoTo 0

    scene.ApplyScreenLayout normalizedCode
End Sub

Public Property Get CurrentSceneId() As String
    CurrentSceneId = m_CurrentSceneId
End Property

Public Sub Reset()
    If Not m_CurrentScene Is Nothing Then
        m_CurrentScene.OnExitScene
    End If
    Set m_CurrentScene = Nothing
    m_CurrentSceneId = ""
    SafeSetGameStateScene ""
End Sub

Public Sub Destroy()
    Reset
    If Not m_Scenes Is Nothing Then
        Set m_Scenes = Nothing
    End If
    m_IsInitialized = False
End Sub

Private Sub SwitchScene(ByVal scene As IGameScene)
    If scene Is Nothing Then Exit Sub

    If Not m_CurrentScene Is Nothing Then
        If m_CurrentScene.SceneId = scene.SceneId Then Exit Sub
        m_CurrentScene.OnExitScene
    End If

    Set m_CurrentScene = scene
    m_CurrentSceneId = scene.SceneId
    m_CurrentScene.OnEnterScene
    SafeSetGameStateScene m_CurrentSceneId
End Sub

Private Function GetScene(ByVal sceneId As String) As IGameScene
    If m_Scenes Is Nothing Then Exit Function
    On Error Resume Next
    Set GetScene = m_Scenes(sceneId)
    On Error GoTo 0
End Function

Private Function FindSceneBySheet(ByVal sheetName As String) As IGameScene
    If m_Scenes Is Nothing Then Exit Function

    Dim normalized As String
    normalized = UCase$(Trim$(sheetName))

    Dim idx As Long
    For idx = 1 To m_Scenes.Count
        Dim candidate As IGameScene
        Set candidate = m_Scenes(idx)
        If UCase$(candidate.SheetName) = normalized Then
            Set FindSceneBySheet = candidate
            Exit Function
        End If
    Next idx
End Function

Private Function ResolveSceneForScreen(ByVal screenCode As String) As IGameScene
    If m_Scenes Is Nothing Then Exit Function

    Dim idx As Long
    For idx = 1 To m_Scenes.Count
        Dim candidate As IGameScene
        Set candidate = m_Scenes(idx)
        If candidate.CanHandleScreen(screenCode) Then
            Set ResolveSceneForScreen = candidate
            Exit Function
        End If
    Next idx
End Function

Private Sub EnsureInitialized()
    If Not m_IsInitialized Then Initialize
End Sub

Private Sub EnsureScenesCollection()
    If m_Scenes Is Nothing Then Set m_Scenes = New Collection
End Sub

Private Function NormalizeScreenCode(ByVal screenCode As String) As String
    NormalizeScreenCode = UCase$(Trim$(screenCode))
End Function

Private Sub SafeSetGameStateScene(ByVal sceneId As String)
    On Error Resume Next
    Dim gs As GameState
    Set gs = GameStateInstance()
    If Not gs Is Nothing Then
        gs.CurrentSceneId = sceneId
    End If
    On Error GoTo 0
End Sub
