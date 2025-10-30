Option Explicit

' Manages screen layouts for games with multiple screens
' Currently only Link's Adventure uses this (70+ screens)

Private m_LinkScene As SceneGame1

Public Sub Initialize()
    Set m_LinkScene = New SceneGame1
End Sub

Public Sub ActivateSceneBySheet(ByVal sheetName As String)
    ' For Link's Adventure, sheet name is "Game1"
    ' Future: could support multiple games here
    If m_LinkScene Is Nothing Then Initialize
End Sub

Public Sub ApplyScreen(ByVal screenCode As String)
    If m_LinkScene Is Nothing Then Initialize
    
    Dim code As String
    code = UCase$(Trim$(screenCode))
    If code = "" Then Exit Sub
    
    ' Check if Link scene can handle this screen
    If m_LinkScene.CanHandleScreen(code) Then
        On Error Resume Next
        ViewportManagerInstance().FocusOnScreen code
        On Error GoTo 0
        
        m_LinkScene.ApplyScreenLayout code
    End If
End Sub

Public Sub Reset()
    ' No cleanup needed currently
End Sub

Public Sub Destroy()
    Set m_LinkScene = Nothing
End Sub
