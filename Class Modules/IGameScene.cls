'Attribute VB_Name = "IGameScene"
Option Explicit

' Interface for scene-specific layout logic

Public Property Get SceneId() As String
End Property

Public Property Get SheetName() As String
End Property

Public Function CanHandleScreen(ByVal screenCode As String) As Boolean
End Function

Public Sub ApplyScreenLayout(ByVal screenCode As String)
End Sub

Public Sub OnEnterScene()
End Sub

Public Sub OnExitScene()
End Sub
