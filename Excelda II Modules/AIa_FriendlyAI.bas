'Attribute VB_Name = "AIa_FriendlyAI"
Option Explicit

Public Sub moveStillFollow(ByVal slot As Long)
    Dim manager As FriendlyManager
    Set manager = FriendlyManagerInstance()
    If manager Is Nothing Then Exit Sub
    If Not manager.FriendlyIsActive(slot) Then Exit Sub

    Dim friendlyShape As Shape
    Set friendlyShape = manager.GetFriendlyShape(slot)
    If friendlyShape Is Nothing Then Exit Sub

    Dim linkShape As Shape
    On Error Resume Next
    Set linkShape = LinkSprite
    On Error GoTo 0
    If linkShape Is Nothing Then Exit Sub

    Dim desiredDir As String
    desiredDir = manager.FriendlyDirection(slot)

    If linkShape.Top < friendlyShape.Top Then
        desiredDir = "U"
    ElseIf linkShape.Top > friendlyShape.Top + 60 Then
        desiredDir = "D"
    ElseIf linkShape.Left < friendlyShape.Left Then
        desiredDir = "L"
    ElseIf linkShape.Left > friendlyShape.Left + 30 Then
        desiredDir = "R"
    End If

    manager.SetFriendlyDirection slot, desiredDir
End Sub
