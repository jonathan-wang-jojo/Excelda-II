'Attribute VB_Name = "AI_Friendlies"
Option Explicit

Public Sub showMarin01()
    SpawnFriendlySlot 1, "Marin", 46
End Sub

Public Sub hideMarin01()
    DespawnFriendlySlot 1, 46
    ResetEnemySlot 1, 46
End Sub

Public Sub showTarin02()
    SpawnFriendlySlot 2, "Tarin", 47
End Sub

Public Sub hideTarin02()
    DespawnFriendlySlot 2, 47
    ResetEnemySlot 2, 47
End Sub

Public Sub showRaccoon01()
    SpawnFriendlySlot 1, "Raccoon", 49
End Sub

Public Sub hideRaccoon01()
    DespawnFriendlySlot 1, 49
    ResetEnemySlot 1, 49
End Sub

Private Sub SpawnFriendlySlot(ByVal slot As Long, ByVal baseName As String, ByVal dataRow As Long)
    Dim manager As FriendlyManager
    Set manager = FriendlyManagerInstance()
    If manager Is Nothing Then Exit Sub

    Dim anchor As Range
    Set anchor = TriggerAnchorCell()
    If anchor Is Nothing Then Exit Sub

    manager.SpawnFriendly slot, baseName, dataRow, anchor
End Sub

Private Sub DespawnFriendlySlot(ByVal slot As Long, ByVal dataRow As Long)
    Dim manager As FriendlyManager
    Set manager = FriendlyManagerInstance()
    If manager Is Nothing Then Exit Sub

    manager.DespawnFriendly slot, dataRow
End Sub

Private Sub ResetEnemySlot(ByVal slot As Long, ByVal dataRow As Long)
    Dim enemyManager As EnemyManager
    Set enemyManager = EnemyManagerInstance()
    If enemyManager Is Nothing Then Exit Sub

    enemyManager.DespawnEnemy slot, dataRow
End Sub

Private Function TriggerAnchorCell() As Range
    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Then Exit Function

    Dim address As String
    address = Trim$(gs.TriggerCellAddress)
    If address = "" Then Exit Function

    Dim hostSheet As Worksheet
    On Error Resume Next
    If gs.CurrentScreen <> "" Then
        Set hostSheet = Sheets(gs.CurrentScreen)
    Else
        Set hostSheet = ActiveSheet
    End If

    If hostSheet Is Nothing Then
        Set TriggerAnchorCell = Nothing
    Else
        Set TriggerAnchorCell = hostSheet.Range(address)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Set TriggerAnchorCell = Nothing
    End If
    On Error GoTo 0
End Function
