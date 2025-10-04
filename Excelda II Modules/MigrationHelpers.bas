Option Explicit
'Helpers to build OOP objects from legacy globals. Non-invasive: reads globals and builds class instances.
Public Enemies As Collection

Sub BuildEnemyObjects()
    On Error GoTo EH
    Set Enemies = New Collection
    Dim i As Integer
    Dim e As Enemy
    For i = 1 To 4
        Set e = New Enemy
        e.InitializeFromGlobals i
        Enemies.Add e
    Next i
    MsgBox "Built " & Enemies.Count & " enemy objects (values copied from globals)."
    Exit Sub
EH:
    MsgBox "BuildEnemyObjects error: " & Err.Number & " - " & Err.Description
End Sub

Sub DumpEnemyObjects()
    If Enemies Is Nothing Then
        MsgBox "No enemy objects built. Run BuildEnemyObjects first."
        Exit Sub
    End If
    Dim idx As Long
    For idx = 1 To Enemies.Count
        Debug.Print Enemies(idx).ToString
    Next idx
    MsgBox "Dump complete (see Immediate window)."
End Sub

Sub CreateEnemy1Instance()
    'Create or replace enemy slot 1 as an Enemy object and store in Enemies(1)
    On Error GoTo EH
    If Enemies Is Nothing Then Set Enemies = New Collection
    Dim e As Enemy
    Set e = New Enemy
    e.InitializeFromGlobals 1

    'Place or replace at index 1
    If Enemies.Count >= 1 Then
        'replace
        Set Enemies(1) = e
    Else
        Enemies.Add e
    End If

    MsgBox "Enemy object for slot 1 created. Use DumpEnemyObjects to inspect."
    Exit Sub
EH:
    MsgBox "CreateEnemy1Instance error: " & Err.Number & " - " & Err.Description
End Sub

