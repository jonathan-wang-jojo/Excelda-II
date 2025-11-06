'@Folder("Engine.Core")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' ENGINE EVENTS
'═══════════════════════════════════════════════════════════════════════════════

Public Sub SpecialEventTrigger(ByVal eventCode As String)
    Dim manager As SpecialEventManager
    Set manager = SpecialEventManagerInstance()
    If manager Is Nothing Then Exit Sub
    manager.Trigger eventCode
End Sub
