'Attribute VB_Name = "AL_SpecialEvents"
Option Explicit

' Compatibility wrapper – delegates to SpecialEventManager

Public Sub SpecialEventTrigger(ByVal eventCode As String)
    Dim manager As SpecialEventManager
    Set manager = SpecialEventManagerInstance()
    If manager Is Nothing Then Exit Sub
    manager.Trigger eventCode
End Sub

Public Sub specialEvent0001()
    SpecialEventTrigger "XXXXXXSE0001XX"
End Sub

Public Sub specialEvent0002()
    SpecialEventTrigger "XXXXXXSE0002XX"
End Sub

Public Sub specialEvent0003()
    SpecialEventTrigger "XXXXXXSE0003XX"
End Sub

Public Sub specialEvent0004()
    SpecialEventTrigger "XXXXXXSE0004XX"
End Sub
