Option Explicit

'===================================================================================
'                              IBEHAVIOR INTERFACE
'===================================================================================
' Interface for all entity behaviors (AI, dialogue, pickups, etc.)
' Implement this interface to create custom entity behaviors
'
' USAGE:
'   Option Explicit
'   Implements IBehavior
'
'   Private Sub IBehavior_Update(entity As Entity, deltaTime As Double)
'       ' Your behavior logic here
'   End Sub
'===================================================================================

' Called every frame - implement your behavior logic here
Public Sub Update(entity As Entity, deltaTime As Double)
    ' Must be overridden by implementing class
End Sub
