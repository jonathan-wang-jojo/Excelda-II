'Attribute VB_Name = "AK_Objects"
Option Explicit

'===================================================================================
' Module: AK_Objects
' Purpose: Legacy compatibility wrappers that delegate to ObjectManager while
'          preserving historical macro signatures used by screen setup sheets.
'===================================================================================

Public Sub swordHitBush(ByVal swordImage As Shape)
    Dim manager As ObjectManager
    Set manager = ObjectManagerInstance()
    If manager Is Nothing Then Exit Sub
    manager.HandleSwordHit swordImage
End Sub

Public Sub resetBushes(ByVal whichBush As String)
    Dim manager As ObjectManager
    Set manager = ObjectManagerInstance()
    If manager Is Nothing Then Exit Sub
    manager.ResetObjects whichBush
End Sub

Public Sub positionObj(ByVal objectName As String, ByVal objectLocation As String, ByVal cellValue As Variant)
    Dim manager As ObjectManager
    Set manager = ObjectManagerInstance()
    If manager Is Nothing Then Exit Sub
    manager.PositionObject objectName, objectLocation, cellValue
End Sub

Public Sub positionMultiple(ByVal objectType As String, _
                            ByVal Obj1 As Variant, ByVal Obj2 As Variant, ByVal Obj3 As Variant, ByVal Obj4 As Variant, _
                            ByVal Obj5 As Variant, ByVal Obj6 As Variant, ByVal Obj7 As Variant, ByVal Obj8 As Variant, _
                            ByVal Obj9 As Variant, ByVal Obj10 As Variant, ByVal Obj11 As Variant, ByVal Obj12 As Variant, _
                            ByVal Obj13 As Variant, ByVal Obj14 As Variant, ByVal Obj15 As Variant, ByVal Obj16 As Variant, _
                            ByVal Obj17 As Variant, ByVal Obj18 As Variant, ByVal Obj19 As Variant, ByVal Obj20 As Variant, _
                            ByVal Obj21 As Variant, ByVal Obj22 As Variant, ByVal Obj23 As Variant, ByVal Obj24 As Variant, _
                            ByVal Obj25 As Variant, ByVal Obj26 As Variant, ByVal Obj27 As Variant, ByVal Obj28 As Variant, _
                            ByVal Obj29 As Variant, ByVal Obj30 As Variant)

    Dim manager As ObjectManager
    Set manager = ObjectManagerInstance()
    If manager Is Nothing Then Exit Sub

    manager.PositionMultiple objectType, _
        Obj1, Obj2, Obj3, Obj4, Obj5, Obj6, Obj7, Obj8, Obj9, Obj10, _
        Obj11, Obj12, Obj13, Obj14, Obj15, Obj16, Obj17, Obj18, Obj19, Obj20, _
        Obj21, Obj22, Obj23, Obj24, Obj25, Obj26, Obj27, Obj28, Obj29, Obj30
End Sub

Public Sub getHeartPiece(ByVal heartNum As Variant)
    Dim manager As ObjectManager
    Set manager = ObjectManagerInstance()
    If manager Is Nothing Then Exit Sub
    manager.CollectHeartPiece CLng(Val(heartNum))
End Sub