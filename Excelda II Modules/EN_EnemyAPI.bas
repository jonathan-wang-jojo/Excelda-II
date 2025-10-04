Option Explicit
Attribute VB_Name = "EN_EnemyAPI"
'Enemy API helpers - non-invasive accessors and utilities for legacy enemy globals.
'This module is intended to centralize reads/writes for enemy slots (1-4) so
'future refactors can replace global usage with object-backed logic.

'Note: This module does not change existing behavior; it provides alternate
'callable helpers. Existing code still uses the globals directly.

Public Sub ResetEnemyGeneric(index As Integer)
    Dim myFrame1 As Shape, myFrame2 As Shape
    On Error GoTo EH
    Select Case index
        Case 1
            'If an object exists for enemy slot 1, allow it to handle reset
            On Error Resume Next
            If Not Enemies Is Nothing Then
                If Enemies.Count >= 1 Then
                    'Attempt to call a Reset method on the object if it exists
                    If TypeName(Enemies(1)) = "Enemy" Then
                        On Error Resume Next
                        CallByName Enemies(1), "Reset", VbMethod
                        If Err.Number = 0 Then Exit Sub
                        On Error GoTo EH
                    End If
                End If
            End If

            On Error GoTo EH
            If RNDenemyName1 <> "" Then
                Set myFrame1 = ActiveSheet.Shapes(RNDenemyFrame1_1)
                Set myFrame2 = ActiveSheet.Shapes(RNDenemyFrame1_2)
                myFrame1.Rotation = 0
                myFrame2.Rotation = 0
                myFrame1.Visible = False
                myFrame2.Visible = False

                RNDenemyName1 = ""
                RNDenemyFrame1_1 = ""
                RNDenemyFrame1_2 = ""
                RNDenemyInitialCount1 = ""
                RNDenemyCount1 = ""

                RNDenemySpeed1 = ""
                RNDenemyBehaviour1 = ""
                RNDenemyChangeRotation1 = ""

                RNDenemyCanShoot1 = ""
                RNDenemyChargeSpeed1 = ""

                RNDenemyCanCollide1 = ""
                RNDenemyCollisionDamage1 = ""
                RNDenemyShootDamage1 = ""
                RNDenemyChargeDamage1 = ""
                RNDenemyHit1 = ""
                RNDenemyLife1 = ""
            End If
        Case 2
            If RNDenemyName2 <> "" Then
                Set myFrame1 = ActiveSheet.Shapes(RNDenemyFrame2_1)
                Set myFrame2 = ActiveSheet.Shapes(RNDenemyFrame2_2)
                myFrame1.Rotation = 0
                myFrame2.Rotation = 0
                myFrame1.Visible = False
                myFrame2.Visible = False

                RNDenemyName2 = ""
                RNDenemyFrame2_1 = ""
                RNDenemyFrame2_2 = ""
                RNDenemyInitialCount2 = ""
                RNDenemyCount2 = ""

                RNDenemySpeed2 = ""
                RNDenemyBehaviour2 = ""
                RNDenemyChangeRotation2 = ""

                RNDenemyCanShoot2 = ""
                RNDenemyChargeSpeed2 = ""

                RNDenemyCanCollide2 = ""
                RNDenemyCollisionDamage2 = ""
                RNDenemyShootDamage2 = ""
                RNDenemyChargeDamage2 = ""
                RNDenemyHit2 = ""
                RNDenemyLife2 = ""
            End If
        Case 3
            If RNDenemyName3 <> "" Then
                Set myFrame1 = ActiveSheet.Shapes(RNDenemyFrame3_1)
                Set myFrame2 = ActiveSheet.Shapes(RNDenemyFrame3_2)
                myFrame1.Rotation = 0
                myFrame2.Rotation = 0
                myFrame1.Visible = False
                myFrame2.Visible = False

                RNDenemyName3 = ""
                RNDenemyFrame3_1 = ""
                RNDenemyFrame3_2 = ""
                RNDenemyInitialCount3 = ""
                RNDenemyCount3 = ""

                RNDenemySpeed3 = ""
                RNDenemyBehaviour3 = ""
                RNDenemyChangeRotation3 = ""

                RNDenemyCanShoot3 = ""
                RNDenemyChargeSpeed3 = ""

                RNDenemyCanCollide3 = ""
                RNDenemyCollisionDamage3 = ""
                RNDenemyShootDamage3 = ""
                RNDenemyChargeDamage3 = ""
                RNDenemyHit3 = ""
                RNDenemyLife3 = ""
            End If
        Case 4
            If RNDenemyName4 <> "" Then
                Set myFrame1 = ActiveSheet.Shapes(RNDenemyFrame4_1)
                Set myFrame2 = ActiveSheet.Shapes(RNDenemyFrame4_2)
                myFrame1.Rotation = 0
                myFrame2.Rotation = 0
                myFrame1.Visible = False
                myFrame2.Visible = False

                RNDenemyName4 = ""
                RNDenemyFrame4_1 = ""
                RNDenemyFrame4_2 = ""
                RNDenemyInitialCount4 = ""
                RNDenemyCount4 = ""

                RNDenemySpeed4 = ""
                RNDenemyBehaviour4 = ""
                RNDenemyChangeRotation4 = ""

                RNDenemyCanShoot4 = ""
                RNDenemyChargeSpeed4 = ""

                RNDenemyCanCollide4 = ""
                RNDenemyCollisionDamage4 = ""
                RNDenemyShootDamage4 = ""
                RNDenemyChargeDamage4 = ""
                RNDenemyHit4 = ""
                RNDenemyLife4 = ""
            End If
    End Select
    Exit Sub
EH:
    'Non-fatal: if shapes missing, just exit
    Exit Sub
End Sub

Public Function GetEnemyName(index As Integer) As Variant
    Select Case index
        Case 1: GetEnemyName = RNDenemyName1
        Case 2: GetEnemyName = RNDenemyName2
        Case 3: GetEnemyName = RNDenemyName3
        Case 4: GetEnemyName = RNDenemyName4
        Case Else: GetEnemyName = ""
    End Select
End Function

Public Sub SetEnemyName(index As Integer, val As Variant)
    Select Case index
        Case 1: RNDenemyName1 = val
        Case 2: RNDenemyName2 = val
        Case 3: RNDenemyName3 = val
        Case 4: RNDenemyName4 = val
    End Select
End Sub
