VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enemy"
Option Explicit

'Simple Enemy class to hold properties for migration.
Private pName As String
Private pFrame1 As String
Private pFrame2 As String
Private pInitialCount As Variant
Private pCount As Variant
Private pDir As Variant
Private pSpeed As Variant
Private pBehaviour As Variant
Private pChangeRotation As Variant
Private pCanShoot As Variant
Private pChargeSpeed As Variant
Private pCanCollide As Variant
Private pCollisionDamage As Variant
Private pShootDamage As Variant
Private pChargeDamage As Variant
Private pHit As Variant
Private pLife As Variant

Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(val As String)
    pName = val
End Property

Public Property Get Frame1() As String
    Frame1 = pFrame1
End Property
Public Property Let Frame1(val As String)
    pFrame1 = val
End Property

Public Property Get Frame2() As String
    Frame2 = pFrame2
End Property
Public Property Let Frame2(val As String)
    pFrame2 = val
End Property

Public Sub InitializeFromGlobals(index As Integer)
    'Copy values from the existing globals (ZZ_Globals) into this object
    Select Case index
        Case 1
            pName = RNDenemyName1
            pFrame1 = RNDenemyFrame1_1
            pFrame2 = RNDenemyFrame1_2
            pInitialCount = RNDenemyInitialCount1
            pCount = RNDenemyCount1
            pDir = RNDenemyDir1
            pSpeed = RNDenemySpeed1
            pBehaviour = RNDenemyBehaviour1
            pChangeRotation = RNDenemyChangeRotation1
            pCanShoot = RNDenemyCanShoot1
            pChargeSpeed = RNDenemyChargeSpeed1
            pCanCollide = RNDenemyCanCollide1
            pCollisionDamage = RNDenemyCollisionDamage1
            pShootDamage = RNDenemyShootDamage1
            pChargeDamage = RNDenemyChargeDamage1
            pHit = RNDenemyHit1
            pLife = RNDenemyLife1
        Case 2
            pName = RNDenemyName2
            pFrame1 = RNDenemyFrame2_1
            pFrame2 = RNDenemyFrame2_2
            pInitialCount = RNDenemyInitialCount2
            pCount = RNDenemyCount2
            pDir = RNDenemyDir2
            pSpeed = RNDenemySpeed2
            pBehaviour = RNDenemyBehaviour2
            pChangeRotation = RNDenemyChangeRotation2
            pCanShoot = RNDenemyCanShoot2
            pChargeSpeed = RNDenemyChargeSpeed2
            pCanCollide = RNDenemyCanCollide2
            pCollisionDamage = RNDenemyCollisionDamage2
            pShootDamage = RNDenemyShootDamage2
            pChargeDamage = RNDenemyChargeDamage2
            pHit = RNDenemyHit2
            pLife = RNDenemyLife2
        Case 3
            pName = RNDenemyName3
            pFrame1 = RNDenemyFrame3_1
            pFrame2 = RNDenemyFrame3_2
            pInitialCount = RNDenemyInitialCount3
            pCount = RNDenemyCount3
            pDir = RNDenemyDir3
            pSpeed = RNDenemySpeed3
            pBehaviour = RNDenemyBehaviour3
            pChangeRotation = RNDenemyChangeRotation3
            pCanShoot = RNDenemyCanShoot3
            pChargeSpeed = RNDenemyChargeSpeed3
            pCanCollide = RNDenemyCanCollide3
            pCollisionDamage = RNDenemyCollisionDamage3
            pShootDamage = RNDenemyShootDamage3
            pChargeDamage = RNDenemyChargeDamage3
            pHit = RNDenemyHit3
            pLife = RNDenemyLife3
        Case 4
            pName = RNDenemyName4
            pFrame1 = RNDenemyFrame4_1
            pFrame2 = RNDenemyFrame4_2
            pInitialCount = RNDenemyInitialCount4
            pCount = RNDenemyCount4
            pDir = RNDenemyDir4
            pSpeed = RNDenemySpeed4
            pBehaviour = RNDenemyBehaviour4
            pChangeRotation = RNDenemyChangeRotation4
            pCanShoot = RNDenemyCanShoot4
            pChargeSpeed = RNDenemyChargeSpeed4
            pCanCollide = RNDenemyCanCollide4
            pCollisionDamage = RNDenemyCollisionDamage4
            pShootDamage = RNDenemyShootDamage4
            pChargeDamage = RNDenemyChargeDamage4
            pHit = RNDenemyHit4
            pLife = RNDenemyLife4
        Case Else
            'Do nothing
    End Select
End Sub

Public Function ToString() As String
    ToString = "Enemy: " & pName & ", Life=" & CStr(pLife)
End Function

Public Sub Reset()
    'Object-level reset: hide frames if they exist and clear internal fields
    On Error Resume Next
    If pFrame1 <> "" Then ActiveSheet.Shapes(pFrame1).Visible = False
    If pFrame2 <> "" Then ActiveSheet.Shapes(pFrame2).Visible = False

    pName = ""
    pFrame1 = ""
    pFrame2 = ""
    pInitialCount = ""
    pCount = ""
    pDir = ""
    pSpeed = ""
    pBehaviour = ""
    pChangeRotation = ""
    pCanShoot = ""
    pChargeSpeed = ""
    pCanCollide = ""
    pCollisionDamage = ""
    pShootDamage = ""
    pChargeDamage = ""
    pHit = ""
    pLife = ""
End Sub
