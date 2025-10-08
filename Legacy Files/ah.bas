Attribute VB_Name = "AH_Enemies"
'-------------------------------
'Enemy 1

Global RNDenemyName1
Global RNDenemyFrame1_1
Global RNDenemyFrame1_2
Global RNDenemyInitialCount1
Global RNDenemyCount1

Global RNDenemyDir1
Global RNDenemySpeed1
Global RNDenemyBehaviour1
Global RNDenemyChangeRotation1

Global RNDenemyCanShoot1
Global RNDenemyChargeSpeed1

Global RNDenemyCanCollide1
Global RNDenemyCollisionDamage1
Global RNDenemyShootDamage1
Global RNDenemyChargeDamage1

Global RNDenemyHit1
Global RNDenemyLife1

'-------------------------------
'Projectile 1

Global projectileName1
Global projectileSpeed1
Global projectileBehaviour1
Global projectileDir1

'-------------------------------
'Enemy 2
Global RNDenemyName2
Global RNDenemyFrame2_1
Global RNDenemyFrame2_2
Global RNDenemyInitialCount2
Global RNDenemyCount2

Global RNDenemyDir2
Global RNDenemySpeed2
Global RNDenemyBehaviour2
Global RNDenemyChangeRotation2

Global RNDenemyCanShoot2
Global RNDenemyChargeSpeed2

Global RNDenemyCanCollide2
Global RNDenemyCollisionDamage2
Global RNDenemyShootDamage2
Global RNDenemyChargeDamage2
Global RNDenemyHit2
Global RNDenemyLife2

'-------------------------------
'Enemy 3
Global RNDenemyName3
Global RNDenemyFrame3_1
Global RNDenemyFrame3_2
Global RNDenemyInitialCount3
Global RNDenemyCount3

Global RNDenemyDir3
Global RNDenemySpeed3
Global RNDenemyBehaviour3
Global RNDenemyChangeRotation3

Global RNDenemyCanShoot3
Global RNDenemyChargeSpeed3

Global RNDenemyCanCollide3
Global RNDenemyCollisionDamage3
Global RNDenemyShootDamage3
Global RNDenemyChargeDamage3
Global RNDenemyHit3
Global RNDenemyLife3
'-------------------------------
'Enemy 4
Global RNDenemyName4
Global RNDenemyFrame4_1
Global RNDenemyFrame4_2
Global RNDenemyInitialCount4
Global RNDenemyCount4

Global RNDenemyDir4
Global RNDenemySpeed4
Global RNDenemyBehaviour4
Global RNDenemyChangeRotation4

Global RNDenemyCanShoot4
Global RNDenemyChargeSpeed4

Global RNDenemyCanCollide4
Global RNDenemyCollisionDamage4
Global RNDenemyShootDamage4
Global RNDenemyChargeDamage4
Global RNDenemyHit4
Global RNDenemyLife4


'####################### Octoroks ###################################################

Sub showOctorok01()

RNDenemyName1 = "Octorok1F1"
RNDenemyFrame1_1 = "Octorok1F1"
RNDenemyFrame1_2 = "Octorok1F2"
RNDenemyInitialCount1 = Sheets("Data").Range("I56").Value
RNDenemyCount1 = Sheets("Data").Range("I56").Value

RNDenemySpeed1 = Sheets("Data").Range("G56").Value
RNDenemyBehaviour1 = Sheets("Data").Range("J56").Value
RNDenemyChangeRotation1 = Sheets("Data").Range("K56").Value

RNDenemyCanShoot1 = Sheets("Data").Range("M56").Value
RNDenemyChargeSpeed1 = Sheets("Data").Range("O56").Value

RNDenemyCollisionDamage1 = Sheets("Data").Range("L56").Value

If RNDenemyCollisionDamage1 <> "" Then
    RNDenemyCanCollide1 = "Y"
End If

RNDenemyShootDamage1 = Sheets("Data").Range("N56").Value
RNDenemyChargeDamage1 = Sheets("Data").Range("P56").Value
RNDenemyLife1 = Sheets("Data").Range("D56").Value

ActiveSheet.Pictures("Octorok1F1").Visible = True
ActiveSheet.Pictures("Octorok1F1").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("Octorok1F1").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("Octorok1F1").Rotation = 0

Sheets("Data").Range("C56").Value = "Y"

End Sub
Sub hideOctorok01()

resetEnemy1

ActiveSheet.Shapes("Octorok1F1").Rotation = 0
ActiveSheet.Shapes("Octorok1F2").Rotation = 0

ActiveSheet.Pictures("Octorok1F1").Visible = False
ActiveSheet.Pictures("Octorok1F2").Visible = False
Sheets("Data").Range("C56").Value = "N"

End Sub


Sub showOctorok02()

RNDenemyName2 = "Octorok2F1"
RNDenemyFrame2_1 = "Octorok2F1"
RNDenemyFrame2_2 = "Octorok2F2"
RNDenemyInitialCount2 = Sheets("Data").Range("I57").Value
RNDenemyCount2 = Sheets("Data").Range("I57").Value

RNDenemySpeed2 = Sheets("Data").Range("G57").Value
RNDenemyBehaviour2 = Sheets("Data").Range("J57").Value
RNDenemyChangeRotation2 = Sheets("Data").Range("K57").Value

RNDenemyCanShoot2 = Sheets("Data").Range("M57").Value
RNDenemyChargeSpeed2 = Sheets("Data").Range("O57").Value

RNDenemyCollisionDamage2 = Sheets("Data").Range("L57").Value

If RNDenemyCollisionDamage2 <> "" Then
    RNDenemyCanCollide2 = "Y"
End If

RNDenemyShootDamage2 = Sheets("Data").Range("N57").Value
RNDenemyChargeDamage2 = Sheets("Data").Range("P57").Value
RNDenemyLife2 = Sheets("Data").Range("D57").Value

ActiveSheet.Pictures("Octorok2F1").Visible = True
ActiveSheet.Pictures("Octorok2F1").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("Octorok2F1").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("Octorok2F1").Rotation = 0

Sheets("Data").Range("C57").Value = "Y"

End Sub
Sub hideOctorok02()

resetEnemy2

ActiveSheet.Shapes("Octorok2F1").Rotation = 0
ActiveSheet.Shapes("Octorok2F2").Rotation = 0

ActiveSheet.Pictures("Octorok2F1").Visible = False
ActiveSheet.Pictures("Octorok2F2").Visible = False
Sheets("Data").Range("C57").Value = "N"


End Sub

' ######## Sandcrabs ##############################################################

Sub showsandcrab01()

RNDenemyName1 = "Sandcrab1F1"
RNDenemyFrame1_1 = "Sandcrab1F1"
RNDenemyFrame1_2 = "Sandcrab1F2"
RNDenemyInitialCount1 = Sheets("Data").Range("I54").Value
RNDenemyCount1 = Sheets("Data").Range("I54").Value

RNDenemySpeed1 = Sheets("Data").Range("G54").Value
RNDenemyBehaviour1 = Sheets("Data").Range("J54").Value
RNDenemyChangeRotation1 = Sheets("Data").Range("K54").Value

RNDenemyCanShoot1 = Sheets("Data").Range("M54").Value
RNDenemyChargeSpeed1 = Sheets("Data").Range("O54").Value

RNDenemyCollisionDamage1 = Sheets("Data").Range("L54").Value

If RNDenemyCollisionDamage1 <> "" Then
    RNDenemyCanCollide1 = "Y"
End If

RNDenemyShootDamage1 = Sheets("Data").Range("N54").Value
RNDenemyChargeDamage1 = Sheets("Data").Range("P54").Value
RNDenemyLife1 = Sheets("Data").Range("D54").Value

ActiveSheet.Pictures("Sandcrab1F1").Visible = True
ActiveSheet.Pictures("Sandcrab1F1").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("Sandcrab1F1").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("Sandcrab1F1").Rotation = 0

Sheets("Data").Range("C54").Value = "Y"

End Sub

Sub hidesandcrab01()

resetEnemy1

ActiveSheet.Pictures("Sandcrab1F1").Visible = False
ActiveSheet.Pictures("Sandcrab1F2").Visible = False
Sheets("Data").Range("C54").Value = "N"

End Sub

Sub showsandcrab02()

RNDenemyName2 = "Sandcrab2F1"
RNDenemyFrame2_1 = "Sandcrab2F1"
RNDenemyFrame2_2 = "Sandcrab2F2"
RNDenemyInitialCount2 = Sheets("Data").Range("I55").Value
RNDenemyCount2 = Sheets("Data").Range("I55").Value

RNDenemySpeed2 = Sheets("Data").Range("G55").Value
RNDenemyBehaviour2 = Sheets("Data").Range("J55").Value
RNDenemyChangeRotation2 = Sheets("Data").Range("K55").Value

RNDenemyCanShoot2 = Sheets("Data").Range("M55").Value
RNDenemyChargeSpeed2 = Sheets("Data").Range("O55").Value

RNDenemyCollisionDamage2 = Sheets("Data").Range("L55").Value

If RNDenemyCollisionDamage2 <> "" Then
    RNDenemyCanCollide2 = "Y"
End If

RNDenemyShootDamage2 = Sheets("Data").Range("N55").Value
RNDenemyChargeDamage2 = Sheets("Data").Range("P55").Value
RNDenemyLife2 = Sheets("Data").Range("D55").Value

ActiveSheet.Pictures("Sandcrab2F1").Visible = True
ActiveSheet.Pictures("Sandcrab2F1").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("Sandcrab2F1").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("Sandcrab2F1").Rotation = 0

Sheets("Data").Range("C55").Value = "Y"

End Sub

Sub hidesandcrab02()

resetEnemy2

ActiveSheet.Pictures("Sandcrab2F1").Visible = False
ActiveSheet.Pictures("Sandcrab2F2").Visible = False
Sheets("Data").Range("C55").Value = "N"

End Sub

'######### Gordos ########################

Sub showgordo01()

RNDenemyName1 = "Gordo1F1"
RNDenemyFrame1_1 = "Gordo1F1"
RNDenemyFrame1_2 = "Gordo1F2"
RNDenemyInitialCount1 = Sheets("Data").Range("I60").Value
RNDenemyCount1 = Sheets("Data").Range("I60").Value

RNDenemySpeed1 = Sheets("Data").Range("G60").Value
RNDenemyBehaviour1 = Sheets("Data").Range("J60").Value
RNDenemyChangeRotation1 = Sheets("Data").Range("K60").Value

RNDenemyCanShoot1 = Sheets("Data").Range("M60").Value
RNDenemyChargeSpeed1 = Sheets("Data").Range("O60").Value

RNDenemyCollisionDamage1 = Sheets("Data").Range("L60").Value

If RNDenemyCollisionDamage1 <> "" Then
    RNDenemyCanCollide1 = "Y"
End If

RNDenemyShootDamage1 = Sheets("Data").Range("N60").Value
RNDenemyChargeDamage1 = Sheets("Data").Range("P60").Value
RNDenemyLife1 = Sheets("Data").Range("D60").Value


ActiveSheet.Pictures("Gordo1F1").Visible = True
ActiveSheet.Pictures("Gordo1F1").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("Gordo1F1").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("Gordo1F1").Rotation = 0

Sheets("Data").Range("C60").Value = "Y"

End Sub
Sub hidegordo01()

resetEnemy1

ActiveSheet.Pictures("Gordo1F1").Visible = False
ActiveSheet.Pictures("Gordo1F2").Visible = False
Sheets("Data").Range("C60").Value = "N"

End Sub
Sub showgordo02()

RNDenemyName2 = "Gordo2F1"
RNDenemyFrame2_1 = "Gordo2F1"
RNDenemyFrame2_2 = "Gordo2F2"
RNDenemyInitialCount2 = Sheets("Data").Range("I61").Value
RNDenemyCount2 = Sheets("Data").Range("I61").Value

RNDenemySpeed2 = Sheets("Data").Range("G61").Value
RNDenemyBehaviour2 = Sheets("Data").Range("J61").Value
RNDenemyChangeRotation2 = Sheets("Data").Range("K61").Value

RNDenemyCanShoot2 = Sheets("Data").Range("M61").Value
RNDenemyChargeSpeed2 = Sheets("Data").Range("O61").Value

RNDenemyCollisionDamage2 = Sheets("Data").Range("L61").Value

If RNDenemyCollisionDamage2 <> "" Then
    RNDenemyCanCollide2 = "Y"
End If

RNDenemyShootDamage2 = Sheets("Data").Range("N61").Value
RNDenemyChargeDamage2 = Sheets("Data").Range("P61").Value
RNDenemyLife2 = Sheets("Data").Range("D61").Value

ActiveSheet.Pictures("Gordo2F1").Visible = True
ActiveSheet.Pictures("Gordo2F1").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("Gordo2F1").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("Gordo2F1").Rotation = 0

Sheets("Data").Range("C61").Value = "Y"

End Sub
Sub hidegordo02()

resetEnemy2

ActiveSheet.Pictures("Gordo2F1").Visible = False
ActiveSheet.Pictures("Gordo2F2").Visible = False
Sheets("Data").Range("C61").Value = "N"

End Sub
Sub showgordo03()

RNDenemyName3 = "Gordo3F1"
RNDenemyFrame3_1 = "Gordo3F1"
RNDenemyFrame3_2 = "Gordo3F2"
RNDenemyInitialCount3 = Sheets("Data").Range("I62").Value
RNDenemyCount3 = Sheets("Data").Range("I62").Value

RNDenemySpeed3 = Sheets("Data").Range("G62").Value
RNDenemyBehaviour3 = Sheets("Data").Range("J62").Value
RNDenemyChangeRotation3 = Sheets("Data").Range("K62").Value

RNDenemyCanShoot3 = Sheets("Data").Range("M62").Value
RNDenemyChargeSpeed3 = Sheets("Data").Range("O62").Value

RNDenemyCollisionDamage3 = Sheets("Data").Range("L62").Value

If RNDenemyCollisionDamage3 <> "" Then
    RNDenemyCanCollide3 = "Y"
End If

RNDenemyShootDamage3 = Sheets("Data").Range("N62").Value
RNDenemyChargeDamage3 = Sheets("Data").Range("P62").Value
RNDenemyLife3 = Sheets("Data").Range("D62").Value

ActiveSheet.Pictures("Gordo3F1").Visible = True
ActiveSheet.Pictures("Gordo3F1").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("Gordo3F1").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("Gordo3F1").Rotation = 0

Sheets("Data").Range("C62").Value = "Y"

End Sub
Sub hidegordo03()

resetEnemy3

ActiveSheet.Pictures("Gordo3F1").Visible = False
ActiveSheet.Pictures("Gordo3F2").Visible = False
Sheets("Data").Range("C62").Value = "N"

End Sub


