'Attribute VB_Name = "AI_Friendlies"
Sub showMarin01()

RNDenemyName1 = "MarinD"
RNDenemyFrame1_1 = "MarinD"
RNDenemyFrame1_2 = "MarinD"
RNDenemyInitialCount1 = Sheets("Data").Range("I46").Value
RNDenemyCount1 = Sheets("Data").Range("I46").Value

RNDenemySpeed1 = Sheets("Data").Range("G46").Value
RNDenemyBehaviour1 = Sheets("Data").Range("J46").Value
RNDenemyChangeRotation1 = Sheets("Data").Range("K46").Value

RNDenemyCanShoot1 = Sheets("Data").Range("M46").Value
RNDenemyChargeSpeed1 = Sheets("Data").Range("O46").Value

RNDenemyCollisionDamage1 = Sheets("Data").Range("L46").Value

If RNDenemyCollisionDamage1 <> "" Then
    RNDenemyCanCollide1 = "Y"
End If

RNDenemyShootDamage1 = Sheets("Data").Range("N46").Value
RNDenemyChargeDamage1 = Sheets("Data").Range("P46").Value
RNDenemyLife1 = Sheets("Data").Range("D46").Value

ActiveSheet.Pictures("MarinD").Visible = True
ActiveSheet.Pictures("MarinD").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("MarinD").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("MarinU").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("MarinU").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("MarinL").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("MarinL").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("MarinR").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("MarinR").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("MarinD").Rotation = 0
ActiveSheet.Shapes("MarinU").Rotation = 0
ActiveSheet.Shapes("MarinL").Rotation = 0
ActiveSheet.Shapes("MarinR").Rotation = 0

Sheets("Data").Range("C46").Value = "Y"

End Sub

Sub hideMarin01()

resetEnemy1

ActiveSheet.Shapes("MarinU").Rotation = 0
ActiveSheet.Shapes("MarinD").Rotation = 0
ActiveSheet.Shapes("MarinR").Rotation = 0
ActiveSheet.Shapes("MarinL").Rotation = 0

ActiveSheet.Pictures("MarinU").Visible = False
ActiveSheet.Pictures("MarinD").Visible = False
ActiveSheet.Pictures("MarinL").Visible = False
ActiveSheet.Pictures("MarinR").Visible = False

Sheets("Data").Range("C46").Value = "N"

End Sub


Sub showTarin02()

RNDenemyName2 = "TarinD"
RNDenemyFrame2_1 = "TarinD"
RNDenemyFrame2_2 = "TarinD"
RNDenemyInitialCount2 = Sheets("Data").Range("I47").Value
RNDenemyCount2 = Sheets("Data").Range("I47").Value

RNDenemySpeed2 = Sheets("Data").Range("G47").Value
RNDenemyBehaviour2 = Sheets("Data").Range("J47").Value
RNDenemyChangeRotation2 = Sheets("Data").Range("K47").Value

RNDenemyCanShoot2 = Sheets("Data").Range("M47").Value
RNDenemyChargeSpeed2 = Sheets("Data").Range("O47").Value

RNDenemyCollisionDamage2 = Sheets("Data").Range("L47").Value

If RNDenemyCollisionDamage2 <> "" Then
    RNDenemyCanCollide2 = "Y"
End If

RNDenemyShootDamage2 = Sheets("Data").Range("N47").Value
RNDenemyChargeDamage2 = Sheets("Data").Range("P47").Value
RNDenemyLife2 = Sheets("Data").Range("D47").Value

ActiveSheet.Pictures("TarinD").Visible = True
ActiveSheet.Pictures("TarinD").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("TarinD").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("TarinU").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("TarinU").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("TarinL").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("TarinL").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("TarinR").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("TarinR").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("TarinD").Rotation = 0
ActiveSheet.Shapes("TarinU").Rotation = 0
ActiveSheet.Shapes("TarinL").Rotation = 0
ActiveSheet.Shapes("TarinR").Rotation = 0

Sheets("Data").Range("C47").Value = "Y"

End Sub

Sub hideTarin02()

resetEnemy2

ActiveSheet.Shapes("TarinU").Rotation = 0
ActiveSheet.Shapes("TarinD").Rotation = 0
ActiveSheet.Shapes("TarinR").Rotation = 0
ActiveSheet.Shapes("TarinL").Rotation = 0

ActiveSheet.Pictures("TarinU").Visible = False
ActiveSheet.Pictures("TarinD").Visible = False
ActiveSheet.Pictures("TarinL").Visible = False
ActiveSheet.Pictures("TarinR").Visible = False

Sheets("Data").Range("C47").Value = "N"

End Sub

Sub showRaccoon01()

RNDenemyName1 = "RaccoonD"
RNDenemyFrame1_1 = "RaccoonD"
RNDenemyFrame1_2 = "RaccoonD"
RNDenemyInitialCount1 = Sheets("Data").Range("I49").Value
RNDenemyCount1 = Sheets("Data").Range("I49").Value

RNDenemySpeed1 = Sheets("Data").Range("G49").Value
RNDenemyBehaviour1 = Sheets("Data").Range("J49").Value
RNDenemyChangeRotation1 = Sheets("Data").Range("K49").Value

RNDenemyCanShoot1 = Sheets("Data").Range("M49").Value
RNDenemyChargeSpeed1 = Sheets("Data").Range("O49").Value

RNDenemyCollisionDamage1 = Sheets("Data").Range("L49").Value

If RNDenemyCollisionDamage1 <> "" Then
    RNDenemyCanCollide1 = "Y"
End If

RNDenemyShootDamage1 = Sheets("Data").Range("N49").Value
RNDenemyChargeDamage1 = Sheets("Data").Range("P49").Value
RNDenemyLife1 = Sheets("Data").Range("D49").Value

ActiveSheet.Pictures("RaccoonD").Visible = True
ActiveSheet.Pictures("RaccoonD").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("RaccoonD").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("RaccoonU").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("RaccoonU").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("RaccoonL").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("RaccoonL").Left = Range(TriggerCel).Left
ActiveSheet.Pictures("RaccoonR").Top = Range(TriggerCel).Top
ActiveSheet.Pictures("RaccoonR").Left = Range(TriggerCel).Left
ActiveSheet.Shapes("RaccoonD").Rotation = 0
ActiveSheet.Shapes("RaccoonU").Rotation = 0
ActiveSheet.Shapes("RaccoonL").Rotation = 0
ActiveSheet.Shapes("RaccoonR").Rotation = 0

Sheets("Data").Range("C49").Value = "Y"

End Sub

Sub hideRaccoon01()

resetEnemy1

ActiveSheet.Shapes("RaccoonU").Rotation = 0
ActiveSheet.Shapes("RaccoonD").Rotation = 0
ActiveSheet.Shapes("RaccoonR").Rotation = 0
ActiveSheet.Shapes("RaccoonL").Rotation = 0

ActiveSheet.Pictures("RaccoonU").Visible = False
ActiveSheet.Pictures("RaccoonD").Visible = False
ActiveSheet.Pictures("RaccoonL").Visible = False
ActiveSheet.Pictures("RaccoonR").Visible = False

Sheets("Data").Range("C46").Value = "N"

End Sub
