'Attribute VB_Name = "AC_Game1Screens"
'###################################################################################
'#
'#
'#      QUADRANTS - Screen setups
'#
'#
'###################################################################################

'The following macros are named after the quadrants that describe each screen on the
'Game1 sheet.  As the screen is scrolled or relocated, the appropriate macro is
'called which describes how the screen is set up (enemies, NPCs, rocks, bushes etc).

Sub AA()
Call ResetAllEnemies
Call resetBushes("All")

End Sub

Sub AB()
Call ResetAllEnemies
Call resetBushes("All")

End Sub


Sub AC()
Call ResetAllEnemies
Call resetBushes("All")

End Sub

Sub AD()
Call ResetAllEnemies


End Sub

Sub BA()
Call ResetAllEnemies



End Sub

Sub BB()
Call ResetAllEnemies



End Sub


Sub BC()
Call ResetAllEnemies


End Sub

Sub BD()
Call ResetAllEnemies


End Sub

Sub CA()
Call ResetAllEnemies


End Sub

Sub CB()
Call ResetAllEnemies


End Sub

Sub CC()
Call ResetAllEnemies


End Sub
Sub CD()
Call ResetAllEnemies


End Sub

Sub DA()
Call ResetAllEnemies


End Sub

Sub DB()
Call ResetAllEnemies


End Sub

Sub DC()
Call ResetAllEnemies


End Sub

Sub DD()
Call ResetAllEnemies


End Sub

Sub EA()
Call ResetAllEnemies


End Sub

Sub EB()
Call ResetAllEnemies


End Sub

Sub EC()
Call ResetAllEnemies


End Sub

Sub ED()
Call ResetAllEnemies


End Sub

Sub FA()
Call ResetAllEnemies


End Sub

Sub FB()
Call ResetAllEnemies

Call resetBushes("All")

Call positionObj("RaccoonD", "DG183", "B")
Call EnemyTrigger("XXXXXXETRC01DDH183")

End Sub

Sub FC()
Call ResetAllEnemies


End Sub

Sub FD()
Call ResetAllEnemies


End Sub

Sub GA()
Call ResetAllEnemies


End Sub

Sub GB()
Call ResetAllEnemies


End Sub

Sub GC()
Call ResetAllEnemies


End Sub

Sub GD()
Call ResetAllEnemies


End Sub

Sub HA()
Call ResetAllEnemies


End Sub

Sub HB()
Call ResetAllEnemies


End Sub

Sub HC()
Call ResetAllEnemies


End Sub

Sub HD()
Call ResetAllEnemies


End Sub

Sub IA()
Call ResetAllEnemies


End Sub

Sub IB()
Call ResetAllEnemies


End Sub

Sub IC()
Call ResetAllEnemies

Call resetBushes("All")

Application.ScreenUpdating = False

Call positionObj("Bush1", "EE275", "B")
Call positionObj("Bush2", "EE279", "B")
Call positionObj("Bush3", "EE283", "B")
Call positionObj("Bush4", "EK283", "B")
Call positionObj("Bush5", "EQ287", "B")
Call positionObj("Bush6", "EQ291", "B")
Call positionObj("Bush7", "EW291", "B")
Call positionObj("Bush8", "FC291", "B")
Call positionObj("Bush9", "FI291", "B")

Application.ScreenUpdating = True

End Sub

Sub ID()
Call ResetAllEnemies


End Sub

Sub JA()
Call ResetAllEnemies

Call resetBushes("All")

Call positionObj("Bush1", "AS305", "B")

End Sub

Sub JB()
Call ResetAllEnemies


End Sub

Sub JC()
Call ResetAllEnemies


End Sub

Sub JD()
Call ResetAllEnemies

Call resetBushes("All")

Application.ScreenUpdating = False

Call positionObj("Bush1", "GS311", "B")
Call positionObj("Bush2", "GS315", "B")
Call positionObj("Bush3", "GS319", "B")
Call positionObj("Bush4", "GY307", "B")
Call positionObj("Bush5", "HE307", "B")
Call positionObj("Bush6", "HK307", "B")
Call positionObj("Bush7", "HQ311", "B")
Call positionObj("Bush8", "HQ315", "B")
Call positionObj("Bush9", "HQ319", "B")

Application.ScreenUpdating = True

End Sub

Sub KA()
Call ResetAllEnemies

Call resetBushes("All")

Call positionObj("Bush1", "AG339", "B")
Call positionObj("Bush2", "AM339", "B")
Call positionObj("Bush3", "AS339", "B")
Call positionObj("Bush4", "BE343", "B")

End Sub

Sub KB()
Call ResetAllEnemies


End Sub

Sub KC()
Call ResetAllEnemies

Call resetBushes("All")

End Sub

Sub KD()

'Bush screen

Call ResetAllEnemies

Call resetBushes("All")

Application.ScreenUpdating = False

Call positionObj("Bush1", "GS339", "B")
Call positionObj("Bush2", "GS343", "B")
Call positionObj("Bush3", "GS347", "B")
Call positionObj("Bush4", "GS351", "B")
Call positionObj("Bush5", "GS355", "B")

Call positionObj("Bush6", "GY339", "B")
Call positionObj("Bush7", "GY343", "B")
Call positionObj("Bush8", "GY347", "B")
Call positionObj("Bush9", "GY351", "B")
Call positionObj("Bush10", "GY355", "B")

Call positionObj("Bush11", "HE339", "B")
Call positionObj("Bush12", "HE343", "B")
Call positionObj("Bush13", "HE347", "B")
Call positionObj("Bush14", "HE351", "B")
Call positionObj("Bush15", "HE355", "B")

Call positionObj("Bush16", "HK339", "B")
Call positionObj("Bush17", "HK343", "B")
Call positionObj("Bush18", "HK347", "B")
Call positionObj("Bush19", "HK351", "B")
Call positionObj("Bush20", "HK355", "B")

Call positionObj("Bush21", "HQ339", "B")
Call positionObj("Bush22", "HQ343", "B")
Call positionObj("Bush23", "HQ347", "B")
Call positionObj("Bush24", "HQ351", "B")
Call positionObj("Bush25", "HQ355", "B")

Call positionObj("Bush26", "HW339", "B")
Call positionObj("Bush27", "HW343", "B")
Call positionObj("Bush28", "HW347", "B")
Call positionObj("Bush29", "HW351", "B")
Call positionObj("Bush30", "HW355", "B")

Application.ScreenUpdating = True
End Sub

Sub LA()
Call ResetAllEnemies
Call resetBushes("All")

Call positionObj("Bush1", "BE379", "B")
Call positionObj("Bush2", "BE383", "B")

End Sub

Sub LB()
Call ResetAllEnemies


End Sub

Sub LC()
Call ResetAllEnemies

Call resetBushes("All")


End Sub

Sub LD()

Call ResetAllEnemies

Call resetBushes("All")

Application.ScreenUpdating = False

Call positionObj("Bush1", "HE367", "B")
Call positionObj("Bush2", "HK367", "B")
Call positionObj("Bush3", "HQ367", "B")

Call positionObj("Bush4", "GY371", "B")
Call positionObj("Bush5", "GY375", "B")
Call positionObj("Bush6", "GY379", "B")
Call positionObj("Bush7", "GY383", "B")

Call positionObj("Bush8", "HW371", "B")
Call positionObj("Bush9", "HW375", "B")
Call positionObj("Bush10", "HW379", "B")
Call positionObj("Bush11", "HW383", "B")

Call positionObj("Bush12", "HE387", "B")
Call positionObj("Bush13", "HK387", "B")

Application.ScreenUpdating = True

End Sub

Sub MA()
Call ResetAllEnemies


End Sub

Sub MB()
Call ResetAllEnemies
Call resetBushes("All")

Call positionObj("Bush1", "DA403", "B")
Call positionObj("Bush2", "DG403", "B")
Call positionObj("Bush3", "CC411", "B")
Call positionObj("Bush4", "CI411", "B")
Call positionObj("Bush5", "CO411", "B")

Call EnemyTrigger("S1XXXXETOC01DCI415")
Call EnemyTrigger("S1XXXXETOC02DDA399")

End Sub

Sub MC()
Call ResetAllEnemies


End Sub

Sub MD()
Call ResetAllEnemies


End Sub

Sub NA()

Call ResetAllEnemies

Call EnemyTrigger("S1XXXXETOC01DV432")


End Sub

Sub NB()
Call ResetAllEnemies
Call resetBushes("All")

Call positionObj("Bush1", "CC451", "B")
Call positionObj("Bush2", "CI451", "B")
Call positionObj("Bush3", "DA447", "B")
Call positionObj("Bush4", "DG447", "B")
Call positionObj("Bush5", "DM447", "B")

End Sub

Sub NC()
Call ResetAllEnemies

End Sub

Sub ND()
Call ResetAllEnemies


End Sub

Sub OA()

Call ResetAllEnemies

Call resetBushes("All")

Call positionObj("Bush1", "AM483", "B")
Call positionObj("Bush2", "AY483", "B")
Call positionObj("Bush3", "BE483", "B")


Call EnemyTrigger("S1XXXXETOC01DAG463")
Call EnemyTrigger("S1XXXXETOC02DR481")

End Sub

Sub OB()
Call ResetAllEnemies

Call resetBushes("All")
Call positionObj("Bush1", "BW483", "B")
Call positionObj("Bush2", "CC483", "B")
Call positionObj("Bush3", "CU463", "B")
Call positionObj("Bush4", "DA463", "B")
Call positionObj("Bush5", "DG463", "B")

Call EnemyTrigger("S1XXXXETGD01DDA475")
Call EnemyTrigger("S1XXXXETGD02DDG475")
Call EnemyTrigger("S1XXXXETGD03Dcc467")

End Sub

Sub OC()
Call ResetAllEnemies


End Sub

Sub OD()
Call ResetAllEnemies


End Sub

Sub PA()
Call ResetAllEnemies

Call EnemyTrigger("S1XXXXETSC01DAK499")
Call EnemyTrigger("S1XXXXETSC02DBD507")


End Sub

Sub PB()
Call ResetAllEnemies

Call EnemyTrigger("S1XXXXETSC01DDA493")
Call EnemyTrigger("S1XXXXETSC02DCN506")

End Sub

Sub PC()
Call ResetAllEnemies

Call resetBushes("All")

Call EnemyTrigger("XXXXXXETGD01DEJ503")

If Sheets("Data").Range("Z4").Value = "" Then
    Call positionObj("SwordUp", "FA509", "")
        ActiveSheet.Range("EW507:EW514").Value = "XXXXXXSE0001XX"
        ActiveSheet.Range("EW507:FG507").Value = "XXXXXXSE0001XX"
        ActiveSheet.Range("FG507:FG514").Value = "XXXXXXSE0001XX"
End If


End Sub

Sub PD()
Call ResetAllEnemies


End Sub

Sub SA()
Call ResetAllEnemies


End Sub

Sub SB()
Call ResetAllEnemies

Call resetBushes("All")

If Sheets("Data").Range("Z3").Value <> "Y" Then
    ActiveSheet.Range(ActiveSheet.Range("DC595"), ActiveSheet.Range("DC595").Offset(11, 9)).Value = "XXXXXXSE0004XX"
    ActiveSheet.Range(ActiveSheet.Range("CQ613"), ActiveSheet.Range("CQ613").Offset(1, 7)).Value = "XXXXXXSE0003XX"
End If

Call positionObj("TarinD", "DG599", "B")
Call EnemyTrigger("XXXXXXETTA02DDG599")

Call positionObj("MarinD", "CU595", "B")
Call EnemyTrigger("XXXXXXETMA01DCU595")


End Sub

Sub SC()
Call ResetAllEnemies


End Sub

Sub SD()
Call ResetAllEnemies


End Sub

Sub TA()
Call ResetAllEnemies

Call resetBushes("All")

If Sheets("Data").Range("AB2").Value = "" Then
    Call positionObj("HeartPiece", "AG639", "XXXXXXSE0002XX")
End If

End Sub

Sub TB()
Call ResetAllEnemies


End Sub

Sub TC()
Call ResetAllEnemies

End Sub

Sub TD()
Call ResetAllEnemies


End Sub

Sub UA()
Call ResetAllEnemies


End Sub

Sub UB()
Call ResetAllEnemies


End Sub

Sub UC()
Call ResetAllEnemies


End Sub

Sub UD()
Call ResetAllEnemies


End Sub

Sub VA()
Call ResetAllEnemies


End Sub

Sub VB()
Call ResetAllEnemies


End Sub

Sub VC()
Call ResetAllEnemies


End Sub

Sub VD()
Call ResetAllEnemies


End Sub
