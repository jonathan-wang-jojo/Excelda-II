'Attribute VB_Name = "SceneGame1"
Option Explicit

Implements IGameScene

Private Const SCENE_ID As String = "Game1"

Private Property Get IGameScene_SceneId() As String
    IGameScene_SceneId = SCENE_ID
End Property

Private Property Get IGameScene_SheetName() As String
    IGameScene_SheetName = SHEET_GAME
End Property

Private Function IGameScene_CanHandleScreen(ByVal screenCode As String) As Boolean
    Dim code As String
    code = NormalizeScreenCode(screenCode)
    If code = "" Then Exit Function

    If Len(code) <> 2 Then Exit Function

    Dim rowCode As String
    Dim colCode As String
    rowCode = Mid(code, 1, 1)
    colCode = Mid(code, 2, 1)

    If rowCode >= "A" And rowCode <= "V" Then
        If colCode >= "A" And colCode <= "D" Then
            IGameScene_CanHandleScreen = True
        End If
    End If
End Function

Private Sub IGameScene_ApplyScreenLayout(ByVal screenCode As String)
    Dim code As String
    code = NormalizeScreenCode(screenCode)
    If code = "" Then Exit Sub

    ResetAllEnemies

    Select Case code
        Case "AA": SetupAA
        Case "AB": SetupAB
        Case "AC": SetupAC
        Case "AD": SetupAD
        Case "BA": SetupBA
        Case "BB": SetupBB
        Case "BC": SetupBC
        Case "BD": SetupBD
        Case "CA": SetupCA
        Case "CB": SetupCB
        Case "CC": SetupCC
        Case "CD": SetupCD
        Case "DA": SetupDA
        Case "DB": SetupDB
        Case "DC": SetupDC
        Case "DD": SetupDD
        Case "EA": SetupEA
        Case "EB": SetupEB
        Case "EC": SetupEC
        Case "ED": SetupED
        Case "FA": SetupFA
        Case "FB": SetupFB
        Case "FC": SetupFC
        Case "FD": SetupFD
        Case "GA": SetupGA
        Case "GB": SetupGB
        Case "GC": SetupGC
        Case "GD": SetupGD
        Case "HA": SetupHA
        Case "HB": SetupHB
        Case "HC": SetupHC
        Case "HD": SetupHD
        Case "IA": SetupIA
        Case "IB": SetupIB
        Case "IC": SetupIC
        Case "ID": SetupID
        Case "JA": SetupJA
        Case "JB": SetupJB
        Case "JC": SetupJC
        Case "JD": SetupJD
        Case "KA": SetupKA
        Case "KB": SetupKB
        Case "KC": SetupKC
        Case "KD": SetupKD
        Case "LA": SetupLA
        Case "LB": SetupLB
        Case "LC": SetupLC
        Case "LD": SetupLD
        Case "MA": SetupMA
        Case "MB": SetupMB
        Case "MC": SetupMC
        Case "MD": SetupMD
        Case "NA": SetupNA
        Case "NB": SetupNB
        Case "NC": SetupNC
        Case "ND": SetupND
        Case "OA": SetupOA
        Case "OB": SetupOB
        Case "OC": SetupOC
        Case "OD": SetupOD
        Case "PA": SetupPA
        Case "PB": SetupPB
        Case "PC": SetupPC
        Case "PD": SetupPD
        Case "SA": SetupSA
        Case "SB": SetupSB
        Case "SC": SetupSC
        Case "SD": SetupSD
        Case "TA": SetupTA
        Case "TB": SetupTB
        Case "TC": SetupTC
        Case Else
            Debug.Print "SceneGame1: Unhandled screen code: " & code
    End Select
End Sub

Private Sub IGameScene_OnEnterScene()
    ' No scene-level setup required yet.
End Sub

Private Sub IGameScene_OnExitScene()
    ' No scene-level teardown required yet.
End Sub

Private Function NormalizeScreenCode(ByVal screenCode As String) As String
    NormalizeScreenCode = UCase(Trim(screenCode))
End Function

Private Function GetObjectManager() As ObjectManager
    Set GetObjectManager = ObjectManagerInstance()
End Function

Private Sub ResetBushes(ByVal whichBush As String)
    Dim manager As ObjectManager
    Set manager = GetObjectManager()
    If manager Is Nothing Then Exit Sub
    manager.ResetObjects whichBush
End Sub

Private Function SuspendScreenUpdating() As Boolean
    SuspendScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
End Function

Private Sub RestoreScreenUpdating(ByVal previousState As Boolean)
    Application.ScreenUpdating = previousState
End Sub

Private Sub PositionObj(ByVal objectName As String, ByVal objectLocation As String, Optional ByVal cellValue As Variant)
    Dim manager As ObjectManager
    Set manager = GetObjectManager()
    If manager Is Nothing Then Exit Sub
    If IsMissing(cellValue) Then
        manager.PositionObject objectName, objectLocation
    Else
        manager.PositionObject objectName, objectLocation, cellValue
    End If
End Sub

Private Sub PositionMultiple(ByVal objectType As String, _
                            ByVal Obj1 As Variant, ByVal Obj2 As Variant, ByVal Obj3 As Variant, ByVal Obj4 As Variant, _
                            ByVal Obj5 As Variant, ByVal Obj6 As Variant, ByVal Obj7 As Variant, ByVal Obj8 As Variant, _
                            ByVal Obj9 As Variant, ByVal Obj10 As Variant, ByVal Obj11 As Variant, ByVal Obj12 As Variant, _
                            ByVal Obj13 As Variant, ByVal Obj14 As Variant, ByVal Obj15 As Variant, ByVal Obj16 As Variant, _
                            ByVal Obj17 As Variant, ByVal Obj18 As Variant, ByVal Obj19 As Variant, ByVal Obj20 As Variant, _
                            ByVal Obj21 As Variant, ByVal Obj22 As Variant, ByVal Obj23 As Variant, ByVal Obj24 As Variant, _
                            ByVal Obj25 As Variant, ByVal Obj26 As Variant, ByVal Obj27 As Variant, ByVal Obj28 As Variant, _
                            ByVal Obj29 As Variant, ByVal Obj30 As Variant)

    Dim manager As ObjectManager
    Set manager = GetObjectManager()
    If manager Is Nothing Then Exit Sub

    manager.PositionMultiple objectType, _
        Obj1, Obj2, Obj3, Obj4, Obj5, Obj6, Obj7, Obj8, Obj9, Obj10, _
        Obj11, Obj12, Obj13, Obj14, Obj15, Obj16, Obj17, Obj18, Obj19, Obj20, _
        Obj21, Obj22, Obj23, Obj24, Obj25, Obj26, Obj27, Obj28, Obj29, Obj30
End Sub

Private Sub GrantHeartPiece(ByVal heartNum As Variant)
    Dim manager As ObjectManager
    Set manager = GetObjectManager()
    If manager Is Nothing Then Exit Sub
    manager.CollectHeartPiece CLng(Val(heartNum))
End Sub

'===================================================================================
'                              SCREEN SETUP ROUTINES
'===================================================================================

Private Sub SetupAA()
    ResetBushes "All"
End Sub

Private Sub SetupAB()
    ResetBushes "All"
End Sub

Private Sub SetupAC()
    ResetBushes "All"
End Sub

Private Sub SetupAD()
    ' Empty screen
End Sub

Private Sub SetupBA()
    ' Empty screen
End Sub

Private Sub SetupBB()
    ' Empty screen
End Sub

Private Sub SetupBC()
    ' Empty screen
End Sub

Private Sub SetupBD()
    ' Empty screen
End Sub

Private Sub SetupCA()
    ' Empty screen
End Sub

Private Sub SetupCB()
    ' Empty screen
End Sub

Private Sub SetupCC()
    ' Empty screen
End Sub

Private Sub SetupCD()
    ' Empty screen
End Sub

Private Sub SetupDA()
    ' Empty screen
End Sub

Private Sub SetupDB()
    ' Empty screen
End Sub

Private Sub SetupDC()
    ' Empty screen
End Sub

Private Sub SetupDD()
    ' Empty screen
End Sub

Private Sub SetupEA()
    ' Empty screen
End Sub

Private Sub SetupEB()
    ' Empty screen
End Sub

Private Sub SetupEC()
    ' Empty screen
End Sub

Private Sub SetupED()
    ' Empty screen
End Sub

Private Sub SetupFA()
    ' Empty screen
End Sub

Private Sub SetupFB()
    ResetBushes "All"
    PositionObj "RaccoonD", "DG183", "B"
    EnemyTrigger "XXXXXXETRC01DDH183"
End Sub

Private Sub SetupFC()
    ' Empty screen
End Sub

Private Sub SetupFD()
    ' Empty screen
End Sub

Private Sub SetupGA()
    ' Empty screen
End Sub

Private Sub SetupGB()
    ' Empty screen
End Sub

Private Sub SetupGC()
    ' Empty screen
End Sub

Private Sub SetupGD()
    ' Empty screen
End Sub

Private Sub SetupHA()
    ' Empty screen
End Sub

Private Sub SetupHB()
    ' Empty screen
End Sub

Private Sub SetupHC()
    ' Empty screen
End Sub

Private Sub SetupHD()
    ' Empty screen
End Sub

Private Sub SetupIA()
    ' Empty screen
End Sub

Private Sub SetupIB()
    ' Empty screen
End Sub

Private Sub SetupIC()
    Dim wasUpdating As Boolean
    wasUpdating = SuspendScreenUpdating()
    ResetBushes "All"
    PositionObj "Bush1", "EE275", "B"
    PositionObj "Bush2", "EE279", "B"
    PositionObj "Bush3", "EE283", "B"
    PositionObj "Bush4", "EK283", "B"
    PositionObj "Bush5", "EQ287", "B"
    PositionObj "Bush6", "EQ291", "B"
    PositionObj "Bush7", "EW291", "B"
    PositionObj "Bush8", "FC291", "B"
    PositionObj "Bush9", "FI291", "B"
    RestoreScreenUpdating wasUpdating
End Sub

Private Sub SetupID()
    ' Empty screen
End Sub

Private Sub SetupJA()
    ResetBushes "All"
    PositionObj "Bush1", "AS305", "B"
End Sub

Private Sub SetupJB()
    ' Empty screen
End Sub

Private Sub SetupJC()
    ' Empty screen
End Sub

Private Sub SetupJD()
    Dim wasUpdating As Boolean
    wasUpdating = SuspendScreenUpdating()
    ResetBushes "All"
    PositionObj "Bush1", "GS311", "B"
    PositionObj "Bush2", "GS315", "B"
    PositionObj "Bush3", "GS319", "B"
    PositionObj "Bush4", "GY307", "B"
    PositionObj "Bush5", "HE307", "B"
    PositionObj "Bush6", "HK307", "B"
    PositionObj "Bush7", "HQ311", "B"
    PositionObj "Bush8", "HQ315", "B"
    PositionObj "Bush9", "HQ319", "B"
    RestoreScreenUpdating wasUpdating
End Sub

Private Sub SetupKA()
    ResetBushes "All"
    PositionObj "Bush1", "AG339", "B"
    PositionObj "Bush2", "AM339", "B"
    PositionObj "Bush3", "AS339", "B"
    PositionObj "Bush4", "BE343", "B"
End Sub

Private Sub SetupKB()
    ' Empty screen
End Sub

Private Sub SetupKC()
    ResetBushes "All"
End Sub

Private Sub SetupKD()
    Dim wasUpdating As Boolean
    wasUpdating = SuspendScreenUpdating()
    ResetBushes "All"
    PositionObj "Bush1", "GS339", "B"
    PositionObj "Bush2", "GS343", "B"
    PositionObj "Bush3", "GS347", "B"
    PositionObj "Bush4", "GS351", "B"
    PositionObj "Bush5", "GS355", "B"
    PositionObj "Bush6", "GY339", "B"
    PositionObj "Bush7", "GY343", "B"
    PositionObj "Bush8", "GY347", "B"
    PositionObj "Bush9", "GY351", "B"
    PositionObj "Bush10", "GY355", "B"
    PositionObj "Bush11", "HE339", "B"
    PositionObj "Bush12", "HE343", "B"
    PositionObj "Bush13", "HE347", "B"
    PositionObj "Bush14", "HE351", "B"
    PositionObj "Bush15", "HE355", "B"
    PositionObj "Bush16", "HK339", "B"
    PositionObj "Bush17", "HK343", "B"
    PositionObj "Bush18", "HK347", "B"
    PositionObj "Bush19", "HK351", "B"
    PositionObj "Bush20", "HK355", "B"
    PositionObj "Bush21", "HQ339", "B"
    PositionObj "Bush22", "HQ343", "B"
    PositionObj "Bush23", "HQ347", "B"
    PositionObj "Bush24", "HQ351", "B"
    PositionObj "Bush25", "HQ355", "B"
    PositionObj "Bush26", "HW339", "B"
    PositionObj "Bush27", "HW343", "B"
    PositionObj "Bush28", "HW347", "B"
    PositionObj "Bush29", "HW351", "B"
    PositionObj "Bush30", "HW355", "B"
    RestoreScreenUpdating wasUpdating
End Sub

Private Sub SetupLA()
    ResetBushes "All"
    PositionObj "Bush1", "BE379", "B"
    PositionObj "Bush2", "BE383", "B"
End Sub

Private Sub SetupLB()
    ' Empty screen
End Sub

Private Sub SetupLC()
    ResetBushes "All"
End Sub

Private Sub SetupLD()
    Dim wasUpdating As Boolean
    wasUpdating = SuspendScreenUpdating()
    ResetBushes "All"
    PositionObj "Bush1", "HE367", "B"
    PositionObj "Bush2", "HK367", "B"
    PositionObj "Bush3", "HQ367", "B"
    PositionObj "Bush4", "GY371", "B"
    PositionObj "Bush5", "GY375", "B"
    PositionObj "Bush6", "GY379", "B"
    PositionObj "Bush7", "GY383", "B"
    PositionObj "Bush8", "HW371", "B"
    PositionObj "Bush9", "HW375", "B"
    PositionObj "Bush10", "HW379", "B"
    PositionObj "Bush11", "HW383", "B"
    PositionObj "Bush12", "HE387", "B"
    PositionObj "Bush13", "HK387", "B"
    RestoreScreenUpdating wasUpdating
End Sub

Private Sub SetupMA()
    ' Empty screen
End Sub

Private Sub SetupMB()
    ResetBushes "All"
    PositionObj "Bush1", "DA403", "B"
    PositionObj "Bush2", "DG403", "B"
    PositionObj "Bush3", "CC411", "B"
    PositionObj "Bush4", "CI411", "B"
    PositionObj "Bush5", "CO411", "B"
    EnemyTrigger "S1XXXXETOC01DCI415"
    EnemyTrigger "S1XXXXETOC02DDA399"
End Sub

Private Sub SetupMC()
    ' Empty screen
End Sub

Private Sub SetupMD()
    ' Empty screen
End Sub

Private Sub SetupNA()
    EnemyTrigger "S1XXXXETOC01DV432"
End Sub

Private Sub SetupNB()
    ResetBushes "All"
    PositionObj "Bush1", "CC451", "B"
    PositionObj "Bush2", "CI451", "B"
    PositionObj "Bush3", "DA447", "B"
    PositionObj "Bush4", "DG447", "B"
    PositionObj "Bush5", "DM447", "B"
End Sub

Private Sub SetupNC()
    ' Empty screen
End Sub

Private Sub SetupND()
    ' Empty screen
End Sub

Private Sub SetupOA()
    ResetBushes "All"
    PositionObj "Bush1", "AM483", "B"
    PositionObj "Bush2", "AY483", "B"
    PositionObj "Bush3", "BE483", "B"
    EnemyTrigger "S1XXXXETOC01DAG463"
    EnemyTrigger "S1XXXXETOC02DR481"
End Sub

Private Sub SetupOB()
    ResetBushes "All"
    PositionObj "Bush1", "BW483", "B"
    PositionObj "Bush2", "CC483", "B"
    PositionObj "Bush3", "CU463", "B"
    PositionObj "Bush4", "DA463", "B"
    PositionObj "Bush5", "DG463", "B"
    EnemyTrigger "S1XXXXETGD01DDA475"
    EnemyTrigger "S1XXXXETGD02DDG475"
    EnemyTrigger "S1XXXXETGD03Dcc467"
End Sub

Private Sub SetupOC()
    ' Empty screen
End Sub

Private Sub SetupOD()
    ' Empty screen
End Sub

Private Sub SetupPA()
    EnemyTrigger "S1XXXXETSC01DAK499"
    EnemyTrigger "S1XXXXETSC02DBD507"
End Sub

Private Sub SetupPB()
    EnemyTrigger "S1XXXXETSC01DDA493"
    EnemyTrigger "S1XXXXETSC02DCN506"
End Sub

Private Sub SetupPC()
    ResetBushes "All"
    EnemyTrigger "XXXXXXETGD01DEJ503"
    If Sheets(SHEET_DATA).Range("Z4").Value = "" Then
        ActiveSheet.Range(ActiveSheet.Range("EW507"), ActiveSheet.Range("EW507").Offset(6, 10)).Value = "XXXXXXSE0001XX"
    End If
End Sub

Private Sub SetupPD()
    ' Empty screen
End Sub

Private Sub SetupSA()
    ' Empty screen
End Sub

Private Sub SetupSB()
    ResetBushes "All"
    If Sheets(SHEET_DATA).Range("Z3").Value <> "Y" Then
        ActiveSheet.Range(ActiveSheet.Range("CQ613"), ActiveSheet.Range("CQ613").Offset(1, 7)).Value = "XXXXXXSE0003XX"
    End If
    PositionObj "TarinD", "DG599", "B"
    EnemyTrigger "XXXXXXETTA02DDG599"
    PositionObj "MarinD", "CU595", "B"
    EnemyTrigger "XXXXXXETMA01DCU595"
End Sub

Private Sub SetupSC()
    ' Empty screen
End Sub

Private Sub SetupSD()
    ' Empty screen
End Sub

Private Sub SetupTA()
    ResetBushes "All"
    If Sheets(SHEET_DATA).Range("AB2").Value = "" Then
        PositionObj "HeartPiece", "AG639", "XXXXXXSE0002XX"
    End If
End Sub

Private Sub SetupTB()
    ' Empty screen
End Sub

Private Sub SetupTC()
    ' Empty screen
End Sub