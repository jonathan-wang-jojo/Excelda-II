'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "SpecialEventManager"
Option Explicit

Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Integer) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'###################################################################################
'#                              SPECIAL EVENT FLOW                                #
'###################################################################################

Private Sub Class_Initialize()
    Initialize
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

Public Sub Initialize()
    ' No persistent state to seed yet; defined for lifecycle consistency.
End Sub

Public Sub Reset()
    ' No persistent state to reset.
End Sub

Public Sub Destroy()
    ' No persistent state to release.
End Sub

Public Sub Trigger(ByVal eventCode As String)
    Dim eventId As String
    eventId = ExtractEventId(eventCode)
    If eventId = "" Then Exit Sub

    Select Case eventId
        Case "0001": HandleEvent0001
        Case "0002": HandleEvent0002
        Case "0003": HandleEvent0003
        Case "0004": HandleEvent0004
        Case Else
            Debug.Print "Unhandled special event: " & eventId
    End Select
End Sub

'-----------------------------------------------------------------------------------
'                               Event Handlers
'-----------------------------------------------------------------------------------

Private Sub HandleEvent0001()
    Dim ws As Worksheet
    Set ws = CurrentScreenSheet()
    If ws Is Nothing Then Exit Sub

    Dim owl1 As Shape, owl2 As Shape
    Set owl1 = ShapeIfExists(ws, "Owl1")
    Set owl2 = ShapeIfExists(ws, "Owl2")
    If owl1 Is Nothing Or owl2 Is Nothing Then Exit Sub

    Dim startCell As Range
    On Error Resume Next
    Set startCell = ws.Range("DU487")
    On Error GoTo 0
    If startCell Is Nothing Then Exit Sub

    SetShapePosition owl1, startCell
    SetShapePosition owl2, startCell
    owl1.Visible = True
    owl2.Visible = False

    AnimateOwls owl1, owl2, 30, 3, 7

    ShowDialogue DataSheet().Range("T5").Value

    AnimateOwls owl1, owl2, 30, -3, -7

    owl1.Visible = False
    owl2.Visible = False

    GiveLinkSword ws
End Sub

Private Sub HandleEvent0002()
    Dim objManager As ObjectManager
    Set objManager = ObjectManagerInstance()
    If objManager Is Nothing Then Exit Sub
    objManager.CollectHeartPiece 2
End Sub

Private Sub HandleEvent0003()
    ShowDialogue DataSheet().Range("T9").Value

    Dim spriteMgr As SpriteManager
    Set spriteMgr = SpriteManagerInstance()
    If spriteMgr Is Nothing Then Exit Sub

    Dim linkShape As Shape
    Set linkShape = spriteMgr.LinkSprite
    If linkShape Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = CurrentScreenSheet()
    If ws Is Nothing Then Exit Sub

    Dim newTop As Double
    newTop = linkShape.Top - 40

    spriteMgr.AlignSprites linkShape.Left, newTop
    spriteMgr.UpdatePosition
    spriteMgr.UpdateVisibility

    Dim gs As GameState
    Set gs = GameStateInstance()
    If Not gs Is Nothing Then
        gs.LinkCellAddress = linkShape.TopLeftCell.Address
        gs.CodeCell = ""
    End If

    Sheets(SHEET_DATA).Range("C18").Value = linkShape.TopLeftCell.Address
End Sub

Private Sub HandleEvent0004()
    If GetAsyncKeyState(KEY_C) = 0 And GetAsyncKeyState(KEY_D) = 0 Then Exit Sub

    ShowDialogue DataSheet().Range("T10").Value

    Dim ws As Worksheet
    Set ws = CurrentScreenSheet()
    If ws Is Nothing Then Exit Sub

    Dim spriteMgr As SpriteManager
    Set spriteMgr = SpriteManagerInstance()
    If spriteMgr Is Nothing Then Exit Sub

    Dim linkSprite As Shape
    Set linkSprite = spriteMgr.LinkSprite
    If linkSprite Is Nothing Then Exit Sub

    Dim linkWin As Shape
    Set linkWin = ShapeIfExists(ws, "LinkWin")
    Dim shieldShape As Shape
    Set shieldShape = ShapeIfExists(ws, "LinkShieldDown")
    If linkWin Is Nothing Or shieldShape Is Nothing Then Exit Sub

    linkWin.Top = linkSprite.Top
    linkWin.Left = linkSprite.Left
    linkSprite.Visible = False
    linkWin.Visible = True

    shieldShape.Top = linkWin.Top - 45
    shieldShape.Left = linkWin.Left
    shieldShape.Visible = True

    TickFrame ws, 2000

    ShowDialogue DataSheet().Range("T11").Value

    linkWin.Visible = False
    shieldShape.Visible = False
    linkSprite.Visible = True

    ShowDialogue DataSheet().Range("T12").Value

    Dim dataSheet As Worksheet
    Set dataSheet = DataSheet()
    dataSheet.Range("Z3").Value = "Y"
    dataSheet.Range("C26").Value = "Shield"

    ActionManagerInstance.Reset

    ws.Range(ws.Range("DC595"), ws.Range("DC595").Offset(11, 9)).Value = ""
    ws.Range(ws.Range("CQ613"), ws.Range("CQ613").Offset(1, 7)).Value = ""
End Sub

'-----------------------------------------------------------------------------------
'                               Helper Methods
'-----------------------------------------------------------------------------------

Private Sub GiveLinkSword(ByVal ws As Worksheet)
    Dim spriteMgr As SpriteManager
    Set spriteMgr = SpriteManagerInstance()
    If spriteMgr Is Nothing Then Exit Sub

    Dim linkSprite As Shape
    Set linkSprite = spriteMgr.LinkSprite
    If linkSprite Is Nothing Then Exit Sub

    Dim linkWin As Shape
    Set linkWin = ShapeIfExists(ws, "LinkWin")
    Dim swordUp As Shape
    Set swordUp = ShapeIfExists(ws, "SwordUp")
    If linkWin Is Nothing Or swordUp Is Nothing Then Exit Sub

    linkWin.Top = linkSprite.Top
    linkWin.Left = linkSprite.Left
    linkSprite.Visible = False
    linkWin.Visible = True

    swordUp.Top = linkWin.Top - 45
    swordUp.Left = linkWin.Left + 20
    swordUp.Visible = True

    TickFrame ws, 2000

    swordUp.Visible = False
    Call swordSpin

    linkSprite.Visible = True
    linkWin.Visible = False

    ShowDialogue DataSheet().Range("T6").Value

    ws.Range("EW507:EW514").Value = ""
    ws.Range("EW507:FG507").Value = ""
    ws.Range("FG507:FG514").Value = ""

    Dim dataSheet As Worksheet
    Set dataSheet = DataSheet()
    dataSheet.Range("Z4").Value = "Y"
    dataSheet.Range("C27").Value = "Sword"

    ActionManagerInstance.Reset
End Sub

Private Sub AnimateOwls(ByVal owl1 As Shape, ByVal owl2 As Shape, ByVal steps As Long, ByVal deltaTop As Double, ByVal deltaLeft As Double)
    Dim ws As Worksheet
    Set ws = owl1.TopLeftCell.Worksheet

    Dim frameCounter As Long
    Dim i As Long
    For i = 1 To steps
        frameCounter = frameCounter + 1
        UpdateOwlFrames owl1, owl2, frameCounter

        owl1.Top = owl1.Top + deltaTop
        owl1.Left = owl1.Left + deltaLeft
        owl2.Top = owl1.Top
        owl2.Left = owl1.Left

        TickFrame ws, 25
    Next i
End Sub

Private Sub UpdateOwlFrames(ByVal owl1 As Shape, ByVal owl2 As Shape, ByVal frameCounter As Long)
    Dim phase As Long
    phase = frameCounter Mod 6

    If phase < 3 Then
        owl1.Visible = True
        owl2.Visible = False
    Else
        owl1.Visible = False
        owl2.Visible = True
    End If
End Sub

Private Function ExtractEventId(ByVal eventCode As String) As String
    If Len(eventCode) >= 12 Then
        ExtractEventId = Mid$(eventCode, 9, 4)
    Else
        ExtractEventId = ""
    End If
End Function

Private Function CurrentScreenSheet() As Worksheet
    On Error Resume Next
    Dim gs As GameState
    Set gs = GameStateInstance()
    If gs Is Nothing Or gs.CurrentScreen = "" Then Exit Function
    Set CurrentScreenSheet = Sheets(gs.CurrentScreen)
    On Error GoTo 0
End Function

Private Function ShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    On Error Resume Next
    Set ShapeIfExists = ws.Shapes(shapeName)
    If Err.Number <> 0 Then
        Err.Clear
        Set ShapeIfExists = Nothing
    End If
    On Error GoTo 0
End Function

Private Sub SetShapePosition(ByVal shp As Shape, ByVal anchorCell As Range)
    shp.Top = anchorCell.Top
    shp.Left = anchorCell.Left
End Sub

Private Sub TickFrame(ByVal ws As Worksheet, ByVal delayMs As Long)
    On Error Resume Next
    ws.Range("A1").Copy ws.Range("A2")
    On Error GoTo 0
    If delayMs > 0 Then Sleep delayMs
End Sub

Private Sub ShowDialogue(ByVal textValue As Variant)
    Dim dataSheet As Worksheet
    Set dataSheet = DataSheet()
    If dataSheet Is Nothing Then Exit Sub
    dataSheet.Range("C42").Value = textValue
    DialogueForm.Show
End Sub

Private Function DataSheet() As Worksheet
    On Error Resume Next
    Set DataSheet = Sheets(SHEET_DATA)
    On Error GoTo 0
End Function