'Attribute VB_Name = "AG_LinkActions"
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'####################################################################################
'#    Link context helpers
'####################################################################################
Private Property Get LinkSprite() As Shape
    Dim manager As SpriteManager
    Set manager = SpriteManagerInstance()
    If manager Is Nothing Then Exit Property
    Set LinkSprite = manager.LinkSprite
End Property

Private Sub SyncLinkState(ByVal link As Shape, ByVal manager As SpriteManager, ByVal gs As GameState)
    If link Is Nothing Then Exit Sub
    If manager Is Nothing Then Exit Sub
    If gs Is Nothing Then Exit Sub

    manager.LinkSpriteTop = link.Top
    manager.LinkSpriteLeft = link.Left

    On Error Resume Next
    gs.LinkCellAddress = link.TopLeftCell.Address
    Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value = gs.LinkCellAddress
    On Error GoTo 0
End Sub

Private Function EnsureLinkContext(ByRef gs As GameState, ByRef manager As SpriteManager, _
                                   ByRef link As Shape, ByRef linkSheet As Worksheet) As Boolean
    Set gs = GameStateInstance()
    Set manager = SpriteManagerInstance()
    Set link = LinkSprite

    If gs Is Nothing Then Exit Function
    If manager Is Nothing Then Exit Function
    If link Is Nothing Then Exit Function

    Set linkSheet = link.Parent
    EnsureLinkContext = True
End Function

Private Function NormalizeDirection(ByVal currentDir As String, ByVal fallbackDir As String) As String
    Dim resolved As String
    resolved = Trim$(currentDir)
    If resolved = "" Then resolved = Trim$(fallbackDir)
    If resolved = "" Then resolved = "D"
    NormalizeDirection = UCase$(resolved)
End Function

Private Function TryGetShape(ByVal hostSheet As Worksheet, ByVal shapeName As String) As Shape
    If hostSheet Is Nothing Then Exit Function
    On Error Resume Next
    Set TryGetShape = hostSheet.Shapes(shapeName)
    On Error GoTo 0
End Function

Private Sub PositionShape(ByVal sprite As Shape, ByVal topOffset As Double, ByVal leftOffset As Double)
    If sprite Is Nothing Then Exit Sub
    sprite.Top = topOffset
    sprite.Left = leftOffset
End Sub

Private Function LoadSwordSprites(ByVal linkSheet As Worksheet, _
                                  ByRef swordUp As Shape, _
                                  ByRef swordDown As Shape, _
                                  ByRef swordLeft As Shape, _
                                  ByRef swordRight As Shape, _
                                  ByRef swipeUpLeft As Shape, _
                                  ByRef swipeUpRight As Shape, _
                                  ByRef swipeDownLeft As Shape, _
                                  ByRef swipeDownRight As Shape) As Boolean
    Set swordUp = TryGetShape(linkSheet, "SwordUp")
    Set swordDown = TryGetShape(linkSheet, "SwordDown")
    Set swordLeft = TryGetShape(linkSheet, "SwordLeft")
    Set swordRight = TryGetShape(linkSheet, "SwordRight")

    Set swipeUpLeft = TryGetShape(linkSheet, "SwordSwipeUpLeft")
    Set swipeUpRight = TryGetShape(linkSheet, "SwordSwipeUpRight")
    Set swipeDownLeft = TryGetShape(linkSheet, "SwordSwipeDownLeft")
    Set swipeDownRight = TryGetShape(linkSheet, "SwordSwipeDownRight")

    LoadSwordSprites = Not (swordUp Is Nothing Or swordDown Is Nothing Or swordLeft Is Nothing Or swordRight Is Nothing)
End Function

Private Function LoadLinkFacingSprites(ByVal linkSheet As Worksheet, _
                                       ByRef linkLeft1 As Shape, _
                                       ByRef linkRight1 As Shape, _
                                       ByRef linkUp1 As Shape, _
                                       ByRef linkDown1 As Shape) As Boolean
    Set linkLeft1 = TryGetShape(linkSheet, "LinkLeft1")
    Set linkRight1 = TryGetShape(linkSheet, "LinkRight1")
    Set linkUp1 = TryGetShape(linkSheet, "LinkUp1")
    Set linkDown1 = TryGetShape(linkSheet, "LinkDown1")

    LoadLinkFacingSprites = Not (linkLeft1 Is Nothing Or linkRight1 Is Nothing Or linkUp1 Is Nothing Or linkDown1 Is Nothing)
End Function

Private Function LoadShieldSprites(ByVal linkSheet As Worksheet, _
                                   ByRef shieldUp As Shape, _
                                   ByRef shieldDown As Shape, _
                                   ByRef shieldLeft As Shape, _
                                   ByRef shieldRight As Shape) As Boolean
    Set shieldUp = TryGetShape(linkSheet, "LinkShieldUp")
    Set shieldDown = TryGetShape(linkSheet, "LinkShieldDown")
    Set shieldLeft = TryGetShape(linkSheet, "LinkShieldLeft")
    Set shieldRight = TryGetShape(linkSheet, "LinkShieldRight")

    LoadShieldSprites = Not (shieldUp Is Nothing Or shieldDown Is Nothing Or shieldLeft Is Nothing Or shieldRight Is Nothing)
End Function

'####################################################################################
'#    Animation utilities
'####################################################################################
Private Sub TriggerFrameTick(ByVal hostSheet As Worksheet, ByVal delayMs As Long)
    If hostSheet Is Nothing Then Exit Sub
    hostSheet.Range("A1").Copy hostSheet.Range("A2")
    Sleep delayMs
End Sub

Private Sub HandleScrollTrigger(ByVal tileCode As String)
    Dim scrollCode As String
    tileCode = Trim$(tileCode)
    If Len(tileCode) < 2 Then Exit Sub

    scrollCode = Left$(tileCode, 2)

    Select Case scrollCode
        Case "S1": myScroll 1
        Case "S2": myScroll 2
    End Select
End Sub

Private Sub AdvanceJumpFrame(ByVal frame As Shape, ByVal link As Shape, ByVal manager As SpriteManager, _
                             ByVal gs As GameState, ByVal linkSheet As Worksheet)
    Dim stepIndex As Long
    Dim tileCode As String

    If frame Is Nothing Then Exit Sub

    For stepIndex = 1 To 10
        frame.Top = frame.Top + 2
        link.Top = frame.Top
        SyncLinkState link, manager, gs

        tileCode = CStr(frame.TopLeftCell.Value)
        HandleScrollTrigger tileCode
        TriggerFrameTick linkSheet, 10
    Next stepIndex
End Sub

Private Sub PlayFallFrame(ByVal frame As Shape, ByVal linkSheet As Worksheet)
    Dim iteration As Long

    If frame Is Nothing Then Exit Sub

    frame.Visible = True
    For iteration = 1 To 30
        TriggerFrameTick linkSheet, 10
    Next iteration
    frame.Visible = False
End Sub

Private Sub HideShapes(ByVal shapeSet As Variant)
    Dim frame As Variant
    For Each frame In shapeSet
        If Not frame Is Nothing Then frame.Visible = False
    Next frame
End Sub

Private Sub RegisterSwordHits(ByVal swordFrame As Shape)
    If swordFrame Is Nothing Then Exit Sub

    Dim enemyManager As EnemyManager
    Set enemyManager = EnemyManagerInstance()
    If Not enemyManager Is Nothing Then
        enemyManager.HandleSwordHit swordFrame
    End If

    Dim objectManager As ObjectManager
    Set objectManager = ObjectManagerInstance()
    If Not objectManager Is Nothing Then
        objectManager.HandleSwordHit swordFrame
    End If
End Sub

'####################################################################################
'#    Falling / jumping
'####################################################################################
Public Sub Falling()
    Dim gs As GameState
    Dim manager As SpriteManager
    Dim link As Shape
    Dim linkSheet As Worksheet
    If Not EnsureLinkContext(gs, manager, link, linkSheet) Then Exit Sub

    Sheets(SHEET_DATA).Range(RANGE_FALL_SEQUENCE).Value = "Y"

    Dim targetCode As String
    targetCode = Mid$(gs.CodeCell, 5, 4)
    If targetCode = "XXXX" Then
        targetCode = Sheets(SHEET_DATA).Range(RANGE_CURRENT_CELL).Value
    End If

    Dim direction As String
    direction = gs.MoveDir
    If direction = "" Then direction = gs.LastDir

    Dim fallFrames(1 To 3) As Shape
    Set fallFrames(1) = TryGetShape(linkSheet, "LinkFall1")
    Set fallFrames(2) = TryGetShape(linkSheet, "LinkFall2")
    Set fallFrames(3) = TryGetShape(linkSheet, "LinkFall3")

    If fallFrames(1) Is Nothing Or fallFrames(2) Is Nothing Or fallFrames(3) Is Nothing Then
        link.Visible = True
        Sheets(SHEET_DATA).Range(RANGE_FALL_SEQUENCE).Value = "N"
        Exit Sub
    End If

    Dim baseTop As Double
    Dim baseLeft As Double
    baseTop = link.Top
    baseLeft = link.Left

    Select Case direction
        Case "U"
            baseTop = link.Top - 15
        Case "D"
            baseTop = link.Top + 50
        Case "L"
            baseLeft = link.Left - 20
        Case "R"
            baseLeft = link.Left + 20
    End Select

    Dim index As Long
    For index = LBound(fallFrames) To UBound(fallFrames)
        PositionShape fallFrames(index), baseTop, baseLeft
    Next index

    link.Visible = False

    For index = LBound(fallFrames) To UBound(fallFrames)
        PlayFallFrame fallFrames(index), linkSheet
    Next index

    Relocate targetCode

    If EnsureLinkContext(gs, manager, link, linkSheet) Then
        SyncLinkState link, manager, gs
    End If

    Sheets(SHEET_DATA).Range(RANGE_FALL_SEQUENCE).Value = "N"
End Sub

Public Sub JumpDown()
    Dim gs As GameState
    Dim manager As SpriteManager
    Dim link As Shape
    Dim linkSheet As Worksheet
    If Not EnsureLinkContext(gs, manager, link, linkSheet) Then Exit Sub

    Sheets(SHEET_DATA).Range(RANGE_FALL_SEQUENCE).Value = "Y"
    Sheets(SHEET_DATA).Range(RANGE_SCROLL_COOLDOWN).Value = "0"

    Dim startCell As Range
    Set startCell = link.TopLeftCell

    Dim jumpColumn As Long
    jumpColumn = startCell.Column

    Dim jumpRow As Long
    jumpRow = CLng(Val(Mid$(gs.CodeCell, 5, 3)))
    If jumpRow = 0 Then jumpRow = startCell.Row

    Dim jumpTarget As Range
    Set jumpTarget = linkSheet.Cells(jumpRow, jumpColumn)

    Dim shadow As Shape
    Set shadow = TryGetShape(linkSheet, "LinkShadow")
    If Not shadow Is Nothing Then
        PositionShape shadow, jumpTarget.Top + 5, jumpTarget.Left - 5
        shadow.Visible = True
    End If

    Dim jumpFrames(1 To 3) As Shape
    Set jumpFrames(1) = TryGetShape(linkSheet, "LinkJump1")
    Set jumpFrames(2) = TryGetShape(linkSheet, "LinkJump2")
    Set jumpFrames(3) = TryGetShape(linkSheet, "LinkJump3")

    If jumpFrames(1) Is Nothing Or jumpFrames(2) Is Nothing Or jumpFrames(3) Is Nothing Then
        If Not shadow Is Nothing Then shadow.Visible = False
        link.Visible = True
        Sheets(SHEET_DATA).Range(RANGE_FALL_SEQUENCE).Value = "N"
        Exit Sub
    End If

    PositionShape jumpFrames(1), link.Top + 10, link.Left
    PositionShape jumpFrames(2), jumpFrames(1).Top + 30, link.Left
    PositionShape jumpFrames(3), jumpFrames(2).Top + 30, link.Left

    link.Visible = False

    Dim stage As Long
    For stage = LBound(jumpFrames) To UBound(jumpFrames)
        If stage > LBound(jumpFrames) Then
            jumpFrames(stage - 1).Visible = False
        End If

        jumpFrames(stage).Visible = True
        AdvanceJumpFrame jumpFrames(stage), link, manager, gs, linkSheet
    Next stage

    jumpFrames(UBound(jumpFrames)).Visible = False

    link.Visible = True
    SyncLinkState link, manager, gs
    gs.CodeCell = ""

    Do
        link.Top = link.Top + 4
        SyncLinkState link, manager, gs
        HandleScrollTrigger CStr(link.TopLeftCell.Value)
        TriggerFrameTick linkSheet, 10
    Loop Until link.Top >= jumpTarget.Top - 30

    If Not shadow Is Nothing Then shadow.Visible = False

    Sheets(SHEET_DATA).Range(RANGE_FALL_SEQUENCE).Value = "N"
End Sub

'####################################################################################
'#    Sword actions
'####################################################################################
Public Sub swordSwipe(ByVal actionSlot As Long, ByVal pressCount As Long)
    Dim gs As GameState
    Dim manager As SpriteManager
    Dim link As Shape
    Dim linkSheet As Worksheet
    If Not EnsureLinkContext(gs, manager, link, linkSheet) Then Exit Sub

    Dim swordUp As Shape
    Dim swordDown As Shape
    Dim swordLeft As Shape
    Dim swordRight As Shape
    Dim swipeUpLeft As Shape
    Dim swipeUpRight As Shape
    Dim swipeDownLeft As Shape
    Dim swipeDownRight As Shape

    If Not LoadSwordSprites(linkSheet, swordUp, swordDown, swordLeft, swordRight, _
                            swipeUpLeft, swipeUpRight, swipeDownLeft, swipeDownRight) Then Exit Sub

    HideShapes Array(swordUp, swordDown, swordLeft, swordRight, swipeUpLeft, swipeUpRight, swipeDownLeft, swipeDownRight)

    Dim frame1 As Shape
    Dim frame2 As Shape
    Dim frame3 As Shape

    Dim direction As String
    direction = NormalizeDirection(gs.MoveDir, gs.LastDir)

    Select Case direction
        Case "L"
            PositionShape swordUp, link.Top - 30, link.Left - 10
            PositionShape swipeUpLeft, link.Top - 30, link.Left - 50
            PositionShape swordLeft, link.Top, link.Left - 50

            Set frame1 = swordUp
            Set frame2 = swipeUpLeft
            Set frame3 = swordLeft

        Case "R"
            PositionShape swordUp, link.Top - 30, link.Left + 30
            PositionShape swipeUpRight, link.Top - 30, link.Left + 45
            PositionShape swordRight, link.Top, link.Left + 45

            Set frame1 = swordUp
            Set frame2 = swipeUpRight
            Set frame3 = swordRight

        Case "U", "RU", "LU"
            PositionShape swordUp, link.Top - 45, link.Left + 5
            PositionShape swipeUpRight, link.Top - 45, link.Left + 25
            PositionShape swordRight, link.Top - 15, link.Left + 35

            Set frame1 = swordRight
            Set frame2 = swipeUpRight
            Set frame3 = swordUp

        Case Else   ' Down-facing including diagonals
            PositionShape swordLeft, link.Top, link.Left - 50
            PositionShape swipeDownLeft, link.Top + 30, link.Left - 45
            PositionShape swordDown, link.Top + 40, link.Left - 25

            Set frame1 = swordLeft
            Set frame2 = swipeDownLeft
            Set frame3 = swordDown
    End Select

    If frame1 Is Nothing Or frame2 Is Nothing Or frame3 Is Nothing Then Exit Sub

    Select Case pressCount
        Case Is <= 1
            frame1.Visible = True
            TriggerFrameTick linkSheet, 25

            frame1.Visible = False
            frame2.Visible = True
            TriggerFrameTick linkSheet, 25

            frame2.Visible = False
            frame3.Visible = True
            TriggerFrameTick linkSheet, 25

            RegisterSwordHits frame3
            frame3.Visible = False

        Case 2 To 20
            ' Charging phase – keep sword hidden but retain context.

        Case Else
            frame3.Visible = True
            RegisterSwordHits frame3
    End Select
End Sub

Public Sub swordSpin()
    Dim gs As GameState
    Dim manager As SpriteManager
    Dim link As Shape
    Dim linkSheet As Worksheet
    If Not EnsureLinkContext(gs, manager, link, linkSheet) Then Exit Sub

    Dim swordUp As Shape
    Dim swordDown As Shape
    Dim swordLeft As Shape
    Dim swordRight As Shape
    Dim swipeUpLeft As Shape
    Dim swipeUpRight As Shape
    Dim swipeDownLeft As Shape
    Dim swipeDownRight As Shape

    If Not LoadSwordSprites(linkSheet, swordUp, swordDown, swordLeft, swordRight, _
                            swipeUpLeft, swipeUpRight, swipeDownLeft, swipeDownRight) Then Exit Sub

    Dim linkLeft1 As Shape
    Dim linkRight1 As Shape
    Dim linkUp1 As Shape
    Dim linkDown1 As Shape

    Call LoadLinkFacingSprites(linkSheet, linkLeft1, linkRight1, linkUp1, linkDown1)

    HideShapes Array(swordUp, swordDown, swordLeft, swordRight, swipeUpLeft, swipeUpRight, swipeDownLeft, swipeDownRight)
    HideShapes Array(linkLeft1, linkRight1, linkUp1, linkDown1)

    PositionShape swordUp, link.Top - 30, link.Left
    PositionShape swordRight, link.Top, link.Left + 35
    PositionShape swordLeft, link.Top, link.Left - 50
    PositionShape swordDown, link.Top + 40, link.Left - 25
    PositionShape swipeUpLeft, link.Top - 30, link.Left - 50
    PositionShape swipeUpRight, link.Top - 45, link.Left + 25
    PositionShape swipeDownRight, link.Top + 45, link.Left + 35
    PositionShape swipeDownLeft, link.Top + 30, link.Left - 45

    Dim swordOrder(1 To 8) As Shape
    Dim linkOrder(1 To 4) As Shape

    Dim direction As String
    direction = NormalizeDirection(gs.MoveDir, gs.LastDir)

    Select Case direction
        Case "L"
            Set swordOrder(1) = swordLeft
            Set swordOrder(2) = swipeDownLeft
            Set swordOrder(3) = swordDown
            Set swordOrder(4) = swipeDownRight
            Set swordOrder(5) = swordRight
            Set swordOrder(6) = swipeUpRight
            Set swordOrder(7) = swordUp
            Set swordOrder(8) = swipeUpLeft

            Set linkOrder(1) = linkLeft1
            Set linkOrder(2) = linkDown1
            Set linkOrder(3) = linkRight1
            Set linkOrder(4) = linkUp1

        Case "R", "RU", "LU"
            Set swordOrder(1) = swordRight
            Set swordOrder(2) = swipeDownRight
            Set swordOrder(3) = swordDown
            Set swordOrder(4) = swipeDownLeft
            Set swordOrder(5) = swordLeft
            Set swordOrder(6) = swipeUpLeft
            Set swordOrder(7) = swordUp
            Set swordOrder(8) = swipeUpRight

            Set linkOrder(1) = linkRight1
            Set linkOrder(2) = linkDown1
            Set linkOrder(3) = linkLeft1
            Set linkOrder(4) = linkUp1

        Case "U"
            Set swordOrder(1) = swordUp
            Set swordOrder(2) = swipeUpLeft
            Set swordOrder(3) = swordLeft
            Set swordOrder(4) = swipeDownLeft
            Set swordOrder(5) = swordDown
            Set swordOrder(6) = swipeDownRight
            Set swordOrder(7) = swordRight
            Set swordOrder(8) = swipeUpRight

            Set linkOrder(1) = linkUp1
            Set linkOrder(2) = linkLeft1
            Set linkOrder(3) = linkDown1
            Set linkOrder(4) = linkRight1

        Case Else
            Set swordOrder(1) = swordDown
            Set swordOrder(2) = swipeDownRight
            Set swordOrder(3) = swordRight
            Set swordOrder(4) = swipeUpRight
            Set swordOrder(5) = swordUp
            Set swordOrder(6) = swipeUpLeft
            Set swordOrder(7) = swordLeft
            Set swordOrder(8) = swipeDownLeft

            Set linkOrder(1) = linkDown1
            Set linkOrder(2) = linkRight1
            Set linkOrder(3) = linkUp1
            Set linkOrder(4) = linkLeft1
    End Select

    If swordOrder(1) Is Nothing Then Exit Sub

    Dim stepIndex As Long
    For stepIndex = LBound(swordOrder) To UBound(swordOrder)
        swordOrder(stepIndex).Visible = True
        If stepIndex > LBound(swordOrder) Then
            swordOrder(stepIndex - 1).Visible = False
        End If

        Select Case stepIndex
            Case 1
                If Not linkOrder(1) Is Nothing Then linkOrder(1).Visible = True
            Case 3
                If Not linkOrder(1) Is Nothing Then linkOrder(1).Visible = False
                If Not linkOrder(2) Is Nothing Then linkOrder(2).Visible = True
            Case 5
                If Not linkOrder(2) Is Nothing Then linkOrder(2).Visible = False
                If Not linkOrder(3) Is Nothing Then linkOrder(3).Visible = True
            Case 7
                If Not linkOrder(3) Is Nothing Then linkOrder(3).Visible = False
                If Not linkOrder(4) Is Nothing Then linkOrder(4).Visible = True
        End Select

        TriggerFrameTick linkSheet, 25
    Next stepIndex

    swordOrder(UBound(swordOrder)).Visible = False
    swordOrder(LBound(swordOrder)).Visible = True
    If Not linkOrder(4) Is Nothing Then linkOrder(4).Visible = False
    If Not linkOrder(1) Is Nothing Then linkOrder(1).Visible = True
    TriggerFrameTick linkSheet, 25

    swordOrder(LBound(swordOrder)).Visible = False
    If Not linkOrder(1) Is Nothing Then linkOrder(1).Visible = False
End Sub

Public Sub showShield()
    Dim gs As GameState
    Dim manager As SpriteManager
    Dim link As Shape
    Dim linkSheet As Worksheet
    If Not EnsureLinkContext(gs, manager, link, linkSheet) Then Exit Sub

    Sheets(SHEET_DATA).Range(RANGE_SHIELD_STATE).Value = "Y"

    Dim shieldUp As Shape
    Dim shieldDown As Shape
    Dim shieldLeft As Shape
    Dim shieldRight As Shape

    If Not LoadShieldSprites(linkSheet, shieldUp, shieldDown, shieldLeft, shieldRight) Then Exit Sub

    HideShapes Array(shieldUp, shieldDown, shieldLeft, shieldRight)

    Dim direction As String
    direction = NormalizeDirection(gs.MoveDir, gs.LastDir)

    Dim shieldSprite As Shape

    Select Case direction
        Case "D", "LD", "RD"
            Set shieldSprite = shieldDown
        Case "U", "RU", "LU"
            Set shieldSprite = shieldUp
        Case "L"
            Set shieldSprite = shieldLeft
        Case "R"
            Set shieldSprite = shieldRight
        Case Else
            Set shieldSprite = shieldDown
    End Select

    If shieldSprite Is Nothing Then Exit Sub

    PositionShape shieldSprite, link.Top, link.Left
    shieldSprite.Visible = True
End Sub

