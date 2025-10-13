Option Explicit

'###################################################################################
'                              SPRITE MANAGEMENT
'###################################################################################
' Handles all sprite operations - positions, frames, visibility
'###################################################################################

Private Type SpriteData
    LinkSprite As Shape
    Top As Double
    Left As Double
    Frame As Integer
    FrameCounter As Long
End Type

Private m_Sprite As SpriteData

' Initialization
Private Sub Class_Initialize()
    m_Sprite.Frame = 1
    m_Sprite.FrameCounter = 0
End Sub

Public Sub Initialize(ByVal screenName As String, ByVal spriteName As String)
    On Error GoTo ErrorHandler
    
    ' Find and set the sprite
    Set m_Sprite.LinkSprite = Sheets(screenName).Shapes(spriteName)
    
    If m_Sprite.LinkSprite Is Nothing Then
        Err.Raise vbObjectError + 1, "SpriteManager.Initialize", "Sprite '" & spriteName & "' not found on " & screenName
    End If
    
    ' Set initial position from sprite's actual location
    m_Sprite.Top = m_Sprite.LinkSprite.Top
    m_Sprite.Left = m_Sprite.LinkSprite.Left
    Sheets(SHEET_DATA).Range(RANGE_ANIM_COUNTER).Value = 0
    
    ' Determine frame from sprite name (LinkDown1 = frame 1, LinkDown2 = frame 2, etc.)
    If InStr(spriteName, "1") > 0 Then
        m_Sprite.Frame = 1
    ElseIf InStr(spriteName, "2") > 0 Then
        m_Sprite.Frame = 2
    Else
        m_Sprite.Frame = 2  ' Default to frame 2
    End If
    
    Exit Sub

ErrorHandler:
    ' Log error but don't crash - let caller handle it
    Debug.Print "SpriteManager.Initialize Error: " & Err.Description
    On Error GoTo 0
End Sub

' Core lifecycle
Public Sub Reset()
    On Error Resume Next
    
    Dim gs As GameState
    Set gs = GameStateInstance
    
    If Not gs Is Nothing And gs.CurrentScreen <> "" Then
        ' Try to reset to last known sprite or default to LinkDown2
        Set m_Sprite.LinkSprite = Sheets(gs.CurrentScreen).Shapes("LinkDown2")
        
        If Not m_Sprite.LinkSprite Is Nothing Then
            m_Sprite.Top = m_Sprite.LinkSprite.Top
            m_Sprite.Left = m_Sprite.LinkSprite.Left
            m_Sprite.Frame = 1
            m_Sprite.FrameCounter = 0
            Sheets(SHEET_DATA).Range(RANGE_ANIM_COUNTER).Value = 0
        End If
    End If
    
    On Error GoTo 0
End Sub

Public Sub Destroy()
    Set m_Sprite.LinkSprite = Nothing
    m_Sprite.Top = 0
    m_Sprite.Left = 0
    m_Sprite.Frame = 0
    m_Sprite.FrameCounter = 0
End Sub


' Essential properties
Public Property Get LinkSprite() As Object
    Set LinkSprite = m_Sprite.LinkSprite
End Property

Public Property Get LinkSpriteTop() As Double
    LinkSpriteTop = m_Sprite.Top
End Property

Public Property Let LinkSpriteTop(ByVal Value As Double)
    m_Sprite.Top = Value
End Property

Public Property Get LinkSpriteLeft() As Double
    LinkSpriteLeft = m_Sprite.Left
End Property

Public Property Let LinkSpriteLeft(ByVal Value As Double)
    m_Sprite.Left = Value
End Property

Public Property Get LinkSpriteFrame() As Integer
    LinkSpriteFrame = m_Sprite.Frame
End Property

Public Property Let LinkSpriteFrame(ByVal Value As Integer)
    m_Sprite.Frame = Value
End Property

' Core sprite operations - delegates to optimized logic
Public Sub UpdateFrame(ByVal movementDir As String, ByVal facingDir As String, ByVal moveSpeed As Long)
    On Error GoTo CleanExit
    Dim gs As GameState
    Set gs = GameStateInstance

    If gs Is Nothing Or gs.CurrentScreen = "" Then GoTo CleanExit

    m_Sprite.FrameCounter = (m_Sprite.FrameCounter + 1) Mod 10
    If m_Sprite.FrameCounter >= 5 Then
        m_Sprite.Frame = 1
    Else
        m_Sprite.Frame = 2
    End If

    On Error Resume Next
    Sheets(SHEET_DATA).Range(RANGE_ANIM_COUNTER).Value = m_Sprite.FrameCounter
    On Error GoTo CleanExit
    
    If facingDir = "" Then
        If Not gs Is Nothing Then
            facingDir = gs.LastDir
        Else
            facingDir = ""
        End If
    End If
    Dim baseDir As String
    baseDir = ResolveDirection(facingDir)
    
    Dim ws As Worksheet
    Set ws = Sheets(gs.CurrentScreen)

    Dim target As Shape
    Dim targetName As String
    targetName = "Link" & baseDir & CStr(m_Sprite.Frame)
    Set target = GetShapeIfExists(ws, targetName)
    If target Is Nothing Then
        Dim fallbackFrame As Integer
        fallbackFrame = IIf(m_Sprite.Frame = 1, 2, 1)
        targetName = "Link" & baseDir & CStr(fallbackFrame)
        Set target = GetShapeIfExists(ws, targetName)
        If target Is Nothing Then
            Debug.Print "SpriteManager.UpdateFrame missing sprite: " & targetName
            GoTo CleanExit
        End If
        m_Sprite.Frame = fallbackFrame
    End If
    Set m_Sprite.LinkSprite = target
    
    ' Sync global for legacy code compatibility
    Set LinkSprite = m_Sprite.LinkSprite
    
    Dim newTop As Double, newLeft As Double
    newTop = m_Sprite.Top
    newLeft = m_Sprite.Left

    ' Adjust position based on movement components
    If InStr(movementDir, "U") > 0 Then newTop = newTop - moveSpeed
    If InStr(movementDir, "D") > 0 Then newTop = newTop + moveSpeed
    If InStr(movementDir, "L") > 0 Then newLeft = newLeft - moveSpeed
    If InStr(movementDir, "R") > 0 Then newLeft = newLeft + moveSpeed

    m_Sprite.Top = newTop
    m_Sprite.Left = newLeft
    
CleanExit:
    On Error GoTo 0
End Sub

Public Sub UpdatePosition()
    On Error Resume Next
    m_Sprite.LinkSprite.Top = m_Sprite.Top
    m_Sprite.LinkSprite.Left = m_Sprite.Left
    SyncAllLinkFrames
    On Error GoTo 0
End Sub

Public Sub UpdateVisibility()
    On Error Resume Next
    
    Dim wsGame As Worksheet: Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    Dim s As Variant
    
    For Each s In Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", _
                        "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
        wsGame.Shapes(s).Visible = False
    Next s
    
    If Not m_Sprite.LinkSprite Is Nothing Then
        m_Sprite.LinkSprite.Visible = True
    End If
    
    On Error GoTo 0
End Sub

Public Sub AlignSprites(ByVal leftPos As Double, ByVal topPos As Double)
    On Error Resume Next
    
    m_Sprite.Left = leftPos
    m_Sprite.Top = topPos
    
    Dim wsGame As Worksheet: Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    Dim s As Variant
    
    For Each s In Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", _
                        "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
        wsGame.Shapes(s).Left = leftPos
        wsGame.Shapes(s).Top = topPos
    Next s
    
    On Error GoTo 0
End Sub

Private Function ResolveDirection(ByVal direction As String) As String
    If InStr(direction, "U") > 0 Then
        ResolveDirection = "Up"
    ElseIf InStr(direction, "D") > 0 Then
        ResolveDirection = "Down"
    ElseIf InStr(direction, "L") > 0 Then
        ResolveDirection = "Left"
    ElseIf InStr(direction, "R") > 0 Then
        ResolveDirection = "Right"
    Else
        ResolveDirection = "Down"
    End If
End Function

Private Sub SyncAllLinkFrames()
    Dim wsGame As Worksheet
    On Error Resume Next
    Set wsGame = Sheets(GameStateInstance.CurrentScreen)
    If wsGame Is Nothing Then Exit Sub
    Dim s As Variant
    For Each s In Array("LinkUp1", "LinkUp2", "LinkDown1", "LinkDown2", _
                        "LinkLeft1", "LinkLeft2", "LinkRight1", "LinkRight2")
        wsGame.Shapes(s).Top = m_Sprite.Top
        wsGame.Shapes(s).Left = m_Sprite.Left
    Next s
    On Error GoTo 0
End Sub

Private Function GetShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    On Error Resume Next
    Set GetShapeIfExists = ws.Shapes(shapeName)
    If Err.Number <> 0 Then
        Err.Clear
        Set GetShapeIfExists = Nothing
    End If
    On Error GoTo 0
End Function