Option Explicit

'===================================================================================
'                              SPRITE COMPONENT
'===================================================================================
' Handles sprite rendering and animation via BatchRenderer
' Automatically syncs sprite position with Transform component
'===================================================================================

Private m_Transform As TransformComponent
Private m_ShapeName As String
Private m_IsVisible As Boolean

' Animation support
Private m_AnimFrames As Variant ' Array of frame names (e.g., ["octorok_1", "octorok_2"])
Private m_CurrentFrame As Long
Private m_FrameTime As Double
Private m_FrameDuration As Double ' Seconds per frame

'===================================================================================
'                              INITIALIZATION
'===================================================================================

Public Sub Init(transform As TransformComponent)
    Set m_Transform = transform
    m_IsVisible = False
    m_CurrentFrame = 0
    m_FrameTime = 0
    m_FrameDuration = 0.15 ' 150ms per frame (classic Zelda timing)
    m_AnimFrames = Empty
End Sub

'===================================================================================
'                              SPRITE LOADING
'===================================================================================

' Load static sprite (no animation)
Public Sub LoadSprite(shapeName As String)
    m_ShapeName = shapeName
    m_IsVisible = True
    m_AnimFrames = Empty ' Clear animation
    
    ' Queue initial render via BatchRenderer
    Singletons.BatchRenderer.QueueSprite m_ShapeName, m_Transform.x, m_Transform.y, True
End Sub

' Load animated sprite (multiple frames)
Public Sub LoadAnimation(frameNames As Variant, Optional frameDuration As Double = 0.15)
    If Not IsArray(frameNames) Then Exit Sub
    If UBound(frameNames) < 0 Then Exit Sub
    
    m_AnimFrames = frameNames
    m_CurrentFrame = 0
    m_FrameTime = 0
    m_FrameDuration = frameDuration
    m_ShapeName = frameNames(0) ' Start with first frame
    m_IsVisible = True
    
    ' Queue initial render
    Singletons.BatchRenderer.QueueSprite m_ShapeName, m_Transform.x, m_Transform.y, True
End Sub

'===================================================================================
'                              UPDATE
'===================================================================================

Public Sub Update()
    If Not m_IsVisible Then Exit Sub
    
    ' Handle animation frame updates
    If Not IsEmpty(m_AnimFrames) Then
        UpdateAnimation
    End If
    
    ' Queue sprite position update via BatchRenderer
    Singletons.BatchRenderer.QueueSprite m_ShapeName, m_Transform.x, m_Transform.y, m_IsVisible
End Sub

Private Sub UpdateAnimation()
    m_FrameTime = m_FrameTime + 0.016 ' ~60fps (16ms per frame)
    
    If m_FrameTime >= m_FrameDuration Then
        m_FrameTime = 0
        m_CurrentFrame = (m_CurrentFrame + 1) Mod (UBound(m_AnimFrames) + 1)
        m_ShapeName = m_AnimFrames(m_CurrentFrame)
    End If
End Sub

'===================================================================================
'                              VISIBILITY
'===================================================================================

Public Sub Show()
    m_IsVisible = True
    Singletons.BatchRenderer.QueueSprite m_ShapeName, m_Transform.x, m_Transform.y, True
End Sub

Public Sub Hide()
    m_IsVisible = False
    Singletons.BatchRenderer.QueueSprite m_ShapeName, m_Transform.x, m_Transform.y, False
End Sub

'===================================================================================
'                              PROPERTIES
'===================================================================================

Public Property Get ShapeName() As String
    ShapeName = m_ShapeName
End Property

Public Property Get IsVisible() As Boolean
    IsVisible = m_IsVisible
End Property

Public Property Get FrameDuration() As Double
    FrameDuration = m_FrameDuration
End Property

Public Property Let FrameDuration(value As Double)
    m_FrameDuration = value
End Property

' Manually set current frame (for direction-based animations)
Public Sub SetFrame(frameName As String)
    m_ShapeName = frameName
End Sub

' Get current frame index (for debugging)
Public Property Get CurrentFrameIndex() As Long
    CurrentFrameIndex = m_CurrentFrame
End Property
