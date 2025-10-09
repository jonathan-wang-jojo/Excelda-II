Option Explicit

'###################################################################################
'                              ENEMY MANAGEMENT
'###################################################################################
' Handles all enemy state and operations
'###################################################################################

Private Type EnemyData
    Name As String
    Hit As Long
End Type

Private m_Enemies(1 To 4) As EnemyData

' Initialization
Private Sub Class_Initialize()
    Reset
End Sub

' Core lifecycle methods
Public Sub Reset()
    Dim i As Long
    For i = 1 To 4
        m_Enemies(i).Name = ""
        m_Enemies(i).Hit = 0
    Next i
End Sub

Public Sub Destroy()
    Reset
End Sub

' Core enemy processing - delegates to existing functions
Public Sub ProcessEnemy(ByVal index As Integer, ByVal linkSprite As Object)
    If index < 1 Or index > 4 Or m_Enemies(index).Name = "" Then Exit Sub
    
    Call enemyCollision(linkSprite, m_Enemies(index).Name)
    
    If m_Enemies(index).Hit > 0 Then
        Call enemyBounceBack(index)
        m_Enemies(index).Hit = m_Enemies(index).Hit - 1
    Else
        Call RNDEnemyMove(index)
    End If
End Sub

' Simple accessors
Public Property Get Name(ByVal index As Integer) As String
    If index >= 1 And index <= 4 Then Name = m_Enemies(index).Name
End Property

Public Property Let Name(ByVal index As Integer, ByVal Value As String)
    If index >= 1 And index <= 4 Then m_Enemies(index).Name = Value
End Property

Public Property Get Hit(ByVal index As Integer) As Long
    If index >= 1 And index <= 4 Then Hit = m_Enemies(index).Hit
End Property

Public Property Let Hit(ByVal index As Integer, ByVal Value As Long)
    If index >= 1 And index <= 4 Then m_Enemies(index).Hit = Value
End Property