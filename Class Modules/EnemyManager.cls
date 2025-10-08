Option Explicit

'###################################################################################
'                              ENEMY MANAGEMENT
'###################################################################################
' Handles all enemy state and operations
'###################################################################################

Private Type EnemyState
    Name(1 To 4) As String
    Hit(1 To 4) As Long
End Type

Private m_Enemy As EnemyState

' Initialization
Private Sub Class_Initialize()
    InitializeEnemies
End Sub

Private Sub InitializeEnemies()
    Dim i As Long
    For i = 1 To 4
        m_Enemy.Name(i) = ""
        m_Enemy.Hit(i) = 0
    Next i
End Sub

' Reset enemy states (for restarts)
Public Sub Reset()
    Dim i As Long
    For i = 1 To 4
        m_Enemy.Hit(i) = 0
    Next i
End Sub

' Cleanup (for full game quit)
Public Sub Destroy()
    Dim i As Long
    For i = 1 To 4
        m_Enemy.Name(i) = ""
        m_Enemy.Hit(i) = 0
    Next i
End Sub

' Enemy processing
Public Sub HandleEnemy(ByVal enemyIndex As Integer, ByVal linkSprite As Object)
    If enemyIndex < 1 Or enemyIndex > 4 Then Exit Sub

    Dim enemyName As String
    enemyName = m_Enemy.Name(enemyIndex)
    If enemyName = "" Then Exit Sub

    Call enemyCollision(linkSprite, enemyName)

    If m_Enemy.Hit(enemyIndex) > 0 Then
        Call enemyBounceBack(enemyIndex)
    Else
        Call RNDEnemyMove(enemyIndex)
    End If
End Sub

' Enemy property accessors
Public Property Get EnemyName(ByVal index As Integer) As String
    EnemyName = m_Enemy.Name(index)
End Property

Public Property Let EnemyName(ByVal index As Integer, ByVal Value As String)
    m_Enemy.Name(index) = Value
End Property

Public Property Get EnemyHit(ByVal index As Integer) As Long
    EnemyHit = m_Enemy.Hit(index)
End Property

Public Property Let EnemyHit(ByVal index As Integer, ByVal Value As Long)
    m_Enemy.Hit(index) = Value
End Property