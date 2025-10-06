Option Explicit

'###################################################################################
'                              ENEMY MANAGEMENT
'###################################################################################
' Handles all enemy state and operations
'###################################################################################

' Enemy state
Private Type EnemyState
    RNDenemyName1 As String
    RNDenemyName2 As String
    RNDenemyName3 As String
    RNDenemyName4 As String
    RNDenemyHit1 As Long
    RNDenemyHit2 As Long
    RNDenemyHit3 As Long
    RNDenemyHit4 As Long
End Type

Private m_Enemy As EnemyState

' Singleton pattern
Private m_Instance As EnemyManager
Private m_Initialized As Boolean

Private Sub Class_Initialize()
    If Not m_Initialized Then
        InitializeEnemies
        m_Initialized = True
    End If
End Sub

Private Sub InitializeEnemies()
    ' Initialize enemy state (values will be set by other modules)
    With m_Enemy
        .RNDenemyName1 = ""
        .RNDenemyName2 = ""
        .RNDenemyName3 = ""
        .RNDenemyName4 = ""
        .RNDenemyHit1 = 0
        .RNDenemyHit2 = 0
        .RNDenemyHit3 = 0
        .RNDenemyHit4 = 0
    End With
End Sub

' Enemy name properties
Public Property Get RNDenemyName1() As String
    RNDenemyName1 = m_Enemy.RNDenemyName1
End Property

Public Property Let RNDenemyName1(ByVal Value As String)
    m_Enemy.RNDenemyName1 = Value
End Property

Public Property Get RNDenemyName2() As String
    RNDenemyName2 = m_Enemy.RNDenemyName2
End Property

Public Property Let RNDenemyName2(ByVal Value As String)
    m_Enemy.RNDenemyName2 = Value
End Property

Public Property Get RNDenemyName3() As String
    RNDenemyName3 = m_Enemy.RNDenemyName3
End Property

Public Property Let RNDenemyName3(ByVal Value As String)
    m_Enemy.RNDenemyName3 = Value
End Property

Public Property Get RNDenemyName4() As String
    RNDenemyName4 = m_Enemy.RNDenemyName4
End Property

Public Property Let RNDenemyName4(ByVal Value As String)
    m_Enemy.RNDenemyName4 = Value
End Property

' Enemy hit properties
Public Property Get RNDenemyHit1() As Long
    RNDenemyHit1 = m_Enemy.RNDenemyHit1
End Property

Public Property Let RNDenemyHit1(ByVal Value As Long)
    m_Enemy.RNDenemyHit1 = Value
End Property

Public Property Get RNDenemyHit2() As Long
    RNDenemyHit2 = m_Enemy.RNDenemyHit2
End Property

Public Property Let RNDenemyHit2(ByVal Value As Long)
    m_Enemy.RNDenemyHit2 = Value
End Property

Public Property Get RNDenemyHit3() As Long
    RNDenemyHit3 = m_Enemy.RNDenemyHit3
End Property

Public Property Let RNDenemyHit3(ByVal Value As Long)
    m_Enemy.RNDenemyHit3 = Value
End Property

Public Property Get RNDenemyHit4() As Long
    RNDenemyHit4 = m_Enemy.RNDenemyHit4
End Property

Public Property Let RNDenemyHit4(ByVal Value As Long)
    m_Enemy.RNDenemyHit4 = Value
End Property

' Enemy methods
Public Sub HandleEnemy(ByVal enemyIndex As Integer, ByVal linkSprite As Object)
    ' Get enemy name based on index
    Dim enemyName As String
    Select Case enemyIndex
        Case 1: enemyName = m_Enemy.RNDenemyName1
        Case 2: enemyName = m_Enemy.RNDenemyName2
        Case 3: enemyName = m_Enemy.RNDenemyName3
        Case 4: enemyName = m_Enemy.RNDenemyName4
        Case Else: Exit Sub
    End Select
    
    ' Skip if enemy doesn't exist
    If enemyName = "" Then Exit Sub
    
    ' Process enemy
    Call enemyCollision(linkSprite, enemyName)
    
    ' Check hit state
    Dim isHit As Boolean
    Select Case enemyIndex
        Case 1: isHit = (m_Enemy.RNDenemyHit1 > 0)
        Case 2: isHit = (m_Enemy.RNDenemyHit2 > 0)
        Case 3: isHit = (m_Enemy.RNDenemyHit3 > 0)
        Case 4: isHit = (m_Enemy.RNDenemyHit4 > 0)
    End Select
    
    ' Handle hit or movement
    If isHit Then
        Call enemyBounceBack(enemyIndex)
    Else
        Call RNDEnemyMove(enemyIndex)
    End If
End Sub

' Singleton accessor
Public Property Get Instance() As EnemyManager
    If m_Instance Is Nothing Then
        Set m_Instance = New EnemyManager
    End If
    Set Instance = m_Instance
End Property
