'Attribute VB_Name = "AH_Enemies"
Option Explicit

'###################################################################################
'                              ENEMY SPAWNING SYSTEM
'###################################################################################
' Data-driven enemy spawning using Excel sheet as single source of truth
' No more global variables - everything managed by EnemyManager
'###################################################################################

' Enemy type to Data sheet row mapping
Private Const ENEMY_DATA_START_ROW As Long = 46  ' First enemy data row (Marin)

' Enemy type registry - maps trigger codes to enemy types and rows
Private Type EnemyTypeData
    TypeName As String
    DisplayName As String
    BaseRow As Long
End Type

Private m_EnemyTypes() As EnemyTypeData
Private m_Initialized As Boolean

'###################################################################################
'                              INITIALIZATION
'###################################################################################

Private Sub InitializeEnemyRegistry()
    ' Build enemy registry from actual Data sheet structure
    If m_Initialized Then Exit Sub
    
    ReDim m_EnemyTypes(1 To 20)
    
    Dim i As Long: i = 1
    
    ' NPCs (rows 46-49)
    m_EnemyTypes(i).TypeName = "Marin"
    m_EnemyTypes(i).DisplayName = "Marin"
    m_EnemyTypes(i).BaseRow = 46
    i = i + 1
    
    m_EnemyTypes(i).TypeName = "Tarin"
    m_EnemyTypes(i).DisplayName = "Tarin"
    m_EnemyTypes(i).BaseRow = 47
    i = i + 1
    
    m_EnemyTypes(i).TypeName = "Raccoon"
    m_EnemyTypes(i).DisplayName = "Raccoon"
    m_EnemyTypes(i).BaseRow = 49
    i = i + 1
    
    ' Skeletons (rows 52-53)
    m_EnemyTypes(i).TypeName = "skeleton"
    m_EnemyTypes(i).DisplayName = "skeleton"
    m_EnemyTypes(i).BaseRow = 52
    i = i + 1

    ' Sandcrabs (rows 54-55)
    m_EnemyTypes(i).TypeName = "Sandcrab"
    m_EnemyTypes(i).DisplayName = "sandcrab"
    m_EnemyTypes(i).BaseRow = 54
    i = i + 1
    
    ' Octoroks (rows 56-57)
    m_EnemyTypes(i).TypeName = "Octorok"
    m_EnemyTypes(i).DisplayName = "Octorok"
    m_EnemyTypes(i).BaseRow = 56
    i = i + 1
    
    ' Sandspinners (rows 58-59)
    m_EnemyTypes(i).TypeName = "Sandspinner"
    m_EnemyTypes(i).DisplayName = "Sandspinner"
    m_EnemyTypes(i).BaseRow = 58
    i = i + 1
    
    ' Gordos (rows 60-62)
    m_EnemyTypes(i).TypeName = "Gordo"
    m_EnemyTypes(i).DisplayName = "gordo"
    m_EnemyTypes(i).BaseRow = 60
    i = i + 1
    
    ' Moblins (row 66+)
    m_EnemyTypes(i).TypeName = "Moblin"
    m_EnemyTypes(i).DisplayName = "Moblin"
    m_EnemyTypes(i).BaseRow = 66

    m_Initialized = True
End Sub

Private Function GetEnemyRow(enemyType As String, slotNumber As Long) As Long
    ' Get data row for enemy type and slot
    Call InitializeEnemyRegistry
    
    Dim i As Long
    For i = LBound(m_EnemyTypes) To UBound(m_EnemyTypes)
        If m_EnemyTypes(i).TypeName = enemyType Then
            GetEnemyRow = m_EnemyTypes(i).BaseRow + slotNumber - 1
            Exit Function
        End If
    Next i
    
    GetEnemyRow = 0
End Function

'###################################################################################
'                              UNIFIED SPAWNING FUNCTIONS
'###################################################################################

Public Sub ShowEnemy(enemyType As String, slotNumber As Long)
    ' Universal enemy spawner - works for all enemy types
    On Error Resume Next
    
    Dim manager As EnemyManager
    Set manager = EnemyManagerInstance()

    Dim dataRow As Long
    dataRow = GetEnemyRow(enemyType, slotNumber)
    If dataRow = 0 Then Exit Sub

    Dim gs As GameState
    Set gs = GameStateInstance()
    Dim anchorAddress As String
    If Not gs Is Nothing Then anchorAddress = Trim$(gs.TriggerCellAddress)
    If anchorAddress = "" Then Exit Sub
    
    ' Spawn through manager
    manager.SpawnEnemy enemyType, slotNumber, dataRow, anchorAddress
End Sub

Public Sub HideEnemy(enemyType As String, slotNumber As Long)
    ' Universal enemy despawner
    On Error Resume Next
    
    Dim dataRow As Long
    dataRow = GetEnemyRow(enemyType, slotNumber)
    
    Dim manager As EnemyManager
    Set manager = EnemyManagerInstance()
    
    manager.DespawnEnemy slotNumber, dataRow
End Sub