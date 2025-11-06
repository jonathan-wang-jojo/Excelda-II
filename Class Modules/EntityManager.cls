'@Folder("Entity.Core")
Option Explicit

'===================================================================================
'                              ENTITY MANAGER
'===================================================================================
' Central manager for all entities (enemies, NPCs, objects, pickups)
' Handles creation, updating, and cleanup
'===================================================================================

Private m_Entities As Collection
Private m_NextID As Long

'===================================================================================
'                              INITIALIZATION
'===================================================================================

Private Sub Class_Initialize()
    Set m_Entities = New Collection
    m_NextID = 1
End Sub

'===================================================================================
'                              ENTITY CREATION
'===================================================================================

' Create entity with auto-generated ID
Public Function CreateEntity(entityType As String, x As Double, y As Double) As Entity
    Dim entity As Entity
    Set entity = New Entity
    
    ' Initialize with unique ID
    Dim entityID As String
    entityID = entityType & "_" & m_NextID
    m_NextID = m_NextID + 1
    
    entity.Create entityID, entityType, x, y
    
    ' Add to collection
    m_Entities.Add entity, entityID
    
    Set CreateEntity = entity
End Function

' Create entity with specific ID (for loading from save data)
Public Function CreateEntityWithID(entityID As String, entityType As String, x As Double, y As Double) As Entity
    Dim entity As Entity
    Set entity = New Entity
    
    entity.Create entityID, entityType, x, y
    
    ' Add to collection
    m_Entities.Add entity, entityID
    
    Set CreateEntityWithID = entity
End Function

'===================================================================================
'                              UPDATE LOOP
'===================================================================================

Public Sub UpdateAll(deltaTime As Double)
    Dim entity As Entity
    Dim i As Long
    
    ' Update all active entities
    For i = m_Entities.Count To 1 Step -1
        Set entity = m_Entities(i)
        
        If entity.IsActive Then
            entity.Update deltaTime
            
            ' Auto-remove dead entities with health component
            If entity.HasHealth Then
                If entity.Health.IsDead Then
                    entity.Destroy
                End If
            End If
        End If
        
        ' Remove inactive entities from collection
        If Not entity.IsActive Then
            m_Entities.Remove i
        End If
    Next i
End Sub

'===================================================================================
'                              QUERIES
'===================================================================================

' Find entity by ID
Public Function GetEntity(entityID As String) As Entity
    On Error Resume Next
    Set GetEntity = m_Entities(entityID)
    On Error GoTo 0
End Function

' Get all entities of a specific type
Public Function GetEntitiesByType(entityType As String) As Collection
    Dim results As Collection
    Set results = New Collection
    
    Dim entity As Entity
    For Each entity In m_Entities
        If entity.EntityType = entityType Then
            results.Add entity
        End If
    Next entity
    
    Set GetEntitiesByType = results
End Function

' Get all active entities
Public Function GetAllEntities() As Collection
    Set GetAllEntities = m_Entities
End Function

' Get entity count
Public Property Get Count() As Long
    Count = m_Entities.Count
End Property

' Get entities near a position (for AI detection)
Public Function GetEntitiesNear(x As Double, y As Double, radius As Double) As Collection
    Dim results As Collection
    Set results = New Collection
    
    Dim entity As Entity
    Dim distance As Double
    
    For Each entity In m_Entities
        If entity.IsActive Then
            distance = entity.Transform.DistanceTo(x, y)
            If distance <= radius Then
                results.Add entity
            End If
        End If
    Next entity
    
    Set GetEntitiesNear = results
End Function

'===================================================================================
'                              CLEANUP
'===================================================================================

' Clear all entities (scene transitions, game restart)
Public Sub Clear()
    Dim entity As Entity
    For Each entity In m_Entities
        entity.Destroy
    Next entity
    
    Set m_Entities = New Collection
    m_NextID = 1
End Sub

' Remove specific entity by ID
Public Sub RemoveEntity(entityID As String)
    On Error Resume Next
    Dim entity As Entity
    Set entity = m_Entities(entityID)
    
    If Not entity Is Nothing Then
        entity.Destroy
        m_Entities.Remove entityID
    End If
    On Error GoTo 0
End Sub

' Remove all entities of a specific type
Public Sub RemoveEntitiesByType(entityType As String)
    Dim entity As Entity
    Dim i As Long
    
    For i = m_Entities.Count To 1 Step -1
        Set entity = m_Entities(i)
        If entity.EntityType = entityType Then
            entity.Destroy
            m_Entities.Remove i
        End If
    Next i
End Sub
