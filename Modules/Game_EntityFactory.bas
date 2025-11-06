'@Folder("Entity.Core")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' ENTITY FACTORY
'═══════════════════════════════════════════════════════════════════════════════
'                              ENEMIES
'===================================================================================

' Spawn classic Zelda Octorok enemy
Public Function SpawnOctorok(x As Double, y As Double) As Entity
    ' Create entity with Transform + Sprite (automatic)
    Dim octorok As Entity
    Set octorok = EntityManagerInstance.CreateEntity("Octorok", x, y)
    
    ' Add components via fluent API
    octorok.WithHealth(3) _
           .WithCollision(8) _
           .WithBehavior(New OctorokBehavior)
    
    ' Setup sprite (TODO: Replace with actual sprite name from Excel)
    ' For now, using placeholder - replace "Octorok1F1" with actual shape name
    octorok.Sprite.LoadSprite "Octorok1F1"
    octorok.Transform.Speed = 2
    
    Set SpawnOctorok = octorok
End Function

' TODO: Add more enemy spawners (Moblin, Tektite, etc.)

'===================================================================================
'                              NPCs (Non-combat entities)
'===================================================================================

' TODO: SpawnOldMan(x, y, dialogue)
' TODO: SpawnShopkeeper(x, y)

'===================================================================================
'                              OBJECTS
'===================================================================================

' TODO: SpawnPot(x, y)
' TODO: SpawnChest(x, y, itemID)

'===================================================================================
'                              PICKUPS
'===================================================================================

' TODO: SpawnHeart(x, y)
' TODO: SpawnRupee(x, y, value)
' TODO: SpawnKey(x, y)

'===================================================================================
'                              TESTING
'===================================================================================

' Test function - spawn an Octorok in current room
Public Sub TestSpawnOctorok()
    Dim octorok As Entity
    Set octorok = SpawnOctorok(200, 200)
    Debug.Print "Spawned Octorok at (200, 200) - EntityID: " & octorok.EntityID
End Sub
