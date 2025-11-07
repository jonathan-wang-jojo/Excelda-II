'@Folder("Legacy.Entities")
Option Explicit

'═══════════════════════════════════════════════════════════════════════════════
' ENEMY CLASS (LEGACY - Migrate to Entity system)
'═══════════════════════════════════════════════════════════════════════════════

' Enemy identification
Public Name As String
Public EnemyType As String
Public SlotNumber As Long

' Animation frames
Public Frame1 As String
Public Frame2 As String
Public CurrentFrame As String

' Movement and behavior
Public Direction As String
Public Speed As Long
Public Behavior As String
Public InitialCount As Long
Public Count As Long
Public ChangeRotation As String

' Combat properties
Public Life As Long
Public Hit As Long
Public CollisionDamage As Long
Public ShootDamage As Long
Public ChargeDamage As Long
Public ChargeSpeed As Long

' Capabilities
Public CanCollide As String
Public CanShoot As String

' State flags
Public IsActive As Boolean
Public IsVisible As Boolean

'===================================================================================
'                              INITIALIZATION
'===================================================================================

Private Sub Class_Initialize()
    Reset
End Sub

Public Sub Reset()
    ' Clear all properties
    Name = ""
    EnemyType = ""
    SlotNumber = 0
    
    Frame1 = ""
    Frame2 = ""
    CurrentFrame = ""
    
    Direction = ""
    Speed = 0
    Behavior = ""
    InitialCount = 0
    Count = 0
    ChangeRotation = ""
    
    Life = 0
    Hit = 0
    CollisionDamage = 0
    ShootDamage = 0
    ChargeDamage = 0
    ChargeSpeed = 0
    
    CanCollide = ""
    CanShoot = ""
    
    IsActive = False
    IsVisible = False
End Sub

'===================================================================================
'                              CORE METHODS
'===================================================================================

Public Sub LoadFromData(enemyTypeName As String, slotNum As Long, dataRow As Long)
    ' Load enemy data from Data sheet - matches actual Excel structure
    On Error Resume Next
    
    SlotNumber = slotNum
    EnemyType = enemyTypeName
    
    ' Build frame names (e.g., "Octorok1F1")
    Frame1 = EnemyType & slotNum & "F1"
    Frame2 = EnemyType & slotNum & "F2"
    Name = Frame1
    CurrentFrame = Frame1
    
    ' Load from actual Excel columns (based on provided structure)
    With GameRegistryInstance().GetGameDataSheet()
        ' Column B: Name (display only)
        ' Column C: Visible? (Y/N)
        Life = Val(.Cells(dataRow, 4).Value)              ' Col D: Life
        Direction = Trim$(.Cells(dataRow, 6).Value)        ' Col F: Direction (often blank)
        Speed = Val(.Cells(dataRow, 7).Value)             ' Col G: Speed
        ' Column H: Frame Count (unused)
        InitialCount = Val(.Cells(dataRow, 9).Value)      ' Col I: Enemy Count (animation speed)
        Count = InitialCount
        Behavior = Trim$(.Cells(dataRow, 10).Value)        ' Col J: Behaviour
        ChangeRotation = Trim$(.Cells(dataRow, 11).Value)  ' Col K: Change Rotation?
        
        CollisionDamage = Val(.Cells(dataRow, 12).Value)  ' Col L: Collision damage
        CanShoot = Trim$(.Cells(dataRow, 13).Value)        ' Col M: Can shoot? (projectile name or N)
        ShootDamage = Val(.Cells(dataRow, 14).Value)      ' Col N: Shoot Damage
        ChargeSpeed = Val(.Cells(dataRow, 15).Value)      ' Col O: Can Charge (speed)
        ChargeDamage = Val(.Cells(dataRow, 16).Value)     ' Col P: Charge damage
        ' Column Q: Drop?
        ' Column R: bounceback
        
        ' Set collision flag
        CanCollide = IIf(CollisionDamage > 0, "Y", "N")
        
        ' Set initial direction if not specified
        If Direction = "" Then Direction = "S"
    End With
    
    IsActive = True
End Sub

Public Sub Show(cellAddress As String)
    ' Show enemy at specified location
    On Error Resume Next
    
    If Not IsActive Then Exit Sub
    
    With ActiveSheet.Shapes(Frame1)
        .Visible = True
        .Top = Range(cellAddress).Top
        .Left = Range(cellAddress).Left
        .Rotation = 0
    End With
    
    IsVisible = True
    CurrentFrame = Frame1
    Name = Frame1
End Sub

Public Sub Hide()
    ' Hide enemy and reset
    On Error Resume Next
    
    If Frame1 <> "" Then
        ActiveSheet.Shapes(Frame1).Visible = False
        ActiveSheet.Shapes(Frame1).Rotation = 0
    End If
    
    If Frame2 <> "" Then
        ActiveSheet.Shapes(Frame2).Visible = False
        ActiveSheet.Shapes(Frame2).Rotation = 0
    End If
    
    IsVisible = False
    Reset
End Sub

Public Sub SwapFrame()
    ' Swap between animation frames
    On Error Resume Next
    
    If Frame1 = "" Or Frame2 = "" Then Exit Sub
    
    If CurrentFrame = Frame1 Then
        ' Switch to frame 2
        ActiveSheet.Shapes(Frame2).Top = ActiveSheet.Shapes(Frame1).Top
        ActiveSheet.Shapes(Frame2).Left = ActiveSheet.Shapes(Frame1).Left
        ActiveSheet.Shapes(Frame2).Rotation = ActiveSheet.Shapes(Frame1).Rotation
        ActiveSheet.Shapes(Frame1).Visible = False
        ActiveSheet.Shapes(Frame2).Visible = True
        CurrentFrame = Frame2
        Name = Frame2
    Else
        ' Switch to frame 1
        ActiveSheet.Shapes(Frame1).Top = ActiveSheet.Shapes(Frame2).Top
        ActiveSheet.Shapes(Frame1).Left = ActiveSheet.Shapes(Frame2).Left
        ActiveSheet.Shapes(Frame1).Rotation = ActiveSheet.Shapes(Frame2).Rotation
        ActiveSheet.Shapes(Frame2).Visible = False
        ActiveSheet.Shapes(Frame1).Visible = True
        CurrentFrame = Frame1
        Name = Frame1
    End If
End Sub

Public Function GetShape() As Shape
    ' Get current frame shape
    On Error Resume Next
    Set GetShape = ActiveSheet.Shapes(Name)
End Function

Public Function GetAddress() As String
    ' Get current cell address
    On Error Resume Next
    Dim shp As Shape
    Set shp = GetShape()
    If Not shp Is Nothing Then
        GetAddress = shp.TopLeftCell.Address
    End If
End Function
