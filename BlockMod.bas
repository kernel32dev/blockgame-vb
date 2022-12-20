Attribute VB_Name = "BlockMod"
Option Explicit

Public Texture As Direct3DTexture8

Type Block
B As Long
D As Long
End Type

Enum TV 'Textures Vertecies
XPtu1 = 0
XPtv1 = 1
XPtu2 = 2
XPtv2 = 3
XNtu1 = 4
XNtv1 = 5
XNtu2 = 6
XNtv2 = 7
YPtu1 = 8
YPtv1 = 9
YPtu2 = 10
YPtv2 = 11
YNtu1 = 12
YNtv1 = 13
YNtu2 = 14
YNtv2 = 15
ZPtu1 = 16
ZPtv1 = 17
ZPtu2 = 18
ZPtv2 = 19
ZNtu1 = 20
ZNtv1 = 21
ZNtu2 = 22
ZNtv2 = 23
End Enum

Enum BlockProperty
Invisible = 1
Trasparent = 0 'Hides Neighbour Blocks Sides
Solid = 2
Collide = 4 'Colision
NoCollide = 0
End Enum

Type BlockId
Name As String
p As BlockProperty
t(23) As Single
End Type

Public BlockIds() As BlockId
Public MaxBIds As Long

Const ResourceSizeX As Single = 16
Const ResourceSizeY As Single = 16
Const ResourceStepX As Single = 1 / ResourceSizeX
Const ResourceStepY As Single = 1 / ResourceSizeY

Sub InitializeBlocks()
Dim NoTexture(23) As Single
MaxBIds = -1
Set Texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\Resource.bmp")
AddBlockId "air", Invisible, NoTexture
'AddBlockId "grass", Solid Or Collide, CreateOneSideTexture(12, 8)
'AddBlockId "grass", Solid Or Collide, CreateFullTexture(12, 8, 12, 8, 9, 7, 12, 7, 12, 8, 12, 8)
'AddBlockId "grass", Solid Or Collide, CreateFullTexture(8, 10, 9, 10, 10, 10, 11, 10, 12, 10, 13, 10)
AddBlockId "user", Solid Or Collide, CreateFullTexture(11, 3, 11, 3, 11, 2, 4, 0, 12, 3, 12, 3)
AddBlockId "grass", Solid Or Collide, CreateFullTexture(3, 0, 3, 0, 0, 0, 2, 0, 3, 0, 3, 0)
AddBlockId "dirt", Solid Or Collide, CreateFullTexture(2, 0, 2, 0, 2, 0, 2, 0, 2, 0, 2, 0)
AddBlockId "stone", Solid Or Collide, CreateFullTexture(1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0)

End Sub

Private Sub AddBlockId(Name As String, p As BlockProperty, Texture() As Single)
MaxBIds = MaxBIds + 1
Dim Z As Long
ReDim Preserve BlockIds(MaxBIds)
With BlockIds(MaxBIds)
.Name = Name
.p = p
    For Z = 0 To 23
    .t(Z) = Texture(Z)
    Next
End With
End Sub

Private Function CreateOneSideTexture(X As Long, Y As Long) As Single()
Dim t(23) As Single, Z As Long
For Z = 0 To 20 Step 4
t(Z) = X / ResourceSizeX
t(Z + 1) = Y / ResourceSizeY
t(Z + 2) = t(Z) + ResourceStepX
t(Z + 3) = t(Z + 1) + ResourceStepY
Next
CreateOneSideTexture = t
End Function

Private Function CreateFullTexture( _
XP_x As Long, XP_y As Long, _
XN_x As Long, XN_y As Long, _
YP_x As Long, YP_y As Long, _
YN_x As Long, YN_y As Long, _
ZP_x As Long, ZP_y As Long, _
ZN_x As Long, ZN_y As Long) As Single()
Dim t(23) As Single, Z As Long
t(0) = XP_x / ResourceSizeX
t(1) = XP_y / ResourceSizeY
t(4) = XN_x / ResourceSizeX
t(5) = XN_y / ResourceSizeY
t(8) = YP_x / ResourceSizeX
t(9) = YP_y / ResourceSizeY
t(12) = YN_x / ResourceSizeX
t(13) = YN_y / ResourceSizeY
t(16) = ZP_x / ResourceSizeX
t(17) = ZP_y / ResourceSizeY
t(20) = ZN_x / ResourceSizeX
t(21) = ZN_y / ResourceSizeY
For Z = 0 To 20 Step 4
t(Z + 2) = t(Z) + ResourceStepX
t(Z + 3) = t(Z + 1) + ResourceStepY
Next
CreateFullTexture = t
End Function

