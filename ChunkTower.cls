VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChunkTower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Grid(15, 255, 15) As Block

Private Type Sector
DrawnFlag As Boolean
VBuffer As Direct3DVertexBuffer8
PrimitiveCount As Long
VertexCount As Long
IsBufferEmpty As Boolean
End Type

Dim VSs(15) As Sector

Dim StatePtr As Long
Dim State As ChunkState_E
Public WorldPosX As Long
Public WorldPosZ As Long

Public ChunkXP As ChunkTower
Public ChunkXN As ChunkTower
Public ChunkZP As ChunkTower
Public ChunkZN As ChunkTower

Property Let ChunkState(NewValue As ChunkState_E)
State = NewValue
If StatePtr Then
CopyMemory ByVal StatePtr, State, 4
End If
End Property

Property Get ChunkState() As ChunkState_E
ChunkState = State
End Property

Property Let StateOutputPtr(NewValue As Long)
StatePtr = NewValue
ChunkState = State
End Property

'X And -16 -> False = [0 >= X <= 15]
'          -> True  = [X < 0 Or X > 15]

Sub TriggerToRedraw(Optional ByVal ChunkToRedraw As Long = -1, Optional Now As Boolean = False)
If Now Then
Build ChunkToRedraw
Else
    If ChunkToRedraw And -16 Then
        For ChunkToRedraw = 0 To 15
        VSs(ChunkToRedraw).DrawnFlag = False
        Next
    Else
        VSs(ChunkToRedraw).DrawnFlag = False
    End If
End If
End Sub

Function Redraw(Optional TriggerRedraw As Boolean) As Boolean
On Error GoTo Err
Dim Z As Long
For Z = 0 To 15
    With VSs(Z)
        If Not .DrawnFlag And ChunkState >= ToBeBuilt Then
        Build Z
        End If
        If .DrawnFlag Then
        D3DDevice.SetStreamSource 0, .VBuffer, LitVertexSize
        D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, .PrimitiveCount
        End If
    End With
Next
Redraw = True
Err:
End Function

Sub GenerateTerrain()
FillBlocks 0, 127, 0, 15, 127, 15, 2
FillBlocks 0, 124, 0, 15, 126, 15, 3
FillBlocks 0, 0, 0, 15, 123, 15, 4
End Sub

Sub Build(Optional ChunkToBuild As Long = -1)
Dim Z As Long
If ChunkToBuild And -16 Then
    For Z = 0 To 15
    If Not VSs(Z).DrawnFlag Then Build Z
    Next
    Exit Sub
End If
'On Error GoTo Err
Const Color As Long = 16777215
Dim BlockId As BlockId
Dim TM() As Single
Dim Cube(35) As LitVertex
Dim Mesh() As LitVertex
Dim MaxM As Long
Dim SquareSize As Long
Dim X As Long, Y As Long

Dim OffX As Single, OffZ As Single
Dim XX0 As Single, YY0 As Single, ZZ0 As Single
Dim XX1 As Single, YY1 As Single, ZZ1 As Single
OffX = WorldPosX * 16
OffZ = WorldPosZ * 16

MaxM = -1
SquareSize = LitVertexSize * 6

For X = 0 To 15
    For Y = ChunkToBuild * 16 To ChunkToBuild * 16 + 15
        For Z = 0 To 15
            BlockId = BlockIds(Grid(X, Y, Z).B)
            If (BlockId.p And Invisible) = 0 Then
                XX0 = OffX + CSng(X):   XX1 = XX0 + 1
                YY0 = CSng(Y):          YY1 = YY0 + 1
                ZZ0 = OffZ + CSng(Z):   ZZ1 = ZZ0 + 1
                
                TM = BlockId.t
                If IsSolid(X + 1, Y, Z) Then
                MaxM = MaxM + 6
                ReDim Preserve Mesh(MaxM)
                Cube(0) = CreateLitVertex(XX1, YY1, ZZ0, Color, TM(TV.XPtu2), TM(TV.XPtv1)) 'X+
                Cube(2) = CreateLitVertex(XX1, YY1, ZZ1, Color, TM(TV.XPtu1), TM(TV.XPtv1))
                Cube(1) = CreateLitVertex(XX1, YY0, ZZ0, Color, TM(TV.XPtu2), TM(TV.XPtv2))
                Cube(5) = CreateLitVertex(XX1, YY0, ZZ1, Color, TM(TV.XPtu1), TM(TV.XPtv2))
                Cube(3) = Cube(2): Cube(4) = Cube(1)
                CopyMemory Mesh(MaxM - 5), Cube(0), SquareSize
                End If
                If IsSolid(X - 1, Y, Z) Then
                MaxM = MaxM + 6
                ReDim Preserve Mesh(MaxM)
                Cube(0) = CreateLitVertex(XX0, YY1, ZZ0, Color, TM(TV.XNtu1), TM(TV.XNtv1)) 'X-
                Cube(5) = CreateLitVertex(XX0, YY0, ZZ1, Color, TM(TV.XNtu2), TM(TV.XNtv2))
                Cube(2) = CreateLitVertex(XX0, YY0, ZZ0, Color, TM(TV.XNtu1), TM(TV.XNtv2))
                Cube(1) = CreateLitVertex(XX0, YY1, ZZ1, Color, TM(TV.XNtu2), TM(TV.XNtv1))
                Cube(4) = Cube(1): Cube(3) = Cube(2)
                CopyMemory Mesh(MaxM - 5), Cube(0), SquareSize
                End If
                If IsSolid(X, Y + 1, Z) Then
                MaxM = MaxM + 6
                ReDim Preserve Mesh(MaxM)
                Cube(0) = CreateLitVertex(XX0, YY1, ZZ1, Color, TM(TV.YPtu1), TM(TV.YPtv2)) 'Y+
                Cube(2) = CreateLitVertex(XX1, YY1, ZZ1, Color, TM(TV.YPtu2), TM(TV.YPtv2))
                Cube(1) = CreateLitVertex(XX0, YY1, ZZ0, Color, TM(TV.YPtu1), TM(TV.YPtv1))
                Cube(5) = CreateLitVertex(XX1, YY1, ZZ0, Color, TM(TV.YPtu2), TM(TV.YPtv1))
                Cube(3) = Cube(2): Cube(4) = Cube(1)
                CopyMemory Mesh(MaxM - 5), Cube(0), SquareSize
                End If
                If IsSolid(X, Y - 1, Z) Then
                MaxM = MaxM + 6
                ReDim Preserve Mesh(MaxM)
                Cube(0) = CreateLitVertex(XX0, YY0, ZZ1, Color, TM(TV.YNtu2), TM(TV.YNtv2)) 'Y-
                Cube(1) = CreateLitVertex(XX1, YY0, ZZ1, Color, TM(TV.YNtu1), TM(TV.YNtv2))
                Cube(2) = CreateLitVertex(XX0, YY0, ZZ0, Color, TM(TV.YNtu2), TM(TV.YNtv1))
                Cube(5) = CreateLitVertex(XX1, YY0, ZZ0, Color, TM(TV.YNtu1), TM(TV.YNtv1))
                Cube(4) = Cube(1): Cube(3) = Cube(2)
                CopyMemory Mesh(MaxM - 5), Cube(0), SquareSize
                End If
                If IsSolid(X, Y, Z + 1) Then
                MaxM = MaxM + 6
                ReDim Preserve Mesh(MaxM)
                Cube(0) = CreateLitVertex(XX0, YY1, ZZ1, Color, TM(TV.ZPtu1), TM(TV.ZPtv1)) 'Z+
                Cube(1) = CreateLitVertex(XX1, YY1, ZZ1, Color, TM(TV.ZPtu2), TM(TV.ZPtv1))
                Cube(2) = CreateLitVertex(XX0, YY0, ZZ1, Color, TM(TV.ZPtu1), TM(TV.ZPtv2))
                Cube(5) = CreateLitVertex(XX1, YY0, ZZ1, Color, TM(TV.ZPtu2), TM(TV.ZPtv2))
                Cube(3) = Cube(2): Cube(4) = Cube(1)
                CopyMemory Mesh(MaxM - 5), Cube(0), SquareSize
                End If
                If IsSolid(X, Y, Z - 1) Then
                MaxM = MaxM + 6
                ReDim Preserve Mesh(MaxM)
                Cube(0) = CreateLitVertex(XX0, YY1, ZZ0, Color, TM(TV.ZNtu2), TM(TV.ZNtv1)) 'Z-
                Cube(2) = CreateLitVertex(XX1, YY1, ZZ0, Color, TM(TV.ZNtu1), TM(TV.ZNtv1))
                Cube(1) = CreateLitVertex(XX0, YY0, ZZ0, Color, TM(TV.ZNtu2), TM(TV.ZNtv2))
                Cube(5) = CreateLitVertex(XX1, YY0, ZZ0, Color, TM(TV.ZNtu1), TM(TV.ZNtv2))
                Cube(3) = Cube(2): Cube(4) = Cube(1)
                CopyMemory Mesh(MaxM - 5), Cube(0), SquareSize
                End If
            End If
        Next
    Next
Next

With VSs(ChunkToBuild)
    .IsBufferEmpty = (MaxM = -1)
    .PrimitiveCount = (MaxM + 1) / 3
    .VertexCount = MaxM + 1
    .DrawnFlag = True
    If Not .IsBufferEmpty Then
        Set .VBuffer = D3DDevice.CreateVertexBuffer(LitVertexSize * .VertexCount, 0, Lit_FVF, D3DPOOL_DEFAULT)
        If .VBuffer Is Nothing Then Exit Sub
        D3DVertexBuffer8SetData .VBuffer, 0, LitVertexSize * .VertexCount, 0, Mesh(0)
    End If
End With

'For Z = 0 To 15
'If Not VSs(Z).DrawnFlag Then Exit Sub
'Next

Err:
End Sub

Private Function IsSolid(X As Long, Y As Long, Z As Long) As Boolean
On Error GoTo Err
IsSolid = ((BlockIds(Grid(X, Y, Z).B).p And Solid) = 0)
Exit Function
Err:
IsSolid = True
End Function

Sub SetBlock(X As Long, Y As Long, Z As Long, Block As Long, Optional Data As Long = 0)
On Error GoTo Err
Dim L(1) As Long
If Grid(X, Y, Z).B <> Block Or Grid(X, Y, Z).D <> Data Then
Grid(X, Y, Z).B = Block
Grid(X, Y, Z).D = Data
L(0) = ChunkPosByBlock(Y)
L(1) = WToC(Y)
    VSs(L(0)).DrawnFlag = False
    If L(1) = 0 And L(0) > 0 Then
    VSs(L(0) - 1).DrawnFlag = False
    ElseIf L(1) = 15 And L(0) < 15 Then
    VSs(L(0) + 1).DrawnFlag = False
    End If
End If
Err:
End Sub

Friend Function GetBlock(X As Long, Y As Long, Z As Long) As Block
On Error Resume Next
GetBlock = Grid(X, Y, Z)
End Function

Sub FillBlocks(X1 As Long, Y1 As Long, Z1 As Long, X2 As Long, Y2 As Long, Z2 As Long, Block As Long, Optional Data As Long = 0)
Dim X As Long
Dim Y As Long
Dim Z As Long
For X = X1 To X2
    For Y = Y1 To Y2
        For Z = Z1 To Z2
        Grid(X, Y, Z).B = Block
        Grid(X, Y, Z).D = Data
        Next
    Next
Next
End Sub

Private Sub Class_Initialize()
State = ToBeFilled
End Sub
