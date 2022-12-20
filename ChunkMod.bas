Attribute VB_Name = "ChunkMod"
Option Explicit

Type WorldPos
X As Long
Z As Long
End Type

Enum ChunkState_E
Busy = -2
Dead = -1
ToBeFilled = 0
ToBeBuilt = 1
Done = 2
ToBeKilled = 4
End Enum

Dim ChunkWorldPos() As WorldPos
Dim States() As ChunkState_E
Dim Chunks() As ChunkTower
Dim MaxC As Long

Public CenterX As Long
Public CenterZ As Long
Public MaxLoadDist As Long

Dim WorkersWork As Long
Dim WorkerState As Boolean 'True = Working 'False = Available
Dim Worker As ChunkBuilder

Function RenderAll() As Boolean
Dim W As Long
RenderAll = True
For W = 0 To MaxC
    If States(W) = Done Then
    RenderAll = RenderAll And Chunks(W).Redraw
    End If
Next
End Function

Sub TickChunks()
Dim C As ChunkTower
Dim WP As WorldPos
If Not WorkerState Then
    If FindNearestFreeSpot(CenterX, CenterZ, WP, MaxLoadDist) Then
    Set C = CreateChunk(WP.X, WP.Z)
    If Not C Is Nothing Then Worker.AssingWork C
    End If
End If
End Sub

Function IsChunkUnlocked(ChunkObjPtr As Long) As Boolean
IsChunkUnlocked = (ChunkObjPtr <> WorkersWork)
End Function

Sub InitializeChunks()
MaxC = -1
MaxLoadDist = 6
Set Worker = CreateObject("Graft.ChunkBuilder")
Worker.StateOutputPtr = VarPtr(WorkerState)
End Sub

Function FindNearestState(PosX As Long, PosZ As Long, State As ChunkState_E, Optional MaxDist As Long = -1) As ChunkTower
Dim W As Long
Dim DistC As Long
Dim DistB As Long
If MaxDist < 0 Then
DistC = -1
Else
DistC = MaxDist * MaxDist
End If
For W = 0 To MaxC
    If States(W) = State Then
        DistB = (ChunkWorldPos(W).X - PosX) ^ 2 + (ChunkWorldPos(W).Z - PosZ) ^ 2
        If DistB <= DistC Or DistC = -1 Then
        Set FindNearestState = Chunks(W)
        DistC = DistB
        If DistC = 0 Then Exit Function
        End If
    End If
Next
End Function

Private Function FindNearestFreeSpot(PosX As Long, PosZ As Long, Output As WorldPos, Optional MaxDist As Long = -1) As Boolean
Dim X As Long
Dim Z As Long
Dim W As Long
Dim DistC As Long
Dim DistB As Long
If MaxDist < 0 Then
DistB = -1
Else
DistB = MaxDist * MaxDist
End If
For DistC = 0 To MaxDist
    For X = PosX - DistC To PosX + DistC
        For Z = PosX - DistC To PosZ + DistC
            If DistB = -1 Then GoTo Skip
            If (X - PosX) ^ 2 + (Z - PosZ) ^ 2 <= DistB * DistB Then
Skip:
                For W = 0 To MaxC
                If ChunkWorldPos(W).X = X And ChunkWorldPos(W).Z = Z And States(W) <> Dead Then Exit For
                Next
                If W > MaxC Then 'Success
                    Output.X = X
                    Output.Z = Z
                    FindNearestFreeSpot = True
                    Exit Function
                End If
            End If
        Next
    Next
Next
End Function

Sub UpdateSideChunks(Chunk As ChunkTower)

End Sub

Function CreateChunk(X As Long, Z As Long, Optional ReturnNothingOnFail As Boolean = True) As ChunkTower
Dim W As Long
For W = 0 To MaxC
    If ChunkWorldPos(W).X = X Then
        If ChunkWorldPos(W).Z = Z Then
            If States(W) = Dead Then
                Exit For
            ElseIf ReturnNothingOnFail Or States(W) = Busy Then
                Exit Function
            Else
                Set CreateChunk = Chunks(W)
                Exit Function
            End If
        Exit Function
        End If
    End If
Next
If W > MaxC Then
MaxC = MaxC + 1
ReDim Preserve ChunkWorldPos(MaxC)
ReDim Preserve States(MaxC)
ReDim Preserve Chunks(MaxC)
W = MaxC
End If
Set Chunks(W) = New ChunkTower
Chunks(W).StateOutputPtr = VarPtr(States(W))
Chunks(W).WorldPosX = X
Chunks(W).WorldPosZ = Z
Chunks(W).ChunkState = ToBeFilled
ChunkWorldPos(W).X = X
ChunkWorldPos(W).Z = Z
'Chunks(MaxC).GenerateTerrain
'Chunks(MaxC).Build
UpdateSideChunks Chunks(MaxC)
Set CreateChunk = Chunks(MaxC)
End Function

Function FindChunk(X As Long, Z As Long) As ChunkTower
Dim W As Long
For W = 0 To MaxC
    If States(W) <> Dead Then
        If ChunkWorldPos(W).X = X Then
            If ChunkWorldPos(W).Z = Z Then
            Set FindChunk = Chunks(W)
            Exit Function
            End If
        End If
    End If
Next
End Function

Function NToB(W As Single) As Long
NToB = Int(W) ' + (W < 0)
End Function

Function WToC(W As Long) As Long
WToC = (W - (W < 0)) Mod 16 - (W < 0) * 15
End Function

Function ChunkPosByBlock(W As Long) As Long
ChunkPosByBlock = (W - (W < 0)) \ 16 + (W < 0)
End Function

Function ChunkPosBySingle(W As Single) As Long
ChunkPosBySingle = ChunkPosByBlock(Int(W))
End Function

Function GetBlock(X As Long, Y As Long, Z As Long) As Block
On Error Resume Next
GetBlock = FindChunk(ChunkPosByBlock(X), ChunkPosByBlock(Z)).GetBlock(WToC(X), Y, WToC(Z))
End Function

Sub SetBlock(X As Long, Y As Long, Z As Long, Block As Long, Optional Data As Long = 0)
On Error Resume Next
FindChunk(ChunkPosByBlock(X), ChunkPosByBlock(Z)).SetBlock WToC(X), Y, WToC(Z), Block, Data
End Sub
