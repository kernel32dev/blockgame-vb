VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChunkBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim WithEvents Timer As SelfTimer
Attribute Timer.VB_VarHelpID = -1

Dim StatePtr As Long
Dim Chunk As ChunkTower
Dim State As ChunkState_E
'Dim Run As Boolean

Private Sub Class_Initialize()
Set Timer = New SelfTimer
Timer.Enabled = False
'Timer.Interval = 10 'Milliseconds in between attempts to find chunks
Timer.Interval = 1
End Sub

Property Let StateOutputPtr(NewValue As Long)
StatePtr = NewValue
End Property

'Property Let Enabled(NewValue As Boolean)
'Run = NewValue
'End Property

'Property Get Enabled() As Boolean
'Enabled = Run
'End Property

Friend Function AssingWork(Chnk As ChunkTower)
CopyMemory ByVal StatePtr, True, 2
State = Chnk.ChunkState
Chnk.ChunkState = Busy
Timer.Enabled = True
Set Chunk = Chnk
End Function

Private Sub Timer_Timer(ByVal Seconds As Currency) 'Main
Timer.Enabled = False

'Set Chunk = ChunkMod.FindNearestState(CenterX, CenterZ, ToBeFilled, MaxLoadDist)
'If Not Chunk Is Nothing Then

Chunk.GenerateTerrain
Chunk.Build
Chunk.ChunkState = Done

CopyMemory ByVal StatePtr, False, 2
End Sub
