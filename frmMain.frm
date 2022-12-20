VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Graft"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PointL
X As Long
Y As Long
End Type

Private Type PointS
X As Long
Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointL) As Long
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim FPS_LastCheck As Long
Dim FPS_Count As Long
Dim FPS_Highest As Integer

Public EndLoop As Boolean

Dim CamX As Single
Dim CamY As Single
Dim PlayerX As Single
Dim PlayerY As Single
Dim PlayerZ As Single

Dim CB As Long

Enum Keys_E
Up = 1
Down = 2
Left = 4
Right = 8
Space = 16
Shift = 32
Ctrl = 64
End Enum

Dim Keys As Keys_E

Dim StopWatchTimer As Currency
Dim StopWatchTimerRes As Currency

Private Sub StartTimer()
QueryPerformanceCounter StopWatchTimer
End Sub

Private Function EndTimer() As Currency
If StopWatchTimer Then
QueryPerformanceCounter EndTimer
EndTimer = (EndTimer - StopWatchTimer) / Frequency
StopWatchTimerRes = EndTimer
StopWatchTimer = 0
Else
EndTimer = StopWatchTimerRes
End If
End Function

Public Function Render() As Boolean
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, RGB(254, 198, 177), 1#, 0
D3DDevice.SetTransform D3DTS_VIEW, matView
D3DDevice.BeginScene
D3DDevice.SetTexture 0, Texture
Render = ChunkMod.RenderAll
D3DDevice.EndScene
On Error Resume Next
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function

Private Sub Form_DblClick()
Form_MouseDown (CB), 0, ScreenSizeX / 2, ScreenSizeY / 2
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Debug.Print "Up:", KeyCode, Shift
If KeyCode = 87 Then Keys = Keys And Not Up
If KeyCode = 83 Then Keys = Keys And Not Down
If KeyCode = 68 Then Keys = Keys And Not Left
If KeyCode = 65 Then Keys = Keys And Not Right
If KeyCode = 32 Then Keys = Keys And Not Space
If KeyCode = 16 Then Keys = Keys And Not 32 'Shift
If KeyCode = 17 Then Keys = Keys And Not Ctrl
End Sub

Function RayTraceAngle(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
AngleX As Single, AngleY As Single, StepSize As Single, _
MaxSteps As Long, OutX As Long, OutY As Long, OutZ As Long, Optional ReturnPrevious As Boolean) As Boolean

RayTraceAngle = RayTraceStep(X, Y, Z, _
StepSize * Sin(AngleX * (pi / 180)) * Abs(Cos(AngleY * (pi / 180))), _
StepSize * Sin(AngleY * (pi / 180)), _
StepSize * Cos(AngleX * (pi / 180)) * Abs(Cos(AngleY * (pi / 180))), _
MaxSteps, OutX, OutY, OutZ, ReturnPrevious)

End Function

Function RayTraceStep(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
StepX As Single, StepY As Single, StepZ As Single, _
MaxSteps As Long, OutX As Long, OutY As Long, OutZ As Long, Optional ReturnPrevious As Boolean) As Boolean
Dim W As Long
Dim OldBlock(2) As Long
OldBlock(0) = NToB(X) + 1
OldBlock(1) = NToB(Y) + 1
OldBlock(2) = NToB(Z) + 1
For W = 1 To MaxSteps
    OutX = NToB(X)
    OutY = NToB(Y)
    OutZ = NToB(Z)
    If OldBlock(0) <> OutX Or OldBlock(1) <> OutY Or OldBlock(2) <> OutZ Then
        OldBlock(0) = OutX
        OldBlock(1) = OutY
        OldBlock(2) = OutZ
        If (BlockIds(GetBlock(OutX, OutY, OutZ).B).p And Invisible) = 0 Then
        RayTraceStep = True
            If ReturnPrevious Then
            X = X - StepX
            Y = Y - StepY
            Z = Z - StepZ
            OutX = NToB(X)
            OutY = NToB(Y)
            OutZ = NToB(Z)
            End If
        Exit Function
        End If
    End If
    X = X + StepX
    Y = Y + StepY
    Z = Z + StepZ
Next
End Function

Sub MainLoop()
Dim RotateAngle As Single
Dim matTemp As D3DMATRIX
Dim PlayerSpeed As Single
Dim DoDebug As Boolean
Dim GameTick As Long
Dim Times(3) As Currency
PlayerX = 16
PlayerY = 130
PlayerZ = 16
CamY = -45
CamX = 105
PlayerSpeed = 0.05
MaxLoadDist = 3
Do
    Times(3) = TimerEx 'Tick
    
    StartTimer
    CenterX = ChunkPosBySingle(PlayerX)
    CenterZ = ChunkPosBySingle(PlayerZ)
    TickChunks
    Times(0) = EndTimer
    
    StartTimer
    If Not Render Then 'RENDER
    MsgBox "Error during rendering", vbCritical, "/!\"
    Exit Do
    End If
    Times(1) = EndTimer
    
    If GetTickCount() - FPS_LastCheck >= 1000 Then
        If FPS_Count >= FPS_Highest Then FPS_Highest = FPS_Count
        FPS_Count = 0
        FPS_LastCheck = GetTickCount()
    End If
    FPS_Count = FPS_Count + 1
    
    StartTimer
    DoEvents
    Times(2) = EndTimer
    
    If Keys Then
        If Keys And Up Then 'Foward
        PlayerX = PlayerX + Sin(CamX * (pi / 180)) * PlayerSpeed
        PlayerZ = PlayerZ + Cos(CamX * (pi / 180)) * PlayerSpeed
        End If
        If Keys And Down Then 'Backward
        PlayerX = PlayerX - Sin(CamX * (pi / 180)) * PlayerSpeed
        PlayerZ = PlayerZ - Cos(CamX * (pi / 180)) * PlayerSpeed
        End If
        If Keys And Left Then 'Left
        PlayerX = PlayerX + Cos(CamX * (pi / 180)) * PlayerSpeed
        PlayerZ = PlayerZ - Sin(CamX * (pi / 180)) * PlayerSpeed
        End If
        If Keys And Right Then 'right
        PlayerX = PlayerX - Cos(CamX * (pi / 180)) * PlayerSpeed
        PlayerZ = PlayerZ + Sin(CamX * (pi / 180)) * PlayerSpeed
        End If
        If Keys And Space Then 'Up
        PlayerY = PlayerY + PlayerSpeed
        End If
        If Keys And Shift Then 'Down
        PlayerY = PlayerY - PlayerSpeed
        End If
    UpdateCamera
    End If
    
    Times(3) = TimerEx - Times(3)
    
    If GameTick Mod 50 = 0 Then
    FilePrintLn "Time       = " & TimerEx & "ms"
    FilePrintLn "TickChunks = " & Times(0) & "ms"
    FilePrintLn "Rendering  = " & Times(1) & "ms"
    FilePrintLn "DoEvents   = " & Times(2) & "ms" & vbNewLine
    FilePrintLn "TickTime   = " & Times(3) & "ms" & vbNewLine
    End If
    GameTick = GameTick + 1
Loop Until EndLoop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Debug.Print Chr(KeyCode); KeyCode
Dim B As Boolean
'a=65 '37
'd=68 '40
'w=87 '38
's=83 '39
If KeyCode = 27 Then EndLoop = True
If KeyCode = 87 Then Keys = Keys Or Up
If KeyCode = 83 Then Keys = Keys Or Down
If KeyCode = 68 Then Keys = Keys Or Left
If KeyCode = 65 Then Keys = Keys Or Right
If KeyCode = 32 Then Keys = Keys Or Space
If KeyCode = 16 Then Keys = Keys Or 32 'Shift
If KeyCode = 17 Then Keys = Keys Or Ctrl
If KeyCode = 75 Then
PlayerX = PlayerX + 1 'K
End If
If KeyCode = 76 Then
PlayerZ = PlayerZ + 1 'L
End If

'If KeyCode = 109 And Brush > 0 Then Brush = Brush - 1
'If KeyCode = 107 And Brush < MaxB Then Brush = Brush + 1

'If KeyCode >= 48 And KeyCode <= 57 Then
'Brush = KeyCode - 48
'If Brush = 0 Then Brush = 10
'UpdateFormCaption
'End If

'Form1.Caption = "X:" & CamX & " Y:" & CamY & " Z:" & PlayerZ & " CamXD" & CamXD & " CamYD" & CamYD
End Sub

Sub UpdateCamera()
D3DXMatrixLookAtLH matView, MakeVector(CSng(PlayerX), CSng(PlayerY), CSng(PlayerZ)), MakeVector( _
CSng(PlayerX) + 3 * Sin(CamX * (pi / 180)) * Abs(Cos(CamY * (pi / 180))), _
CSng(PlayerY) + 3 * Sin(CamY * (pi / 180)), _
CSng(PlayerZ) + 3 * Cos(CamX * (pi / 180)) * Abs(Cos(CamY * (pi / 180)))), _
MakeVector(0, 1, 0)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CB = Button
If Button = 1 Or Button = 2 Then
Dim XX As Long
Dim YY As Long
Dim ZZ As Long
    If RayTraceAngle(PlayerX, PlayerY, PlayerZ, CamX, CamY, 0.05, 1000, XX, YY, ZZ, Button = 2) Then
        SetBlock XX, YY, ZZ, Button - 1
    End If
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim p As PointL
GetCursorPos p
CamX = CamX + (p.X - (ScreenSizeX / 2)) / 9.7
CamY = CamY - (p.Y - (ScreenSizeY / 2)) / 9.7
If CamY >= 90 Then CamY = 89.9
If CamY <= -90 Then CamY = -89.9
If CamX > 360 Then CamX = CamX - 360
If CamX < 0 Then CamX = CamX + 360
SetCursorPos ScreenSizeX / 2, ScreenSizeY / 2
GetCursorPos p
UpdateCamera
End Sub
