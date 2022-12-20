Attribute VB_Name = "DXMod"
Option Explicit
'
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal Lenght As Long)

Public Const pi As Single = 3.14159265358979   'can be worked out using (4*atn(1))

Public Dx As DirectX8 'The master Object, everything comes from here
Public D3D As Direct3D8 'This controls all things 3D
Public D3DDevice As Direct3DDevice8 'This actually represents the hardware doing the rendering
Public D3DX As D3DX8 '//A helper library

Public Type LitVertex
X As Single
Y As Single
Z As Single
Color As Long
tu As Single
TV As Single
End Type

Public Const LitVertexSize As Long = 24

'//The Descriptor for this vertex format...
Public Const Lit_FVF = (D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)

Public matWorld As D3DMATRIX '//How the vertices are positioned
Public matView As D3DMATRIX '//Where the camera is/where it's looking
Public matProj As D3DMATRIX '//How the camera projects the 3D world onto the 2D screen

Public ScreenSizeX As Long
Public ScreenSizeY As Long

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public StartTime As Currency
Public Frequency As Currency

Sub Main()
QueryPerformanceCounter StartTime
QueryPerformanceFrequency Frequency
Frequency = Frequency / 1000
Dim L As Long
L = InitialiseDX
InitializeFileMod
If L Then
MsgBox "DX8 Err Code:" & vbNewLine & CStr(L)
Else
InitializeBlocks
InitializeChunks
frmMain.Show
frmMain.MainLoop
End If
Terminate
TerminateFileMod
Unload frmMain
End
End Sub

Function TimerEx() As Currency
QueryPerformanceCounter TimerEx
TimerEx = (TimerEx - StartTime) / Frequency
End Function

Public Function InitialiseDX() As Long
On Error GoTo ErrHandler:

Dim DispMode As D3DDISPLAYMODE '//Describes our Display Mode
Dim D3DWindow As D3DPRESENT_PARAMETERS '//Describes our Viewport

Set Dx = New DirectX8  '//Create our Master Object
Set D3D = Dx.Direct3DCreate() '//Make our Master Object create the Direct3D Interface
Set D3DX = New D3DX8 '//Create our helper library...

'//We're going to use Fullscreen mode because I prefer it to windowed mode :)
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
'DispMode.Format = D3DFMT_X8R8G8B8
'DispMode.Format = D3DFMT_R5G6B5 'If this mode doesn't work try the commented one above...
'DispMode.Width = 1360
'DispMode.Height = 768
'DispMode.RefreshRate = 60

'DispMode.Format = DispMode.Format

ScreenSizeX = DispMode.Width
ScreenSizeY = DispMode.Height

D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
D3DWindow.BackBufferCount = 1 '//1 backbuffer only
D3DWindow.BackBufferFormat = DispMode.Format 'What we specified earlier
D3DWindow.BackBufferWidth = DispMode.Width
D3DWindow.BackBufferHeight = DispMode.Height
D3DWindow.hDeviceWindow = frmMain.hWnd
D3DWindow.EnableAutoDepthStencil = 1

If D3D.CheckDeviceFormat(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, D3DUSAGE_DEPTHSTENCIL, D3DRTYPE_SURFACE, D3DFMT_D16) = D3D_OK Then
D3DWindow.AutoDepthStencilFormat = D3DFMT_D16 '//16 bit Z-Buffer
End If

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, D3DWindow)

D3DDevice.SetVertexShader Lit_FVF

'//Transformed and lit vertices dont need lighting
'   so we disable it...
D3DDevice.SetRenderState D3DRS_LIGHTING, False

D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
'D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE

'//We need to enable our Z Buffer
D3DDevice.SetRenderState D3DRS_ZENABLE, 1

'#####################
'##  NEW STUFF - MATRICES ##
'####################

'//1. The World Matrix
'       This is applied to all vertices that are sent to be rendered
'       Values in here can alter the position of a vertex; common
'       things are rotation, Translation and scaling.

D3DXMatrixIdentity matWorld '//Make an identity matrix
            'the identity matrix, when applied to a vertex will not alter
            'it in any way - useful for making a fresh start on a persistant variable
D3DDevice.SetTransform D3DTS_WORLD, matWorld 'commit this matrix to the device


'//2. The View Matrix
'       This can be thought of as the camera; defined by a start point
'       and an end point, as well as an up vector. The up vector usually
'       points upwards along the Y axis (so +Y becomes up and -Y becomes down)
'       Should you wish Z to be the vertical axis you need to alter this vector

'D3DXMatrixLookAtLH matView, MakeVector(0, 5, 9), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
D3DXMatrixIdentity matView '//Make an identity matrix
        'Make the camera be at [0,4,0] and looking at the origin [0,0,0]
D3DDevice.SetTransform D3DTS_VIEW, matView

'//3. The projection Matrix
'       Once set up this can usually be left along; it defines what area of world-space
'       is visible to be rendered - the nearest and farthest reaches. It also specifies
'       how much we can see horizontally - a wide angle means we can see lots (like a landscape)
'       a small angle appears very zoomed in.
D3DXMatrixPerspectiveFovLH matProj, pi / 4, DispMode.Height / DispMode.Width, 0.1, 500
    'We're going to use a pi/4 view angle (it's in radians) - pi/2 or pi/3 are also used
    'We're giving it an aspect ration of 1:1
    'and we're telling it to only draw triangles greater 0.1 from the camera
    'and less than 500 from the camera (in meters)
D3DDevice.SetTransform D3DTS_PROJECTION, matProj

Exit Function
ErrHandler:
InitialiseDX = Err.Number
End Function

Public Sub Terminate()
On Error Resume Next
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing
End Sub

Public Function CreateLitVertex(X As Single, Y As Single, Z As Single, Diffuse As Long, tu As Single, TV As Single) As LitVertex
        CreateLitVertex.X = X
        CreateLitVertex.Y = Y
        CreateLitVertex.Z = Z
        CreateLitVertex.Color = Diffuse
        CreateLitVertex.tu = tu
        CreateLitVertex.TV = TV
End Function

Public Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.Z = Z
End Function
