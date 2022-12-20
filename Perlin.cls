VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Perlin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' vbAccelerator.com
' Perlin Noise
' Pretyy much optimal algorithm for calculation, converted to VB from
'// JAVA REFERENCE IMPLEMENTATION OF IMPROVED NOISE - COPYRIGHT 2002 KEN PERLIN.
' http://mrl.nyu.edu/~perlin/noise/
' See also http://www.noisemachine.com/
'
' Takes approx 30ms for a 100x100 grid with all optimisations on,
' using AMD Athlon 2k.
'
Private p() As Long
Private permutation() As Long
Dim cSeed As Long

Public Function G(ByVal X As Double, ByVal Y As Double, Optional ByVal Z As Double) As Double
Dim iX As Long
Dim iY As Long
Dim iZ As Long
Dim u As Double
Dim v As Double
Dim W As Double
Dim a As Long
Dim aa As Long
Dim ab As Long
Dim B As Long
Dim ba As Long
Dim bb As Long

   ' Find unit cube that contains point:
   iX = Int(X) And &HFF&
   iY = Int(Y) And &HFF&
   iZ = Int(Z) And &HFF&
   
   ' Find relative x,y,z of point in cube:
   X = X - Int(X)
   Y = Y - Int(Y)
   Z = Z - Int(Z)
   
   ' Compute fade curves for each of x,y,z
   u = fade(X)
   v = fade(Y)
   W = fade(Z)
   
   ' Hash coordinates of the 8 cube corners:
   a = p(iX) + iY
   aa = p(a) + iZ
   ab = p(a + 1) + iZ
   B = p(iX + 1) + iY
   ba = p(B) + iZ
   bb = p(B + 1) + iZ
   
   ' Add the blended results from 8 corners of the cube:
   G = lerp(W, lerp(v, lerp(u, grad(p(aa), X, Y, Z), _
                              grad(p(ba), X - 1, Y, Z)), _
                           lerp(u, grad(p(ab), X, Y - 1, Z), _
                              grad(p(bb), X - 1, Y - 1, Z))), _
                     lerp(v, lerp(u, grad(p(aa + 1), X, Y, Z - 1), _
                              grad(p(ba + 1), X - 1, Y, Z - 1)), _
                           lerp(u, grad(p(ab + 1), X, Y - 1, Z - 1), _
                              grad(p(bb + 1), X - 1, Y - 1, Z - 1))))
   
End Function

Public Function SmallStructure(X As Long, Z As Long, Chance As Double) As Boolean
Dim D As Double
Rnd -cSeed
Randomize (Sqr(X) - 112.34) * (Sqr(Z) + 3.1415) * Chance
D = Rnd
SmallStructure = (D <= Chance)
End Function

Public Function O(ByVal X As Double, ByVal Y As Double, Optional BaseMultiplier As Double = 1) As Double
Dim Z As Long
Dim W As Long
Rnd -Seed
Z = 2
For W = 0 To 7
O = O + G((X + Rnd * 1000) * Z * BaseMultiplier, (Y + Rnd * 1000) * Z * BaseMultiplier) / Z
Z = Z * 2
Next
End Function

Private Function fade(ByVal t As Double) As Double
   fade = t * t * t * (t * (t * 6# - 15#) + 10#)
End Function
Private Function lerp(ByVal t As Double, ByVal a As Double, ByVal B As Double) As Double
   lerp = a + t * (B - a)
End Function
Private Function grad(ByVal hash As Long, ByVal X As Double, ByVal Y As Double, ByVal Z As Double) As Double
Dim h As Long
Dim u As Double
Dim v As Double
Dim G As Double

   '// CONVERT LO 4 BITS OF HASH CODE
   h = hash And &HF&
   
   '// INTO 12 GRADIENT DIRECTIONS.
   If (h < 8) Or (h = 12) Or (h = 13) Then
      u = X
   Else
      u = Y
   End If
   
   If (h < 4) Or (h = 12) Or (h = 13) Then
      v = Y
   Else
      Y = Z
   End If
   
   If ((h And 1) = 0) Then
      G = u
   Else
      G = -u
   End If
   
   If ((h And 2) = 0) Then
      G = G + v
   Else
      G = G - v
   End If
   
   grad = G
   
End Function

Property Get Seed() As Long
Seed = cSeed
End Property

Property Let Seed(NewValue As Long)
cSeed = NewValue
Rnd -cSeed
 
   Dim i As Long
   For i = 0 To 255
      p(i) = Abs(Rnd) * 255
      p(256 + i) = p(i)
   Next i
End Property

Private Sub Class_Initialize()
ReDim p(511) As Long
ReDim permutation(255) As Long
Seed = 0
End Sub
