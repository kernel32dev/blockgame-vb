Attribute VB_Name = "DebugFileMod"
Option Explicit

Dim F As Long

Sub FilePrint(Text As String)
Put F, , Text
End Sub

Sub FilePrintLn(Text As String)
Put F, , Text & vbNewLine
End Sub

Sub InitializeFileMod()
F = FreeFile
Open App.Path & "\Debug.txt" For Output As F
Close F
F = FreeFile
Open App.Path & "\Debug.txt" For Binary As F
End Sub

Sub TerminateFileMod()
Close F
End Sub
