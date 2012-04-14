Attribute VB_Name = "ModDebug"
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)


Public Sub DebugString(DString As String)
OutputDebugString DString
End Sub
