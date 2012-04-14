Attribute VB_Name = "ModDebug"
Private Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)


Public Sub DebugString(DString As String)
OutputDebugString DString
End Sub

//大或高或低后发酵的是发拉三发