Attribute VB_Name = "ModMain"
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Main()
    If App.PrevInstance = True Then
        MessageBox 0&, "不允许同时运行本程序的两个实例！", "打地鼠", vbOKOnly + vbSystemModal
        Dim TempTitle As String
        TempTitle = App.Title
        App.Title = ""
        AppActivate TempTitle
        TerminateApp
    End If
    InitCommonControls
    Load FrmSplash
    FrmSplash.Refresh
    FrmSplash.SetLoadingStatus "Loading..."
    FrmSplash.SetLoadingStatus "Loading Main Window..."
    Load FrmMain
    FrmMain.Show
    FrmMain.Refresh
    FrmSplash.Hide
    Unload FrmSplash
End Sub

Public Sub TerminateApp()
    DebugString "Application Quit.Date:" & Date & " Time:" & Time & " Timer: " & Timer
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
    End
End Sub
