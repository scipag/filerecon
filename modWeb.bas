Attribute VB_Name = "modWeb"
Option Explicit

Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub OpenProjectWebsite()
    Call ShellExecute(frmMain.hwnd, "Open", APP_WEBSITE, "", App.Path, 1)
End Sub

Public Sub OpenUpdateWebsite()
    Call ShellExecute(frmMain.hwnd, "Open", APP_WEBSITE & "?s=download&v=" & APP_NAME, "", App.Path, 1)
End Sub

