Attribute VB_Name = "modReadfile"
Option Explicit

Private Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
Private Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long

Function ReadFile(ByRef sFileName As String) As String
    Dim hFile As Long
    Dim sFileContent As String
    Dim lFileLength As Long
    
    If (Dir$(sFileName, 16) <> "") Then
        hFile = lopen(sFileName, 0)
        lFileLength = llseek(hFile, 0, 2)
        llseek hFile, 0, 0
        
        sFileContent = Space$(lFileLength)
        lread hFile, ByVal sFileContent, lFileLength
        
        lclose hFile
        ReadFile = sFileContent
    Else
        MsgBox "File " & sFileName & " not accessible.", vbExclamation + vbOKOnly, "Open file error"
    End If
End Function
