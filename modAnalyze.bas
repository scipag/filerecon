Attribute VB_Name = "modAnalyze"
Option Explicit

Public Sub IdentifyFile(ByRef sFileContent As String)
    Dim sMagicDatabase() As String
    Dim iMagicDatabaseCount As Integer
    Dim i As Integer
    Dim sMagicDatabaseEntry() As String
    Dim lPositionStart As Long
    Dim lPositionEnd As Long
    
    frmMain.lstResults.Clear
    
    sMagicDatabase = Split(ReadFile(App.Path & "\magicdatabase.txt"), vbCrLf, , vbBinaryCompare)
    iMagicDatabaseCount = UBound(sMagicDatabase)
    
    'On Error Resume Next
    
    With frmMain
        For i = 0 To iMagicDatabaseCount
            If LenB(sMagicDatabase(i)) Then
                If Mid$(sMagicDatabase(i), 1, 1) <> "#" Then
                    sMagicDatabaseEntry = Split(sMagicDatabase(i), vbTab, , vbBinaryCompare)
                    
                    If Mid$(sMagicDatabaseEntry(0), 1, 1) = ">" Then
                        If InStrB(Mid$(sMagicDatabaseEntry(0), 2), sFileContent, sMagicDatabaseEntry(2), vbBinaryCompare) Then
                            .lstResults.AddItem sMagicDatabaseEntry(3)
                        End If
                    ElseIf InStrB(2, sMagicDatabaseEntry(0), "-", vbBinaryCompare) Then
                        lPositionStart = Mid$(sMagicDatabaseEntry(0), 1, InStr(2, sMagicDatabaseEntry(0), "-", vbBinaryCompare) - 1)
                        lPositionEnd = Mid$(sMagicDatabaseEntry(0), InStr(2, sMagicDatabaseEntry(0), "-", vbBinaryCompare) + 1)
            
                        If InStrB(1, Mid$(sFileContent, lPositionStart, lPositionEnd - lPositionStart + 1), sMagicDatabaseEntry(2), vbBinaryCompare) Then
                            .lstResults.AddItem sMagicDatabaseEntry(3)
                        End If
                    Else
                        If Mid$(sFileContent, sMagicDatabaseEntry(0), Len(sMagicDatabaseEntry(2))) = sMagicDatabaseEntry(2) Then
                            .lstResults.AddItem sMagicDatabaseEntry(3)
                        End If
                    End If
                End If
            End If
        Next i
        
        If (.lstResults.ListCount = 0) Then
            .lstResults.AddItem "Unknown file type"
            Call .ResetResultDetails
        Else
            .lstResults.ListIndex = 0
        End If
    End With
End Sub
