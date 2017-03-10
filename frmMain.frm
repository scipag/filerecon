VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "filerecon"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4455
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
      Begin VB.TextBox txtDescription 
         Height          =   615
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   256
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtAccuracy 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   13
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txtLength 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   32
         TabIndex        =   12
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtPattern 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   128
         TabIndex        =   11
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtType 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtPosition 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Accuracy"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Length"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Pattern"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Type"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Position"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog cdgFileOpen 
      Left            =   3840
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open File For Analysis"
      Filter          =   "All Files|*.*"
   End
   Begin VB.ListBox lstResults 
      Height          =   840
      ItemData        =   "frmMain.frx":08CA
      Left            =   120
      List            =   "frmMain.frx":08CC
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   735
      Left            =   3600
      Picture         =   "frmMain.frx":08CE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Open and analyze a file..."
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExitItem 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAboutItem 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpUpdatesItem 
         Caption         =   "&Check for Updates..."
      End
      Begin VB.Menu mnuHelpWebsiteItem 
         Caption         =   "&filerecon Home Page..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpen_Click()
    Call OpenDialog
End Sub

Public Sub OpenDialog()
    Dim sFileName As String
    
    cdgFileOpen.InitDir = App.Path
    
    On Error GoTo Cancel
    
    cdgFileOpen.ShowOpen
    sFileName = cdgFileOpen.FileName
    
    If LenB(sFileName) Then
        If (Dir$(sFileName) <> "") Then
            Call IdentifyFile(ReadFile(sFileName))
            Me.Caption = APP_NAME & " - " & sFileName
            Exit Sub
        End If
    End If

Cancel:
End Sub

Private Sub Form_Load()
    Me.Caption = APP_NAME
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        Me.Height = 4680
        Me.Width = 4575
    End If
End Sub

Private Sub ShowHitDetails(ByRef sSearch As String)
    Dim sMagicDatabase() As String
    Dim iMagicDatabaseCount As Integer
    Dim sMagicDatabaseEntry() As String
    Dim i As Integer
    Dim iMagicDatabaseEntryStringLength As Integer
    Dim sMagicDatabaseEntryPosition As String
    Dim sMagicDatabaseEntryPositionType As String
    Dim sMagicDatabaseEntryType As String
    Dim iMagicDatabaseEntryType As Integer
    Dim sMagicDatabaseEntryTypeDetail As String
    Dim sMagicDatabaseEntryString As String
    Dim sMagicDatabaseEntryDescription As String
    Dim sMagicDatabaseEntryAccuracy As String
    
    sMagicDatabase = Split(ReadFile(App.Path & "\magicdatabase.txt"), vbCrLf, , vbBinaryCompare)
    iMagicDatabaseCount = UBound(sMagicDatabase)
    
    'On Error Resume Next
    
    For i = 0 To iMagicDatabaseCount
        If LenB(sMagicDatabase(i)) Then
            If Mid$(sMagicDatabase(i), 1, 1) <> "#" Then
                sMagicDatabaseEntry = Split(sMagicDatabase(i), vbTab, , vbBinaryCompare)
                
                If (UBound(sMagicDatabaseEntry) = 3) Then
                    If (sMagicDatabaseEntry(3) = sSearch) Then
                        sMagicDatabaseEntryPosition = sMagicDatabaseEntry(0)
                        sMagicDatabaseEntryType = sMagicDatabaseEntry(1)
                        sMagicDatabaseEntryString = sMagicDatabaseEntry(2)
                        sMagicDatabaseEntryDescription = sMagicDatabaseEntry(3)
                        
                        iMagicDatabaseEntryStringLength = Len(sMagicDatabaseEntryString)
                        
                        If Mid$(sMagicDatabaseEntryPosition, 1, 1) = ">" Then
                            iMagicDatabaseEntryType = 3
                            sMagicDatabaseEntryPositionType = "undefined range"
                        ElseIf InStrB(2, sMagicDatabaseEntryPosition, "-", vbBinaryCompare) Then
                            iMagicDatabaseEntryType = 2
                            sMagicDatabaseEntryPositionType = "defined range"
                        Else
                            iMagicDatabaseEntryType = 1
                            sMagicDatabaseEntryPositionType = "exact position"
                            sMagicDatabaseEntryPosition = sMagicDatabaseEntryPosition & "-" & sMagicDatabaseEntryPosition + iMagicDatabaseEntryStringLength
                        End If
                        
                        If (sMagicDatabaseEntryType = "s") Then
                            sMagicDatabaseEntryTypeDetail = "string"
                        Else
                            sMagicDatabaseEntryTypeDetail = "unknown"
                        End If
                        
                        If (iMagicDatabaseEntryStringLength > 3) Then
                            If (iMagicDatabaseEntryType = 1) Then
                                sMagicDatabaseEntryAccuracy = "high"
                            Else
                                sMagicDatabaseEntryAccuracy = "medium"
                            End If
                        ElseIf (iMagicDatabaseEntryStringLength > 1) Then
                            If (iMagicDatabaseEntryType = 1) Then
                                sMagicDatabaseEntryAccuracy = "medium"
                            Else
                                sMagicDatabaseEntryAccuracy = "low"
                            End If
                        Else
                            sMagicDatabaseEntryAccuracy = "low"
                        End If
                        
                        txtPosition.Text = sMagicDatabaseEntryPosition & " (" & sMagicDatabaseEntryPositionType & ")"
                        txtType.Text = sMagicDatabaseEntryTypeDetail & " (" & sMagicDatabaseEntryType & ")"
                        txtPattern.Text = sMagicDatabaseEntryString
                        txtLength.Text = iMagicDatabaseEntryStringLength & " bytes"
                        txtAccuracy.Text = sMagicDatabaseEntryAccuracy & " (" & iMagicDatabaseEntryType & ")"
                        txtDescription.Text = sMagicDatabaseEntryDescription
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Sub lstResults_Click()
    Call ShowHitDetails(lstResults.Text)
End Sub

Private Sub mnuFileExitItem_Click()
    Unload Me
End Sub

Private Sub mnuHelpAboutItem_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnuHelpUpdatesItem_Click()
    Call OpenUpdateWebsite
End Sub

Private Sub mnuHelpWebsiteItem_Click()
    Call OpenProjectWebsite
End Sub

Public Sub ResetResultDetails()
    txtPosition.Text = vbNullString
    txtType.Text = vbNullString
    txtPattern.Text = vbNullString
    txtLength.Text = vbNullString
    txtAccuracy.Text = vbNullString
    txtDescription.Text = vbNullString
End Sub
