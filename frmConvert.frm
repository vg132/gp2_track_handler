VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Converter"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   325
      Left            =   4500
      TabIndex        =   9
      Top             =   3720
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input File"
      Height          =   3015
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   2415
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.FileListBox File1 
         Height          =   2625
         Left            =   2640
         MultiSelect     =   2  'Extended
         Pattern         =   "*.ths;*.set;database.tdb"
         TabIndex        =   5
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   2160
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output File"
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   3120
      Width           =   4215
      Begin VB.TextBox txtFolder 
         Height          =   285
         Left            =   145
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folder and Filename:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Enabled         =   0   'False
      Height          =   325
      Left            =   4500
      TabIndex        =   0
      Top             =   3240
      Width           =   1000
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FilePath As String

Private Sub cmdConvert_Click()
    If File1.FileName = "database.tdb" Then
        LapTimeCon FilePath
        frmMain.LoadTimeData
        MsgBox LoadResString(126), vbInformation, TH
        Exit Sub
    End If
    FileCopy ProgramDir & "\mall.lda", txtFolder.Text
    If FileLen(FilePath) = "14336" Then
        WinTrack2TH FilePath, txtFolder.Text
    ElseIf FileLen(FilePath) = "14288" Then
        Conv1 FilePath, txtFolder.Text
    ElseIf FileLen(FilePath) = "9826" Then
        Conv2 FilePath, txtFolder.Text
    End If
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    txtFolder.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    If Len(File1.Path) = 3 Then
        FilePath = File1.Path & File1.FileName
        txtFolder.Text = File1.Path & "th14" & File1.FileName
    Else
        FilePath = File1.Path & "\" & File1.FileName
        txtFolder.Text = File1.Path & "\th14" & File1.FileName
    End If
    txtFolder.Enabled = True
    If LCase(File1.FileName) = LCase("database.tdb") Then
        frmConvert.Caption = "File Converter - Track Handler v1.3 Lap Time Database"
        cmdConvert.Enabled = True
        txtFolder.Enabled = False
    ElseIf FileLen(FilePath) = "14288" Then
        frmConvert.Caption = "File Converter - Track Handler v1.0/1.1 file"
        cmdConvert.Enabled = True
    ElseIf FileLen(FilePath) = "14336" Then
        frmConvert.Caption = "File Converter - WinTrackMan 1.5 file"
        If Len(File1.Path) = 3 Then
            txtFolder.Text = txtFolder.Text & File1.FileName
        Else
            txtFolder.Text = txtFolder.Text & "\" & Mid(File1.FileName, 1, Len(File1.FileName) - 3) & "ths"
        End If
        cmdConvert.Enabled = True
    ElseIf FileLen(FilePath) = "9826" Then
        frmConvert.Caption = "File Converter - Track Handler v1.2 file"
        cmdConvert.Enabled = True
    Else
        FileNum = FreeFile
        Open FilePath For Binary As FileNum
        Read = String(5, " ")
        Get #FileNum, 1, Read
        Close FileNum
        If UCase(Read) = "#TH14" Then
            frmConvert.Caption = "File Converter - Track Handler v1.4"
        Else
            frmConvert.Caption = "File Converter - Unknown file"
        End If
        cmdConvert.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Drive1.Drive = Mid(ProgramDir, 1, 2)
    Dir1.Path = ProgramDir
    txtFolder.Text = Dir1.Path
End Sub

