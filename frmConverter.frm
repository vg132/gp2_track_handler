VERSION 5.00
Begin VB.Form frmConverter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Converter"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   720
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   13
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Enabled         =   0   'False
      Height          =   325
      Left            =   4560
      TabIndex        =   11
      Top             =   3360
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output File"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   3070
      Width           =   4215
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse"
         Height          =   325
         Left            =   3120
         TabIndex        =   10
         Top             =   200
         Width           =   1000
      End
      Begin VB.TextBox txtFolder 
         Height          =   285
         Left            =   145
         TabIndex        =   7
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folder and Filename:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1470
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input File"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtFileName 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2550
         Width           =   4335
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   2640
         MultiSelect     =   2  'Extended
         Pattern         =   "*.ths;*.gif"
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   2415
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
         TabIndex        =   12
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2595
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   325
      Left            =   4560
      TabIndex        =   0
      Top             =   3840
      Width           =   1000
   End
End
Attribute VB_Name = "frmConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConvertFrom1 As SaveFileInfo
Dim ConvertFrom2 As SaveFileInfo2

Private Sub cmdConvert_Click()
    MDIForm1.MousePointer = 11
    frmConverter.MousePointer = 11
    Read = txtFileName.Text
    If FileLen(Read) = 14288 Then
        ConvertFrom11
    ElseIf FileLen(Read) = 9826 Then
        ConvertFrom12
    Else
        Picture1.Picture = LoadPicture(txtFileName.Text)
        SavePicture Picture1.Image, txtFolder.Text
    End If
    frmConverter.MousePointer = 0
    MDIForm1.MousePointer = 0
End Sub

Private Sub cmdOk_Click()
    Unload Me
    MDIForm1.Show
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    If Len(File1.Path) = 3 Then
        txtFileName.Text = File1.Path + File1.FileName
        CountNr = Len(File1.FileName)
        Read4 = File1.Path & Mid(File1.FileName, 1, CountNr - 4) & "14.ths"
    Else
        txtFileName.Text = File1.Path + "\" + File1.FileName
        CountNr = Len(File1.FileName)
        Read4 = File1.Path & Mid(File1.FileName, 1, CountNr - 4) & "14.ths"
    End If
    If FileLen(txtFileName.Text) = "9826" Then
        lblType.Caption = GP2TH + " 1.2 Save file"
        txtFolder.Text = Read4
        cmdConvert.Enabled = True
        Exit Sub
    ElseIf FileLen(txtFileName.Text) = "14288" Then
        lblType.Caption = GP2TH + " 1.0/1.1 Save file"
        txtFolder.Text = Read4
        cmdConvert.Enabled = True
        Exit Sub
    ElseIf UCase(Mid(File1.FileName, Len(File1.FileName) - 2, 3)) = UCase("gif") Then
        lblType.Caption = "GIF Picture file"
        If Len(File1.Path) = 3 Then
            txtFolder.Text = File1.Path & Mid(File1.FileName, 1, CountNr - 4) & ".bmp"
        Else
            txtFolder.Text = File1.Path & "\" & Mid(File1.FileName, 1, CountNr - 4) & ".bmp"
        End If
        cmdConvert.Enabled = True
        Exit Sub
    End If
    txtFileName.Text = ""
    txtFolder.Text = ""
    lblType.Caption = GP2TH + " 1.3/1.4 Save file"
    cmdConvert.Enabled = False
End Sub

Private Sub Form_Load()
    Dir1.Path = ProgramDir
    Drive1.Drive = ProgramDir
    txtFolder.Text = ProgramDir + "\"
End Sub

Public Sub ConvertFrom11()
    
    RecordLen = Len(ConvertFrom1)
    FileNum = FreeFile
    Open Read For Random As FileNum Len = RecordLen
    Read3 = Mid(txtFolder.Text, Len(txtFolder) - 3, 4)
    If Read3 <> ".ths" Then txtFolder.Text = txtFolder + ".ths"
    Read2 = txtFolder
    Read3 = ProgramDir + "\mall.lda"
    FileCopy Read3, Read2
    X = 1
    Do Until X > 16
        Get #FileNum, X, ConvertFrom1
        If Trim(ConvertFrom1.Country2) <> "No Data" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Adjective", Trim(ConvertFrom1.Country2), Read2)
        If Trim(ConvertFrom1.Country) <> "No Data" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Country", Trim(ConvertFrom1.Country), Read2)
        If Trim(ConvertFrom1.Laps) <> "No" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Laps", Trim(ConvertFrom1.Laps), Read2)
        If Trim(ConvertFrom1.Track) <> "No Data" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Name", Trim(ConvertFrom1.Track), Read2)
        If Trim(ConvertFrom1.Pic) <> "" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "SPic", Trim(ConvertFrom1.Pic), Read2)
        If Trim(ConvertFrom1.Pic2) <> "" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "BPic", Trim(ConvertFrom1.Pic2), Read2)
        If Trim(ConvertFrom1.Path) <> "No Data" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "TPath", Trim(ConvertFrom1.Path), Read2)
        If Trim(ConvertFrom1.Length) <> "No" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Length", Trim(ConvertFrom1.Length), Read2)
        X = X + 1
    Loop
    Close FileNum
    MsgBox "You can now use this file with " + TH + " v1.4!", vbInformation, TH
End Sub

Public Sub ConvertFrom12()
    RecordLen = Len(ConvertFrom2)
    FileNum = FreeFile
    Open Read For Random As FileNum Len = RecordLen
    
    Read3 = Mid(txtFolder.Text, Len(txtFolder) - 3, 4)
    If Read3 <> ".ths" Then txtFolder.Text = txtFolder + ".ths"

    Read2 = txtFolder
    Read3 = ProgramDir + "\mall.lda"
    FileCopy Read3, Read2
    X = 1
    Do Until X > 16
        Get #FileNum, X, ConvertFrom2
        If Trim(ConvertFrom2.Track) <> "NoTh" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Name", Trim(ConvertFrom2.Track), Read2)
        If Trim(ConvertFrom2.Country) <> "NoTh" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Country", Trim(ConvertFrom2.Country), Read2)
        If Trim(ConvertFrom2.Country2) <> "NoTh" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Adjective", Trim(ConvertFrom2.Country2), Read2)
        If Trim(ConvertFrom2.Laps) <> "NoT" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Laps", Trim(ConvertFrom2.Laps), Read2)
        If Trim(ConvertFrom2.Ware) <> "NoTh" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Ware", Trim(ConvertFrom2.Ware), Read2)
        If Trim(ConvertFrom2.Pic) <> "NoTh" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "SPic", Trim(ConvertFrom2.Pic), Read2)
        If Trim(ConvertFrom2.Pic2) <> "NoTh" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "SPic", Trim(ConvertFrom2.Pic2), Read2)
        If Trim(ConvertFrom2.Path) <> "NoTh" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "TPath", Trim(ConvertFrom2.Path), Read2)
        If Trim(ConvertFrom2.Length) <> "NoTh" Then Read = oMisc.WriteINI("Track " + Trim(Str(X)), "Length", Trim(ConvertFrom2.Length), Read2)
        X = X + 1
    Loop
    Close FileNum
    MsgBox "You can now use this file with " + TH + " v1.4!", vbInformation, TH
End Sub
