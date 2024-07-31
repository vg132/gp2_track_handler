VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export To GP2"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5010
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Default         =   -1  'True
      Height          =   325
      Left            =   3840
      TabIndex        =   24
      Top             =   3720
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   325
      Left            =   3840
      TabIndex        =   22
      Top             =   3240
      Width           =   1000
   End
   Begin VB.CheckBox chkPicture 
      Caption         =   "Menu Pictures"
      Height          =   225
      Left            =   1560
      TabIndex        =   14
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 2"
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrackLength 
      Caption         =   "Track Length"
      Height          =   225
      Left            =   240
      TabIndex        =   20
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 4"
      Height          =   225
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 5"
      Height          =   225
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 3"
      Height          =   225
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 1"
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkLap 
      Caption         =   "Lap Data"
      Height          =   225
      Left            =   240
      TabIndex        =   15
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkWare 
      Caption         =   "Tyre Ware"
      Height          =   225
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 16"
      Height          =   225
      Index           =   15
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 15"
      Height          =   225
      Index           =   14
      Left            =   1560
      TabIndex        =   11
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 14"
      Height          =   225
      Index           =   13
      Left            =   1560
      TabIndex        =   10
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 13"
      Height          =   225
      Index           =   12
      Left            =   1560
      TabIndex        =   9
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 12"
      Height          =   225
      Index           =   11
      Left            =   1560
      TabIndex        =   8
      Top             =   1080
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 11"
      Height          =   225
      Index           =   10
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 10"
      Height          =   225
      Index           =   9
      Left            =   1560
      TabIndex        =   6
      Top             =   600
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 9"
      Height          =   225
      Index           =   8
      Left            =   1560
      TabIndex        =   5
      Top             =   360
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 8"
      Height          =   225
      Index           =   7
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 7"
      Height          =   225
      Index           =   6
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkTrack 
      Caption         =   "Track 6"
      Height          =   225
      Index           =   5
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkDosPath 
      Caption         =   "GP2Edit File"
      Height          =   225
      Left            =   1560
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Select/Deselect All"
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Frame frameTrackData 
      Caption         =   "Track Data"
      Height          =   4050
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   3735
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   120
         TabIndex        =   31
         Top             =   2350
         Width           =   3375
      End
      Begin VB.CheckBox chkTimes 
         Caption         =   "Lap Times"
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CheckBox chkPoint 
         Caption         =   "Points"
         Height          =   195
         Left            =   1560
         TabIndex        =   29
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CheckBox chkTheTrack 
         Caption         =   "Track"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2520
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrackName 
         Caption         =   "Track Name/Country"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2100
      End
      Begin VB.CheckBox chkSettings 
         Caption         =   "GP2 Settings"
         Height          =   195
         Left            =   1560
         TabIndex        =   26
         Top             =   3480
         Width           =   1455
      End
   End
   Begin VB.Label lblNote 
      Height          =   2895
      Left            =   3840
      TabIndex        =   23
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim Kalle As Boolean

Private Sub chkAll_Click()
    If (chkAll.Value = 1) And (Kalle = True) Then
        X = 0
        Do Until X > 15
            chkTrack(X).Value = 1
            X = X + 1
        Loop
        chkLap.Value = 1
        chkTrackLength.Value = 1
        chkPicture.Value = 1
        chkWare.Value = 1
        chkTimes.Value = 1
        chkDosPath.Value = 1
        chkPoint.Value = 1
        chkTheTrack.Value = 1
        chkTrackName.Value = 1
        chkSettings.Value = 1
        chkAll.Value = 1
    Else
        If Kalle = True Then
            X = 0
            Do Until X > 15
                chkTrack(X).Value = 0
                X = X + 1
            Loop
            chkLap.Value = 0
            chkPoint.Value = 0
            chkTrackLength.Value = 0
            chkPicture.Value = 0
            chkWare.Value = 0
            chkTimes.Value = 0
            chkDosPath.Value = 0
            chkTheTrack.Value = 0
            chkTrackName.Value = 0
            chkSettings.Value = 0
            chkAll.Value = 0
        End If
    End If
End Sub

Private Sub chkDosPath_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkLap_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkPicture_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkSettings_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkTheTrack_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkTrackName_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub
Private Sub chkPoint_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkTimes_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkTrack_Click(Index As Integer)
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkTrackLength_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub chkWare_Click()
    If chkAll.Value = 1 Then
        Kalle = False
        chkAll.Value = 0
        Kalle = True
    End If
    AllCheck
End Sub

Private Sub cmdCancel_Click()
    frmExport.Hide
    MDIForm1.Show
End Sub

Private Sub cmdExport_Click()
Dim Start As Long
Dim Stopp As Long
Dim Total As Double
    If chkPicture.Value = 1 Then
        Read = oMisc.File_Exists(Gp2Dir + "\gp2hipic.exe")
        If Read = False Then
            SourceFile = ProgramDir + "\GP2Utils\gp2hipic.exe"
            TargetFile = Gp2Dir + "\gp2hipic.exe"
            FileCopy SourceFile, TargetFile
        End If
    End If
    On Error Resume Next
    DeleteFile Gp2Dir + "\_menupic.bat"
    On Error GoTo ErrorTrap
    frmExport.MousePointer = 11
    GetGP2Version
    GP2FileNum = FreeFile
    Open Gp2Dir + "\gp2.exe" For Binary As GP2FileNum
    GetTrackNames
    SetAttribut
    CountExport = 0
    GP2NameFile = ""
    Do Until CountExport > 15
        If chkTrack(CountExport).Value = 1 Then
            If chkLap.Value = 1 Then ExportLaps
            If chkWare.Value = 1 Then ExportWare
            If chkTrackLength.Value = 1 Then ExportLength
            If chkPicture.Value = 1 Then ExportPictures
            If chkTimes.Value = 1 Then
                ExportRaceTime
                ExportRName
                ExportRTeam
                ExportRDate
                ExportQualTime
                ExportQName
                ExportQTeam
                ExportQDate
            End If
            If chkTheTrack.Value = 1 Then ExportTracks
            If chkTrackName.Value = 1 Then
                ExportName
            Else
                GP2NameFile = GP2NameFile + TrackName(CountExport) + Chr(0)
            End If
        Else
            GP2NameFile = GP2NameFile + TrackName(CountExport) + Chr(0)
        End If
        CountExport = CountExport + 1
    Loop
    GP2NameFile = GP2NameFile + String(16, Chr(0))
    CountExport = 0
    Count1 = 0
    Do Until Count1 > 4
        Do Until CountExport > 15
            If chkTrack(CountExport).Value = 1 Then
                ExportCountry
            Else
                GP2NameFile = GP2NameFile + Country(CountExport) + Chr(0)
            End If
            CountExport = CountExport + 1
        Loop
        Count1 = Count1 + 1
        CountExport = 0
        GP2NameFile = GP2NameFile + String(16, Chr(0))
    Loop
    CountExport = 0
    Do Until CountExport > 15
        If chkTrack(CountExport).Value = 1 Then
            ExportAdjectiv
        Else
            GP2NameFile = GP2NameFile + Adj(CountExport) + Chr(0)
        End If
        CountExport = CountExport + 1
    Loop
    GP2NameFile = GP2NameFile + String(16, Chr(0))
    Count1 = Len(GP2NameFile)
    Count2 = 4000 - Count1
    Put #GP2FileNum, oData.Text(GP2V) + 1, GP2NameFile
    If (chkSettings.Value = 1) And (chkSettings.Enabled = True) Then
        'f1gstate.sav
        F1SaveFileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As F1SaveFileNum
        ExportQuickRace
        Close F1SaveFileNum
        'gp2.exe
        ExportNullAsOne
        ExportLevel
        ExportSaveLap
        ExportCarHelp
        ExportPQPower
        ExportPQPower
        ExportPRPower
        ExportPGrip
        ExportPWeight
        ExportCWeight
        ExportSpeed
        ExportUseTeam
    End If
    If (chkDosPath.Value = 1) And (chkDosPath.Enabled = True) Then ExportDos
    If ((chkSettings.Value = 1) And (chkSettings.Enabled = True)) Or ((chkTimes.Value = 1) And (chkTimes.Enabled = True)) Or (chkDosPath.Value = 1) Or (chkPicture.Value = 1) Then
        SourceFile = ProgramDir + "\gp2utils\check.exe"
        TargetFile = Gp2Dir + "\$$check.exe"
        FileCopy SourceFile, TargetFile
        FileNum = FreeFile
        Open Gp2Dir + "\_MenuPic.bat" For Append As FileNum
        Print #FileNum, Mid(Gp2Dir, 1, 2)
        Print #FileNum, "cd " + Gp2Dir
        Print #FileNum, Gp2Dir + "\$$Check f1gstate.sav"
        Print #FileNum, "del $$Check.exe"
        Close FileNum
        Read = oMisc.File_Exists("c:\command.com")
        Dim retval
        If Read = True Then
            ChDir Gp2Dir
            retval = Shell("c:\command.com /c " + Gp2Dir + "\_MenuPic.bat", vbNormalFocus)
        Else
            retval = Shell(Gp2Dir + "\_MenuPic.bat", vbNormalFocus)
        End If
    End If
    If (chkPoint.Value = 1) And (chkPoint.Enabled = True) Then ExportPoints
    Close GP2FileNum
    frmExport.MousePointer = 0
    frmExport.Hide
    If ((chkPicture.Value = 1) And (chkPicture.Enabled = True)) Or ((chkDosPath.Value = 1) And (chkDosPath.Enabled = True)) Or chkTimes.Value = 1 Then Exit Sub
    MDIForm1.Show
    X = 0
    Do Until X > 15
        TrackName(X) = ""
        Country(X) = ""
        Adj(X) = ""
        X = X + 1
    Loop
    Exit Sub

ErrorTrap:
    MsgBox "Error # " + Str(Err.Number) + Err.Description
    frmExport.MousePointer = 0
    X = 0
    Do Until X > 15
        TrackName(X) = ""
        Country(X) = ""
        Adj(X) = ""
        X = X + 1
    Loop
    Close FileNum
    Close FileNum2
    Close GP2FileNum
End Sub

Private Sub Form_Activate()
    Kalle = True
    SaveLastClick
    
    Read = oMisc.File_Exists(Gp2Dir + "\f1gstate.sav")
    If Read = False Then
        chkTimes.Enabled = False
        chkSettings.Enabled = False
        lblNote.Visible = True
    ElseIf Read = True Then
        chkTimes.Enabled = True
        chkSettings.Enabled = True
        lblNote.Visible = False
    End If
    
    Read = oMisc.ReadINI("Misc", "EXEPath", ProgramDir + "\WorkCopy.lda")
    If Read = "" Then chkDosPath.Enabled = False
End Sub

Private Sub Form_Load()
    Kalle = True
End Sub

Public Sub AllCheck()
    If (chkAll.Value = 0) And (Kalle = True) Then
        X = 0
        Do Until X > 15
            If chkTrack(X).Value = 0 Then Exit Sub
            X = X + 1
        Loop
        If chkLap.Value = 0 Then Exit Sub
        If chkTrackLength.Value = 0 Then Exit Sub
        If chkPicture.Value = 0 Then Exit Sub
        If chkWare.Value = 0 Then Exit Sub
        If chkTimes.Value = 0 Then Exit Sub
        If chkDosPath.Value = 0 Then Exit Sub
        If chkPoint.Value = 0 Then Exit Sub
        If chkTheTrack.Value = 0 Then Exit Sub
        If chkTrackName.Value = 0 Then Exit Sub
        If chkSettings.Value = 0 Then Exit Sub
        chkAll.Value = 1
    End If
End Sub

Public Sub GetTrackNames()
    Count1 = 0
    CountNr = 1
    Read4 = String(3000, " ")
    Get #GP2FileNum, oData.Text(GP2V) + 1, Read4
    
    Do Until Count1 > 15
        Read = String(1, " ")
        Read2 = ""
        Do Until Read = Chr(0)
            Read = Mid(Read4, CountNr, 1)
            If Read <> Chr(0) Then Read2 = Read2 + Read
            CountNr = CountNr + 1
        Loop
        TrackName(Count1) = Read2
        Count1 = Count1 + 1
    Loop
    CountNr = CountNr + 16
    Count1 = 0
    Do Until Count1 > 15
        Read = String(1, " ")
        Read2 = ""
        Do Until Read = Chr(0)
            Read = Mid(Read4, CountNr, 1)
            If Read <> Chr(0) Then Read2 = Read2 + Read
            CountNr = CountNr + 1
        Loop
        Country(Count1) = Read2
        Count1 = Count1 + 1
    Loop
    X = 1
    Do Until X > 4
        CountNr = CountNr + 16
        Count1 = 1
        Do Until Count1 > 16
            Read = String(1, " ")
            Read2 = ""
            Do Until Read = Chr(0)
                Read = Mid(Read4, CountNr, 1)
                If Read <> Chr(0) Then Read2 = Read2 + Read
                CountNr = CountNr + 1
            Loop
            Count1 = Count1 + 1
        Loop
        X = X + 1
    Loop
    CountNr = CountNr + 16
    Count1 = 0
    Do Until Count1 > 15
        Read = String(1, " ")
        Read2 = ""
        Do Until Read = Chr(0)
            Read = Mid(Read4, CountNr, 1)
            If Read <> Chr(0) Then Read2 = Read2 + Read
            CountNr = CountNr + 1
        Loop
        Adj(Count1) = Read2
        Count1 = Count1 + 1
    Loop
    Read4 = ""
End Sub
