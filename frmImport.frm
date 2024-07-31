VERSION 5.00
Begin VB.Form frmImport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Import From GP2"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2805
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkWare 
      Caption         =   "Tyre Ware"
      Height          =   225
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkLap 
      Caption         =   "Lap Data"
      Height          =   225
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkTrackLength 
      Caption         =   "Track Length"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox chkPoint 
      Caption         =   "Point Data"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkTrackName 
      Caption         =   "Track Name"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   325
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1000
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Default         =   -1  'True
      Height          =   325
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Data to Import"
      Height          =   1980
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2775
      Begin VB.CheckBox chkSettings 
         Caption         =   "GP2 Settings"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Select/Deselect All"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkTime 
         Caption         =   "Lap Times"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   720
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Kalle2 As Boolean

Private Sub chkAll_Click()
    If (chkAll.Value = 1) And (Kalle2 = True) Then
        chkWare.Value = 1
        chkTrackName.Value = 1
        chkTrackLength.Value = 1
        chkLap.Value = 1
        chkPoint.Value = 1
        chkTime.Value = 1
        chkSettings.Value = 1
        chkAll.Value = 1
    ElseIf Kalle2 = True Then
        chkWare.Value = 0
        chkTrackName.Value = 0
        chkTrackLength.Value = 0
        chkLap.Value = 0
        chkPoint.Value = 0
        chkTime.Value = 0
        chkSettings.Value = 0
        chkAll.Value = 0
    End If
End Sub

Private Sub chkLap_Click()
    If chkAll.Value = 1 Then
        Kalle2 = False
        chkAll.Value = 0
        Kalle2 = True
    End If
End Sub

Private Sub chkPoint_Click()
    If chkAll.Value = 1 Then
        Kalle2 = False
        chkAll.Value = 0
        Kalle2 = True
    End If
End Sub

Private Sub chkSettings_Click()
    If chkAll.Value = 1 Then
        Kalle2 = False
        chkAll.Value = 0
        Kalle2 = True
    End If
End Sub

Private Sub chkTime_Click()
    If chkAll.Value = 1 Then
        Kalle2 = False
        chkAll.Value = 0
        Kalle2 = True
    End If
End Sub

Private Sub chkTrackLength_Click()
    If chkAll.Value = 1 Then
        Kalle2 = False
        chkAll.Value = 0
        Kalle2 = True
    End If
End Sub

Private Sub chkTrackName_Click()
    If chkAll.Value = 1 Then
        Kalle2 = False
        chkAll.Value = 0
        Kalle2 = True
    End If
End Sub

Private Sub chkWare_Click()
    If chkAll.Value = 1 Then
        Kalle2 = False
        chkAll.Value = 0
        Kalle2 = True
    End If
End Sub

Private Sub cmdCancel_Click()
    frmImport.Hide
    MDIForm1.Show
End Sub

Private Sub cmdImport_Click()
    On Error GoTo ErrorTrap
    GP2FileNum = FreeFile
    Open Gp2Dir + "\gp2.exe" For Binary As GP2FileNum
    
    MDIForm1.txtAdjectiv = ""
    MDIForm1.txtName = ""
    MDIForm1.txtCountry = ""
    MDIForm1.txtLaps = ""
    MDIForm1.txtLength = ""
    MDIForm1.txtPath.Enabled = True
    MDIForm1.txtPath = ""
    MDIForm1.txtPath.Enabled = False
    MDIForm1.txtTire = ""
    Read = ""
    Read2 = ""
    Unload frmExport
    Unload frmDosPath
    Unload frmPoint
    Unload frmAbout

    MousePointer = 11
    GetGP2Version
    If frmImport.chkWare.Value = 1 Then ImportWare
    If frmImport.chkTrackName.Value = 1 Then ImportText
    If frmImport.chkPoint.Value = 1 Then ImportPoints
    If frmImport.chkLap.Value = 1 Then ImportLaps
    If frmImport.chkTrackLength.Value = 1 Then ImportLength
    If (frmImport.chkTime.Value = 1) And (chkTime.Enabled = True) Then
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        CountExport = 0
        Do Until CountExport > 15
            ImportQName
            ImportQDate
            ImportQTime
            ImportQTeam
            ImportRName
            ImportRDate
            ImportRaceTime
            ImportRTeam
            CountExport = CountExport + 1
        Loop
        Close FileNum
    End If
    If chkSettings.Value = 1 Then
        ImportNullAsOne
        ImportLevel
        ImportSaveLap
        ImportGameSettings
        ImportPQPower
        ImportPRPower
        ImportPGrip
        ImportSpeed
        ImportCWeight
        ImportPWeight
        ImportUseTeam
        If chkTime.Enabled = True Then
            F1SaveFileNum = FreeFile
            Open Gp2Dir + "\F1gstate.sav" For Binary As F1SaveFileNum
                ImportQuick
            Close F1SaveFileNum
        End If
    End If
    FileInfo.FileName = ""
    FileInfo.FilePath = ""
    FileInfo.FileType = FileImport
    frmImport.MousePointer = 0
    frmImport.Hide
    Close GP2FileNum
    MDIForm1.Show
Exit Sub
ErrorTrap:
    MsgBox "Error # " + Str(Err.Number) + Err.Description
    frmImport.MousePointer = 0
End Sub

Private Sub Form_Activate()
    Read = oMisc.File_Exists(Gp2Dir + "\f1gstate.sav")
    If Read = False Then chkTime.Enabled = False
    If Read = True Then chkTime.Enabled = True
End Sub

Private Sub Form_Load()
    Kalle2 = True
End Sub
