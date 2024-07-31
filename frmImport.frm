VERSION 5.00
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import From GP2"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Select Data to Import"
      Height          =   4185
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   100
         TabIndex        =   27
         Top             =   2270
         Width           =   2570
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 10"
         Height          =   255
         Index           =   9
         Left            =   1560
         TabIndex        =   25
         Top             =   480
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 11"
         Height          =   255
         Index           =   10
         Left            =   1560
         TabIndex        =   23
         Top             =   720
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 12"
         Height          =   255
         Index           =   11
         Left            =   1560
         TabIndex        =   22
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 4"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 5"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 13"
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   19
         Top             =   1200
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 6"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 14"
         Height          =   255
         Index           =   13
         Left            =   1560
         TabIndex        =   17
         Top             =   1440
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 7"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 15"
         Height          =   255
         Index           =   14
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 9"
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 8"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 16"
         Height          =   255
         Index           =   15
         Left            =   1560
         TabIndex        =   11
         Top             =   1920
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Default         =   -1  'True
         Height          =   325
         Left            =   1635
         TabIndex        =   10
         Top             =   3720
         Width           =   1000
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   325
         Left            =   120
         TabIndex        =   9
         Top             =   3720
         Width           =   1000
      End
      Begin VB.CheckBox chkTrackName 
         Caption         =   "Track Names"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkPoint 
         Caption         =   "Point Data"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkTrackLength 
         Caption         =   "Track Length"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkLap 
         Caption         =   "Lap Data"
         Height          =   225
         Left            =   1560
         TabIndex        =   5
         Top             =   2400
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkWare 
         Caption         =   "Tyre Ware"
         Height          =   225
         Left            =   1560
         TabIndex        =   4
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkSettings 
         Caption         =   "GP2 Settings"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Select/Deselect All"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CheckBox chkTime 
         Caption         =   "Lap Times"
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   2880
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
Dim State As Check2
Dim Update As Boolean

Private Enum Check2
    Check = 0
    UnCheckAll = 1
    CheckAll = 2
End Enum

Private Sub chkAll_Click()
    If Update = False Then
        If State = UnCheckAll Then
            Test CheckAll
        ElseIf State = CheckAll Then
            Test UnCheckAll
        End If
    End If
End Sub

Private Sub chkLap_Click()
    Test Check
End Sub

Private Sub chkPoint_Click()
    Test Check
End Sub

Private Sub chkSettings_Click()
    Test Check
End Sub

Private Sub chkTime_Click()
    Test Check
End Sub

Private Sub chkTrackLength_Click()
    Test Check
End Sub

Private Sub chkTrackName_Click()
    Test Check
End Sub

Private Sub chkWare_Click()
    Test Check
End Sub

Private Sub cmdCancel_Click()
    frmImport.Hide
End Sub

Private Sub cmdImport_Click()
    frmImport.MousePointer = 11

    On Error GoTo ErrorTrap
    Exp.GP2FileNum = FreeFile
    Open GP2Dir + "\gp2.exe" For Binary As Exp.GP2FileNum

    If (frmImport.chkTime.Value = 1) And (chkTime.Enabled = True) Then
        Exp.F1FileNum = FreeFile
        Open GP2Dir + "\f1gstate.sav" For Binary As Exp.F1FileNum
    End If

    GetGP2Version
    If chkTrackName.Value = 1 Then ImportText
    For Exp.TrackNr = 0 To 15
        If chkTrack(Exp.TrackNr).Value = 1 Then
            If chkLap.Value = 1 Then ImportLaps
            If chkTrackLength.Value = 1 Then ImportLength
            If chkWare.Value = 1 Then ImportWare
            If chkTrackName.Value = 1 Then
                oMisc.WriteINI "Track " & Exp.TrackNr + 1, "Name", TrackName(Exp.TrackNr), TempFile
                oMisc.WriteINI "Track " & Exp.TrackNr + 1, "Adjective", Adj(Exp.TrackNr), TempFile
                oMisc.WriteINI "Track " & Exp.TrackNr + 1, "Country", Country(Exp.TrackNr), TempFile
                If Exp.TrackNr + 1 > 9 Then
                    Read = GP2Dir & "\Circuits\f1ct" & Exp.TrackNr + 1 & ".dat"
                Else
                    Read = GP2Dir & "\Circuits\f1ct0" & Exp.TrackNr + 1 & ".dat"
                End If
                oMisc.WriteINI "Track " & Exp.TrackNr + 1, "TPath", Read, TempFile
                Tracks(X) = True
            End If
            If (chkTime.Value = 1) And (chkTime.Enabled = True) Then
                ImportQName F1gstate
                ImportQDate F1gstate
                ImportTime Qual, F1gstate
                ImportQTeam F1gstate
                ImportRName F1gstate
                ImportRDate F1gstate
                ImportTime Race, F1gstate
                ImportRTeam F1gstate
            End If
        End If
    Next
    If frmImport.chkPoint.Value = 1 Then ImportPoints
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
        'If chkTime is enabled then f1gstate is installed on the system
        If chkTime.Enabled = True Then
            ImportQuick
        End If
    End If
    Close Exp.GP2FileNum
    Close Exp.F1FileNum

    FileInfo.Name = ""
    FileInfo.Path = ""
    FileInfo.Saved = False
    
    LoadGP2Aid
    LoadFile
    frmImport.MousePointer = 0
    frmImport.Hide
Exit Sub
ErrorTrap:
    frmImport.MousePointer = 0
    Close Exp.GP2FileNum
    Close Exp.F1FileNum
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
    frmImport.Hide
End Sub

Private Sub Form_Activate()
    Read = oMisc.File_Exists(GP2Dir + "\f1gstate.sav")
    If Read = False Then chkTime.Enabled = False
    If Read = True Then chkTime.Enabled = True
    State = UnCheckAll
    Update = False
End Sub

Private Sub Test(ByRef CheckType As Check2)
Dim oCtl As Control
    If CheckType = CheckAll Then
        For Each oCtl In frmImport
            If TypeOf oCtl Is CheckBox Then
                If oCtl.Enabled = True Then oCtl.Value = 1
            End If
        Next
        State = CheckAll
    ElseIf CheckType = UnCheckAll Then
        For Each oCtl In frmImport
            If TypeOf oCtl Is CheckBox Then
                oCtl.Value = 0
            End If
        Next
        State = UnCheckAll
    Else
        Update = True
        chkAll.Value = 1
        For Each oCtl In frmImport
            If TypeOf oCtl Is CheckBox Then
                If oCtl.Value = 0 Then
                    chkAll.Value = 0
                    Exit For
                End If
            End If
        Next
        Update = False
    End If
    Set oCtl = Nothing
End Sub

Private Sub chkTrack_Click(Index As Integer)
    Test Check
End Sub
