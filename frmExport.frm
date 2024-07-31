VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export To GP2"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Default         =   -1  'True
      Height          =   325
      Left            =   3887
      TabIndex        =   0
      Top             =   3965
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   325
      Left            =   3887
      TabIndex        =   1
      Top             =   3480
      Width           =   1000
   End
   Begin VB.Frame frameTrackData 
      Caption         =   "Track Data"
      Height          =   4290
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      Begin VB.CheckBox chkQSetup 
         Caption         =   "Qual Setup"
         Height          =   255
         Left            =   1560
         TabIndex        =   33
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CheckBox chkRSetup 
         Caption         =   "Race Setup"
         Height          =   255
         Left            =   1560
         TabIndex        =   32
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Select/Deselect All"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CheckBox chkDosPath 
         Caption         =   "GP2Edit File"
         Height          =   255
         Left            =   1560
         TabIndex        =   30
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CheckBox chkWare 
         Caption         =   "Tyre Ware"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chkLap 
         Caption         =   "Lap Data"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   3000
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrackLength 
         Caption         =   "Track Length"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2760
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkPicture 
         Caption         =   "Menu Pictures"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 9"
         Height          =   225
         Index           =   8
         Left            =   1560
         TabIndex        =   25
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 1"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 6"
         Height          =   225
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 7"
         Height          =   225
         Index           =   6
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 8"
         Height          =   225
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 10"
         Height          =   225
         Index           =   9
         Left            =   1560
         TabIndex        =   20
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 11"
         Height          =   225
         Index           =   10
         Left            =   1560
         TabIndex        =   19
         Top             =   840
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 12"
         Height          =   225
         Index           =   11
         Left            =   1560
         TabIndex        =   18
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 13"
         Height          =   225
         Index           =   12
         Left            =   1560
         TabIndex        =   17
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 14"
         Height          =   225
         Index           =   13
         Left            =   1560
         TabIndex        =   16
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 15"
         Height          =   225
         Index           =   14
         Left            =   1560
         TabIndex        =   15
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 16"
         Height          =   225
         Index           =   15
         Left            =   1560
         TabIndex        =   14
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 3"
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 5"
         Height          =   225
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 4"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkTrack 
         Caption         =   "Track 2"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   120
         TabIndex        =   9
         Top             =   2350
         Width           =   3375
      End
      Begin VB.CheckBox chkTimes 
         Caption         =   "Lap Times"
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CheckBox chkPoint 
         Caption         =   "Points"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3720
         Width           =   1575
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2100
      End
      Begin VB.CheckBox chkSettings 
         Caption         =   "GP2 Settings"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   3240
         Width           =   1455
      End
   End
   Begin VB.Label lblNote 
      Caption         =   $"frmExport.frx":030A
      Height          =   3255
      Left            =   3840
      TabIndex        =   2
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

Private Sub chkDosPath_Click()
    Test Check
End Sub

Private Sub chkLap_Click()
    Test Check
End Sub

Private Sub chkPicture_Click()
    Test Check
End Sub

Private Sub chkQSetup_Click()
    Test Check
End Sub

Private Sub chkRSetup_Click()
    Test Check
End Sub

Private Sub chkSettings_Click()
    Test Check
End Sub

Private Sub chkTheTrack_Click()
    Test Check
End Sub

Private Sub chkTrackName_Click()
    Test Check
End Sub
Private Sub chkPoint_Click()
    Test Check
End Sub

Private Sub chkTimes_Click()
    Test Check
End Sub

Private Sub chkTrack_Click(Index As Integer)
    Test Check
End Sub

Private Sub chkTrackLength_Click()
    Test Check
End Sub

Private Sub chkWare_Click()
    Test Check
End Sub

Private Sub cmdCancel_Click()
    frmExport.Hide
End Sub

Private Sub cmdExport_Click()
Dim Total As Double
    frmExport.MousePointer = 11
    SetAttr GP2Dir & "\gp2.exe", vbNormal

    If chkPicture.Value = 1 Then
        Read = oMisc.File_Exists(GP2Dir & "\gp2hipic.exe")
        If Read = False Then
            FileCopy ProgramDir & "\GP2Utils\gp2hipic.exe", GP2Dir & "\gp2hipic.exe"
        End If
    End If

    GetGP2Version
    Read = oMisc.File_Exists(GP2Dir + "\f1gstate.sav")
    If Read = True Then
        SetAttr GP2Dir & "\f1gstate.sav", vbNormal
        Exp.F1FileNum = FreeFile
        Open GP2Dir & "\f1gstate.sav" For Binary As Exp.F1FileNum
    End If
    Exp.GP2FileNum = FreeFile
    Open GP2Dir & "\gp2.exe" For Binary As Exp.GP2FileNum
    GetTrackNames
    SetAttribut
    GP2NameFile = ""
    For Exp.TrackNr = 0 To 15
        If chkTrack(Exp.TrackNr).Value = 1 Then
            If chkLap.Value = 1 Then ExportLaps
            If chkWare.Value = 1 Then ExportWare
            If chkTrackLength.Value = 1 Then ExportLength
            If chkPicture.Value = 1 Then ExportPictures
            If chkQSetup.Value = 1 Then ExportqualSetup
            If chkRSetup.Value = 1 Then ExportRaceSetup
            If chkTimes.Value = 1 Then
                ExportTime Race, F1gstate
                ExportRName F1gstate
                ExportRTeam F1gstate
                ExportRDate F1gstate
                ExportTime Qual, F1gstate
                ExportQName F1gstate
                ExportQTeam F1gstate
                ExportQDate F1gstate
            End If
            If chkTheTrack.Value = 1 Then ExportTracks
            If chkTrackName.Value = 1 Then
                ExportName
            Else
                GP2NameFile = GP2NameFile & TrackName(Exp.TrackNr) & Chr(0)
            End If
        Else
            GP2NameFile = GP2NameFile & TrackName(Exp.TrackNr) & Chr(0)
        End If
    Next
    GP2NameFile = GP2NameFile & String(16, Chr(0))
    For Count1 = 0 To 4
        For Exp.TrackNr = 0 To 15
            If chkTrack(Exp.TrackNr).Value = 1 Then
                ExportCountry
            Else
                GP2NameFile = GP2NameFile & Country(Exp.TrackNr) & Chr(0)
            End If
        Next
        GP2NameFile = GP2NameFile & String(16, Chr(0))
    Next
    For Exp.TrackNr = 0 To 15
        If chkTrack(Exp.TrackNr).Value = 1 Then
            ExportAdjectiv
        Else
            GP2NameFile = GP2NameFile & Adj(Exp.TrackNr) & Chr(0)
        End If
    Next
    If chkTrackName.Value = 1 Then
        GP2NameFile = GP2NameFile & String(16, Chr(0))
        Count1 = Len(GP2NameFile)
        Count2 = 4000 - Count1
        Put #Exp.GP2FileNum, oData.Text(GP2V) + 1, GP2NameFile
    Else
        GP2NameFile = ""
    End If
    If (chkPoint.Value = 1) And (chkPoint.Enabled = True) Then ExportPoints
    If (chkSettings.Value = 1) And (chkSettings.Enabled = True) Then
        'f1gstate.sav
        ExportQuickRace
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
    Close Exp.F1FileNum
    Close Exp.GP2FileNum

    Read = oMisc.GetShortName(GP2Dir & "\f1gstate.sav")
    WriteCheckSum Read
    If chkPicture.Value = 1 Or chkDosPath.Value = 1 Then
        X = ShellExecute(frmMain.hwnd, "open", ProgramDir & "\Bat\Export.bat", vbNullString, vbNullString, 1)
    End If

    For X = 0 To 15
        TrackName(X) = ""
        Country(X) = ""
        Adj(X) = ""
    Next
    frmExport.MousePointer = 0
    frmExport.Hide
Exit Sub

ErrorTrap:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
    X = 0
    Do Until X > 15
        TrackName(X) = ""
        Country(X) = ""
        Adj(X) = ""
        X = X + 1
    Loop
    GP2NameFile = ""
    Close FileNum
    Close FileNum2
    Close Exp.GP2FileNum
    Close Exp.F1FileNum
    frmExport.MousePointer = 0
End Sub

Private Sub Form_Activate()
    State = UnCheckAll

    Read = oMisc.File_Exists(GP2Dir + "\f1gstate.sav")
    If Read = False Then
        chkTimes.Enabled = False
        chkSettings.Enabled = False
        chkQSetup.Enabled = False
        chkRSetup.Enabled = False
        lblNote.Visible = True
    ElseIf Read = True Then
        chkTimes.Enabled = True
        chkSettings.Enabled = True
        chkQSetup.Enabled = True
        chkRSetup.Enabled = True
        lblNote.Visible = False
    End If

    Read = oMisc.ReadINI("Misc", "EXEPath", TempFile)
    If Read = "" Then
        chkDosPath.Enabled = False
    Else
        chkDosPath.Enabled = True
    End If
    For X = 0 To 15
        If Tracks(X) = False Then
            chkTrack(X).Value = 0
        Else
            chkTrack(X).Value = 1
        End If
    Next
End Sub

Public Sub GetTrackNames()
    Read = String(3000, " ")
    Get #Exp.GP2FileNum, oData.Text(GP2V) + 1, Read
    
    For Count1 = 0 To 15
        Count2 = InStr(1, Read, Chr(0))
        TrackName(Count1) = Mid(Read, 1, Count2 - 1)
        Read = Mid(Read, Count2 + 1)
    Next
    
    Read = Mid(Read, 17)
    For Count1 = 0 To 15
        Count2 = InStr(1, Read, Chr(0))
        Country(Count1) = Mid(Read, 1, Count2 - 1)
        Read = Mid(Read, Count2 + 1)
    Next

    Read = Mid(Read, 17)
    For Count1 = 0 To 3
        Read2 = String(17, Chr(0))
        Count2 = InStr(1, Read, Read2)
        Read = Mid(Read, Count2 + 1)
    Next
    
    Read = Mid(Read, 17)
    For Count1 = 0 To 15
        Count2 = InStr(1, Read, Chr(0))
        Adj(Count1) = Mid(Read, 1, Count2 - 1)
        Read = Mid(Read, Count2 + 1)
    Next
    Read = ""
End Sub

Private Sub Test(ByRef CheckType As Check2)
Dim oCtl As Control
    If CheckType = CheckAll Then
        For Each oCtl In frmExport
            If TypeOf oCtl Is CheckBox Then
                If oCtl.Enabled = True Then oCtl.Value = 1
            End If
        Next
        State = CheckAll
    ElseIf CheckType = UnCheckAll Then
        For Each oCtl In frmExport
            If TypeOf oCtl Is CheckBox Then
                oCtl.Value = 0
            End If
        Next
        State = UnCheckAll
    Else
        Update = True
        chkAll.Value = 1
        For Each oCtl In frmExport
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
