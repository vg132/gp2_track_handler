VERSION 5.00
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import From GP2"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   Icon            =   "frmImport.frx":0000
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
    'NewFile
    On Error GoTo ErrorTrap
    GP2FileNum = FreeFile
    Open GP2Dir + "\gp2.exe" For Binary As GP2FileNum
    
    With frmMain
        .txtPath.Enabled = True
        .txtPath.Enabled = False
    End With
    Read = ""

    GetGP2Version
    If frmImport.chkWare.Value = 1 Then ImportWare
    If frmImport.chkTrackName.Value = 1 Then
        ImportText
        FileInfo.Import = True
    End If
    If frmImport.chkPoint.Value = 1 Then ImportPoints
    If frmImport.chkLap.Value = 1 Then ImportLaps
    If frmImport.chkTrackLength.Value = 1 Then ImportLength
    If (frmImport.chkTime.Value = 1) And (chkTime.Enabled = True) Then
        FileNum = FreeFile
        Open GP2Dir + "\f1gstate.sav" For Binary As FileNum
        CountExport = 0
        Do Until CountExport > 15
            ImportQName
            ImportQDate
            ImportTimeFromGP2 Qual
            ImportQTeam
            ImportRName
            ImportRDate
            ImportTimeFromGP2 Race
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
            Open GP2Dir + "\F1gstate.sav" For Binary As F1SaveFileNum
                ImportQuick
            Close F1SaveFileNum
        End If
    End If

    FileInfo.Name = ""
    FileInfo.Path = ""
    LoadGP2Aid
    LoadFile
    FileInfo.Saved = False
    frmImport.MousePointer = 0
    frmImport.Hide
    Close GP2FileNum
Exit Sub
ErrorTrap:
    Print #Log, Date & " " & Time & " cmdImport_Click , Error Number: " & Err.Number & ", Error Description: " & Err.Description
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
    frmImport.MousePointer = 0
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
                oCtl.Value = 1
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

