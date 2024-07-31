VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDosPath 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GP2Edit Carset file (EXE)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3900
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   325
      Left            =   2640
      TabIndex        =   16
      Top             =   3120
      Width           =   1000
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "JAM Files"
      Height          =   315
      Index           =   8
      Left            =   1920
      TabIndex        =   15
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Laps Data"
      Height          =   315
      Index           =   15
      Left            =   1920
      TabIndex        =   14
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Sound Effects"
      Height          =   315
      Index           =   14
      Left            =   1920
      TabIndex        =   13
      Top             =   2040
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Pit Crew Colours"
      Height          =   315
      Index           =   13
      Left            =   1920
      TabIndex        =   12
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Cockpit Colours"
      Height          =   315
      Index           =   12
      Left            =   1920
      TabIndex        =   11
      Top             =   1320
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Helmet JAM"
      Height          =   315
      Index           =   11
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Damage Data"
      Height          =   315
      Index           =   10
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Car Settings"
      Height          =   315
      Index           =   9
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Points Data"
      Height          =   315
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Game Settings"
      Height          =   315
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Wheel JAMs"
      Height          =   315
      Index           =   6
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Cockpits"
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Menu Helmets *"
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Car JAMs"
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Camera Data"
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1500
   End
   Begin VB.CheckBox chkEXE 
      Caption         =   "Team Data"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Carset Settings"
      Height          =   4215
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton cmdInfo 
         Caption         =   "Command1"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   325
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "* If you export menu pictures with gp2hipic then don't use this function, the pictures in gp2 will be destoyed."
         Height          =   555
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   3435
      End
   End
End
Attribute VB_Name = "frmDosPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub cmdInfo_Click()
    Info
End Sub

Private Sub cmdOk_Click()
X = 0
Read = ""
Do Until X > 15
    If chkEXE(X).Value = 0 Then
        If X = 0 Then Read = Read + "TD "
        If X = 1 Then Read = Read + "CD "
        If X = 2 Then Read = Read + "CJ "
        If X = 3 Then Read = Read + "MH "
        If X = 4 Then Read = Read + "CP "
        If X = 5 Then Read = Read + "GS "
        If X = 6 Then Read = Read + "WJ "
        If X = 7 Then Read = Read + "PD "
        If X = 8 Then Read = Read + "JF "
        If X = 9 Then Read = Read + "CS "
        If X = 10 Then Read = Read + "DD "
        If X = 11 Then Read = Read + "HJ "
        If X = 12 Then Read = Read + "CC "
        If X = 13 Then Read = Read + "PC "
        If X = 14 Then Read = Read + "SE "
        If X = 15 Then Read = Read + "LD "
    End If
    X = X + 1
Loop
Read4 = oMisc.WriteINI("Misc", "EXE", Read, ProgramDir + "\WorkCopy.lda")
Unload frmDosPath
MDIForm1.Show
End Sub

Private Sub cmdRemove_Click()
    Read = ""
    Read4 = oMisc.WriteINI("Misc", "EXE", Read, ProgramDir + "\WorkCopy.lda")
    Read4 = oMisc.WriteINI("Misc", "EXEPath", Read, ProgramDir + "\WorkCopy.lda")
    Unload frmDosPath
    MDIForm1.Show
End Sub


Private Sub Form_Load()
    Read = oMisc.ReadINI("Misc", "EXE", ProgramDir + "\WorkCopy.lda")
    Read2 = oMisc.ReadINI("Misc", "EXEPath", ProgramDir + "\WorkCopy.lda")
    If Read = "" Then
        X = 0
        Do Until X > 15
            chkEXE(X).Value = 1
            X = X + 1
        Loop
    Else
        Y = 1
        X = 0
        If Mid(Read, Y, 2) = "TD" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 1
        If Mid(Read, Y, 2) = "CD" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 2
        If Mid(Read, Y, 2) = "CJ" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 3
        If Mid(Read, Y, 2) = "MH" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 4
        If Mid(Read, Y, 2) = "CP" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 5
        If Mid(Read, Y, 2) = "GS" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 6
        If Mid(Read, Y, 2) = "WJ" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 7
        If Mid(Read, Y, 2) = "PD" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 8
        If Mid(Read, Y, 2) = "JF" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 9
        If Mid(Read, Y, 2) = "CS" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 10
        If Mid(Read, Y, 2) = "DD" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 11
        If Mid(Read, Y, 2) = "HJ" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 12
        If Mid(Read, Y, 2) = "CC" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 13
        If Mid(Read, Y, 2) = "PC" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 14
        If Mid(Read, Y, 2) = "SE" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
        X = 15
        If Mid(Read, Y, 2) = "LD" Then
            chkEXE(X).Value = 0
            Y = Y + 3
        Else
            chkEXE(X).Value = 1
        End If
    End If
End Sub
