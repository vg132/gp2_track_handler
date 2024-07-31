VERSION 5.00
Begin VB.Form frmPoint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GP2 Point Editor"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   19
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   29
      Top             =   1320
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   24
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   28
      Top             =   240
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   16
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   27
      Top             =   240
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   6
      Left            =   360
      MaxLength       =   2
      TabIndex        =   26
      Top             =   2400
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   25
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   25
      Top             =   600
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   3
      Left            =   360
      MaxLength       =   2
      TabIndex        =   24
      Top             =   1320
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   13
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   23
      Top             =   2040
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   21
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   22
      Top             =   2040
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   0
      Left            =   360
      MaxLength       =   2
      TabIndex        =   21
      Top             =   240
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   18
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   20
      Top             =   960
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   12
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   19
      Top             =   1680
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   15
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   18
      Top             =   2760
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   8
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   17
      Top             =   240
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   17
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   16
      Top             =   600
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   1
      Left            =   360
      MaxLength       =   2
      TabIndex        =   15
      Top             =   600
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   2
      Left            =   360
      MaxLength       =   2
      TabIndex        =   14
      Top             =   960
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   7
      Left            =   360
      MaxLength       =   2
      TabIndex        =   13
      Top             =   2760
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   14
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2400
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   20
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1680
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   4
      Left            =   360
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1680
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   5
      Left            =   360
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2040
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   9
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   8
      Top             =   600
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   22
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2400
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   23
      Left            =   2280
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2760
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   10
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   5
      Top             =   960
      Width           =   400
   End
   Begin VB.TextBox txtPoint 
      Height          =   285
      Index           =   11
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1320
      Width           =   400
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   325
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   325
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   900
   End
   Begin VB.CommandButton cmdCart 
      Caption         =   "CART"
      Height          =   325
      Left            =   3000
      TabIndex        =   1
      Top             =   1800
      Width           =   900
   End
   Begin VB.CommandButton cmdF1 
      Caption         =   "F1"
      Height          =   325
      Left            =   3000
      TabIndex        =   0
      Top             =   1320
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Point Editor (0-99)"
      ClipControls    =   0   'False
      Height          =   3255
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Export"
         Height          =   325
         Left            =   3960
         TabIndex        =   58
         Top             =   2280
         Width           =   900
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "&Import"
         Height          =   325
         Left            =   3960
         TabIndex        =   57
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "8th"
         Height          =   195
         Left            =   70
         TabIndex        =   56
         Top             =   2775
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "10th"
         Height          =   195
         Left            =   900
         TabIndex        =   55
         Top             =   615
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "9th"
         Height          =   195
         Left            =   900
         TabIndex        =   54
         Top             =   255
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "5th"
         Height          =   195
         Left            =   70
         TabIndex        =   53
         Top             =   1695
         Width           =   225
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "7th"
         Height          =   195
         Left            =   70
         TabIndex        =   52
         Top             =   2415
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "4th"
         Height          =   195
         Left            =   70
         TabIndex        =   51
         Top             =   1335
         Width           =   225
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "3th"
         Height          =   195
         Left            =   70
         TabIndex        =   50
         Top             =   975
         Width           =   225
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "6th"
         Height          =   195
         Left            =   70
         TabIndex        =   49
         Top             =   2055
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "2nd"
         Height          =   195
         Left            =   70
         TabIndex        =   48
         Top             =   615
         Width           =   270
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "1st"
         Height          =   195
         Left            =   70
         TabIndex        =   47
         Top             =   300
         Width           =   210
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "20th"
         Height          =   195
         Left            =   1800
         TabIndex        =   46
         Top             =   1335
         Width           =   315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "16th"
         Height          =   195
         Left            =   900
         TabIndex        =   45
         Top             =   2775
         Width           =   315
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "19th"
         Height          =   195
         Left            =   1800
         TabIndex        =   44
         Top             =   975
         Width           =   315
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "14th"
         Height          =   195
         Left            =   900
         TabIndex        =   43
         Top             =   2055
         Width           =   315
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "18th"
         Height          =   195
         Left            =   1800
         TabIndex        =   42
         Top             =   615
         Width           =   315
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "15th"
         Height          =   195
         Left            =   900
         TabIndex        =   41
         Top             =   2415
         Width           =   315
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "13th"
         Height          =   195
         Left            =   900
         TabIndex        =   40
         Top             =   1695
         Width           =   315
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "17th"
         Height          =   195
         Left            =   1800
         TabIndex        =   39
         Top             =   255
         Width           =   315
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "12th"
         Height          =   195
         Left            =   900
         TabIndex        =   38
         Top             =   1335
         Width           =   315
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "11th"
         Height          =   195
         Left            =   900
         TabIndex        =   37
         Top             =   975
         Width           =   315
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "26th"
         Height          =   195
         Left            =   2850
         TabIndex        =   36
         Top             =   615
         Width           =   315
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "22th"
         Height          =   195
         Left            =   1800
         TabIndex        =   35
         Top             =   2055
         Width           =   315
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "25th"
         Height          =   195
         Left            =   2850
         TabIndex        =   34
         Top             =   255
         Width           =   315
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "24th"
         Height          =   195
         Left            =   1800
         TabIndex        =   33
         Top             =   2775
         Width           =   315
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "23th"
         Height          =   195
         Left            =   1800
         TabIndex        =   32
         Top             =   2415
         Width           =   315
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "21st"
         Height          =   195
         Left            =   1800
         TabIndex        =   31
         Top             =   1695
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload frmPoint
    MDIForm1.Show
End Sub

Private Sub cmdCart_Click()
    txtPoint(0).Text = "20"
    txtPoint(1).Text = "16"
    txtPoint(2).Text = "14"
    txtPoint(3).Text = "12"
    txtPoint(4).Text = "10"
    txtPoint(5).Text = "8"
    txtPoint(6).Text = "6"
    txtPoint(7).Text = "5"
    txtPoint(8).Text = "4"
    txtPoint(9).Text = "3"
    txtPoint(10).Text = "2"
    txtPoint(11).Text = "1"
    X = 12
    Do Until X > 25
        txtPoint(X).Text = "0"
        X = X + 1
    Loop
    
End Sub

Private Sub cmdF1_Click()
    txtPoint(0).Text = "10"
    txtPoint(1).Text = "6"
    txtPoint(2).Text = "4"
    txtPoint(3).Text = "3"
    txtPoint(4).Text = "2"
    txtPoint(5).Text = "1"
    X = 6
    Do Until X > 25
        txtPoint(X).Text = "0"
        X = X + 1
    Loop
    
End Sub

Private Sub cmdImport_Click()
    GP2FileNum = FreeFile
    Open Gp2Dir + "\GP2.exe" For Binary As GP2FileNum
    ImportPoints
    Form_Load
    Close GP2FileNum
End Sub

Private Sub cmdOk_Click()
    X = 0
    Read = ""
    Read2 = ""
    Do Until X > 25
        Read2 = txtPoint(X).Text
        If Len(Read2) = 1 Then Read2 = "0" + Read2
        If Len(Read2) = 0 Then Read2 = "00"
        Read = Read + Read2
        X = X + 1
    Loop
    Read = oMisc.WriteINI("Misc", "Point", Read, ProgramDir + "\WorkCopy.lda")
    Unload frmPoint
    MDIForm1.Show
End Sub

Private Sub cmdExport_Click()
    X = 0
    Read = ""
    Read2 = ""
    Do Until X > 25
        Read2 = txtPoint(X).Text
        If Len(Read2) = 1 Then Read2 = "0" + Read2
        If Len(Read2) = 0 Then Read2 = "00"
        Read = Read + Read2
        X = X + 1
    Loop
    Read = oMisc.WriteINI("Misc", "Point", Read, ProgramDir + "\WorkCopy.lda")
    GP2FileNum = FreeFile
    Open Gp2Dir + "\GP2.exe" For Binary As GP2FileNum
    ExportPoints
    Close GP2FileNum
End Sub

Private Sub Form_Load()
    Read = oMisc.ReadINI("Misc", "Point", ProgramDir + "\WorkCopy.lda")
    CountNr = Len(Read)
    X = 0
    Count1 = 1
    Do Until Count1 > CountNr
        Read2 = Mid(Read, Count1, 2)
        If Mid(Read2, 1, 1) = 0 Then Read2 = Mid(Read2, 2, 1)
        txtPoint(X) = Read2
        Count1 = Count1 + 2
        X = X + 1
    Loop
    Do Until X > 25
        txtPoint(X) = "0"
        X = X + 1
    Loop
End Sub
