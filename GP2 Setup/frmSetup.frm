VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gp2 Setup File Editor"
   ClientHeight    =   4260
   ClientLeft      =   4485
   ClientTop       =   1725
   ClientWidth     =   6270
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "frmSetup"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMisc 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   -10000
      TabIndex        =   68
      Top             =   480
      Width           =   4215
      Begin VB.Frame fraBrake 
         Caption         =   "Brake Balans"
         Height          =   1095
         Left            =   2160
         TabIndex        =   92
         Top             =   1320
         Width           =   2055
         Begin VB.HScrollBar hscBrake 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   200
            TabIndex        =   93
            Top             =   720
            Value           =   102
            Width           =   1815
         End
         Begin VB.Label lblFront 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "37500"
            Height          =   195
            Left            =   1485
            TabIndex        =   97
            Top             =   480
            Width           =   450
         End
         Begin VB.Label lblRear 
            AutoSize        =   -1  'True
            Caption         =   "62500"
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Front"
            Height          =   195
            Left            =   1590
            TabIndex        =   95
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Rear"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   345
         End
      End
      Begin VB.Frame fraRollBar 
         Caption         =   "Anti Roll Bar"
         Height          =   1095
         Left            =   2160
         TabIndex        =   87
         Top             =   0
         Width           =   2055
         Begin VB.ComboBox cboRollF 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   240
            Width           =   1110
         End
         Begin VB.ComboBox cboRollR 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   660
            Width           =   1110
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Rear"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Front"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   300
            Width           =   360
         End
      End
      Begin VB.Frame fraGear 
         Caption         =   "Gears"
         ClipControls    =   0   'False
         Height          =   2415
         Left            =   0
         TabIndex        =   74
         Top             =   0
         Width           =   2055
         Begin VB.TextBox txtGear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   80
            Text            =   "59"
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox txtGear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   79
            Text            =   "54"
            Top             =   1680
            Width           =   375
         End
         Begin VB.TextBox txtGear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   78
            Text            =   "48"
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox txtGear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   77
            Text            =   "42"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtGear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   76
            Text            =   "35"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtGear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   75
            Text            =   "28"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "6th Gear (21-80)"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   2055
            Width           =   1155
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "5th Gear (20-79)"
            Height          =   195
            Left            =   120
            TabIndex        =   85
            Top             =   1695
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "4th Gear (19-78)"
            Height          =   195
            Left            =   120
            TabIndex        =   84
            Top             =   1335
            Width           =   1155
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "1st Gear (16-75)"
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   255
            Width           =   1140
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "2nd Gear (17-76)"
            Height          =   195
            Left            =   120
            TabIndex        =   82
            Top             =   615
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "3rd Gear (18-77)"
            Height          =   195
            Left            =   120
            TabIndex        =   81
            Top             =   975
            Width           =   1155
         End
      End
      Begin VB.Frame fraWing 
         Caption         =   "Wings"
         Height          =   975
         Left            =   0
         TabIndex        =   69
         Top             =   2520
         Width           =   2055
         Begin VB.TextBox txtRWing 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   71
            Text            =   "12"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox txtFWing 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   70
            Text            =   "10"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Rear Wing (1-20)"
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   645
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Front Wing (1-20)"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   285
            Width           =   1230
         End
      End
   End
   Begin VB.Frame fraRear 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   -10000
      TabIndex        =   35
      Top             =   480
      Width           =   5895
      Begin VB.Frame fraLeftRear 
         Caption         =   "Left Rear"
         Height          =   2895
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   2895
         Begin VB.ComboBox cboSpringR 
            Height          =   315
            Index           =   0
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar hscHeightR 
            Height          =   255
            Index           =   0
            LargeChange     =   10
            Left            =   1560
            Max             =   160
            Min             =   40
            TabIndex        =   58
            Top             =   2520
            Value           =   116
            Width           =   1215
         End
         Begin VB.TextBox txtSlowReboundR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   57
            Text            =   "15"
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtSlowBumpR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   56
            Text            =   "15"
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtFastReboundR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   55
            Text            =   "4"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtFastBumpR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   54
            Text            =   "4"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtPacR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   53
            Text            =   "0"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblHeightR 
            Caption         =   "58"
            Height          =   195
            Index           =   0
            Left            =   1200
            TabIndex        =   67
            Top             =   2530
            Width           =   315
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Packers (0-80)"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   66
            Top             =   375
            Width           =   1035
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fast Rebound (0-8)"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   65
            Top             =   1095
            Width           =   1365
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Slow Bump (0-24)"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   64
            Top             =   1455
            Width           =   1245
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Ride Height"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   63
            Top             =   2540
            Width           =   840
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Damper Fast Bump (0-8)"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   62
            Top             =   735
            Width           =   1710
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Slow Rebound (0-24)"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   1815
            Width           =   1500
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Spring"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   60
            Top             =   2180
            Width           =   450
         End
      End
      Begin VB.Frame fraRightRead 
         Caption         =   "Right Rear"
         Height          =   2895
         Left            =   3000
         TabIndex        =   36
         Top             =   0
         Width           =   2895
         Begin VB.ComboBox cboSpringR 
            Height          =   315
            Index           =   1
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar hscHeightR 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   1560
            Max             =   160
            Min             =   40
            TabIndex        =   42
            Top             =   2520
            Value           =   116
            Width           =   1215
         End
         Begin VB.TextBox txtSlowReboundR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   41
            Text            =   "15"
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtSlowBumpR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   40
            Text            =   "15"
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtFastReboundR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   39
            Text            =   "4"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtFastBumpR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   38
            Text            =   "4"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtPacR 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   37
            Text            =   "0"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblHeightR 
            Caption         =   "58"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   51
            Top             =   2530
            Width           =   315
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Packers (0-80)"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   50
            Top             =   375
            Width           =   1035
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fast Rebound (0-8)"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   49
            Top             =   1095
            Width           =   1365
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Slow Bump (0-24)"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   1455
            Width           =   1245
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Ride Height"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   2540
            Width           =   840
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Damper Fast Bump (0-8)"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   46
            Top             =   735
            Width           =   1710
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Slow Rebound (0-24)"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Top             =   1815
            Width           =   1500
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Spring"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   2180
            Width           =   450
         End
      End
   End
   Begin VB.Frame fraFrontSettings 
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   -10000
      TabIndex        =   2
      Top             =   480
      Width           =   5895
      Begin VB.Frame fraRightFront 
         Caption         =   "Right Front"
         Height          =   2895
         Left            =   3000
         TabIndex        =   19
         Top             =   0
         Width           =   2895
         Begin VB.ComboBox cboSpringF 
            Height          =   315
            Index           =   1
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar hscHeightF 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   1560
            Max             =   100
            Min             =   30
            TabIndex        =   25
            Top             =   2520
            Value           =   64
            Width           =   1215
         End
         Begin VB.TextBox txtSlowReboundF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   24
            Text            =   "21"
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtSlowBumpF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   23
            Text            =   "21"
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtFastReboundF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   22
            Text            =   "4"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtFastBumpF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   21
            Text            =   "4"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtPacF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   20
            Text            =   "0"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblHeightF 
            Caption         =   "32"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   34
            Top             =   2535
            Width           =   315
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Packers (0-40)"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   375
            Width           =   1035
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fast Rebound (0-8)"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1095
            Width           =   1365
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Slow Bump (0-24)"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   1455
            Width           =   1245
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Ride Height"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   2540
            Width           =   840
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Damper Fast Bump (0-8)"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   735
            Width           =   1710
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Slow Rebound (0-24)"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   1815
            Width           =   1500
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Spring"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   2180
            Width           =   450
         End
      End
      Begin VB.Frame fraLeftFront 
         Caption         =   "Left Front"
         Height          =   2895
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2895
         Begin VB.ComboBox cboSpringF 
            Height          =   315
            Index           =   0
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar hscHeightF 
            Height          =   255
            Index           =   0
            LargeChange     =   10
            Left            =   1560
            Max             =   100
            Min             =   30
            TabIndex        =   9
            Top             =   2520
            Value           =   64
            Width           =   1215
         End
         Begin VB.TextBox txtPacF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   8
            Text            =   "0"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox txtFastBumpF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   7
            Text            =   "4"
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox txtFastReboundF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   1
            TabIndex        =   6
            Text            =   "4"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox txtSlowBumpF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "21"
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox txtSlowReboundF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2400
            MaxLength       =   2
            TabIndex        =   4
            Text            =   "21"
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label lblHeightF 
            Caption         =   "32"
            Height          =   195
            Index           =   0
            Left            =   1200
            TabIndex        =   18
            Top             =   2530
            Width           =   315
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Spring"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   2180
            Width           =   450
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Slow Rebound (0-24)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   1815
            Width           =   1500
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Damper Fast Bump (0-8)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   735
            Width           =   1710
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Ride Height"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   2540
            Width           =   840
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Slow Bump (0-24)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   1455
            Width           =   1245
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fast Rebound (0-8)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   1095
            Width           =   1365
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Packers (0-40)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   375
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame4 
      Height          =   30
      Left            =   -100
      TabIndex        =   0
      Top             =   0
      Width           =   10000
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   7223
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc Settings"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Front Settings"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Rear Settings"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveRQ 
         Caption         =   "Save as Qual And Race"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQual 
         Caption         =   "Save as Qual"
      End
      Begin VB.Menu mnuRace 
         Caption         =   "Save as Race"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuSync 
         Caption         =   "Symmetrical Editing"
         Checked         =   -1  'True
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset all"
         Shortcut        =   +{DEL}
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sync As Boolean

Private Sub cboSpringF_Click(Index As Integer)
    If Sync = True Then
        If Index = 1 Then
            cboSpringF(Index - 1).ListIndex = cboSpringF(Index).ListIndex
        Else
            cboSpringF(Index + 1).ListIndex = cboSpringF(Index).ListIndex
        End If
    End If
End Sub

Private Sub cboSpringR_Click(Index As Integer)
    If Sync = True Then
        If Index = 1 Then
            cboSpringR(Index - 1).ListIndex = cboSpringR(Index).ListIndex
        Else
            cboSpringR(Index + 1).ListIndex = cboSpringR(Index).ListIndex
        End If
    End If
End Sub

Private Sub Form_Load()
    fraMisc.Left = 200
    fraFrontSettings.Left = -10000
    fraRear.Left = -10000
    Sync = False
    Reset
    Sync = True
    SetAllToNr
    frmSetup.Caption = frmSetup.Caption & frmMain.Tag
End Sub

Public Sub Reset()
    cboRollR.AddItem "0"
    cboRollR.AddItem "100"
    cboRollR.AddItem "150"
    cboRollR.AddItem "200"
    cboRollR.AddItem "300"
    cboRollR.AddItem "400"
    cboRollR.AddItem "500"
    cboRollR.AddItem "550"
    cboRollR.AddItem "750"
    cboRollR.AddItem "1000"
    cboRollR.AddItem "1250"
    cboRollR.Text = "100"

    cboRollF.AddItem "0"
    cboRollF.AddItem "500"
    cboRollF.AddItem "1000"
    cboRollF.AddItem "1500"
    cboRollF.AddItem "2000"
    cboRollF.AddItem "3000"
    cboRollF.AddItem "4000"
    cboRollF.AddItem "5500"
    cboRollF.AddItem "7500"
    cboRollF.AddItem "10000"
    cboRollF.AddItem "12500"
    cboRollF.Text = "3000"
    
    For X = 0 To 1
        cboSpringF(X).AddItem "800"
        cboSpringF(X).AddItem "900"
        cboSpringF(X).AddItem "1000"
        cboSpringF(X).AddItem "1100"
        cboSpringF(X).AddItem "1200"
        cboSpringF(X).AddItem "1300"
        cboSpringF(X).AddItem "1400"
        cboSpringF(X).AddItem "1500"
        cboSpringF(X).AddItem "1600"
        cboSpringF(X).Text = "1400"

        cboSpringR(X).AddItem "600"
        cboSpringR(X).AddItem "700"
        cboSpringR(X).AddItem "800"
        cboSpringR(X).AddItem "900"
        cboSpringR(X).AddItem "1000"
        cboSpringR(X).AddItem "1100"
        cboSpringR(X).AddItem "1200"
        cboSpringR(X).AddItem "1300"
        cboSpringR(X).AddItem "1400"
        cboSpringR(X).Text = "1100"

        txtPacF(X).Text = "0"
        txtPacR(X).Text = "0"
        txtFastBumpF(X).Text = "4"
        txtFastBumpR(X).Text = "4"
        txtFastReboundF(X).Text = "4"
        txtFastReboundR(X).Text = "4"
        txtSlowBumpF(X).Text = "21"
        txtSlowBumpR(X).Text = "15"
        txtSlowReboundF(X).Text = "21"
        txtSlowReboundR(X).Text = "15"
        hscHeightF(X).Value = 64
        hscHeightR(X).Value = 116
    Next

    txtFWing.Text = 10
    txtRWing.Text = 12

    txtGear(0).Text = "28"
    txtGear(1).Text = "35"
    txtGear(2).Text = "42"
    txtGear(3).Text = "48"
    txtGear(4).Text = "54"
    txtGear(5).Text = "59"
    hscBrake.Value = 102
End Sub

Private Sub hscBrake_Change()
    lblFront = hscBrake.Value * 125 + 50000
    lblRear = 100000 - lblFront
End Sub

Private Sub hscBrake_Scroll()
    lblFront = hscBrake.Value * 125 + 50000
    lblRear = 100000 - lblFront
End Sub

Private Sub hscHeightF_Change(Index As Integer)
    lblHeightF(Index).Caption = hscHeightF(Index).Value / 2
    If Sync = True Then
        If Index = 1 Then
            hscHeightF(Index - 1).Value = hscHeightF(Index).Value
        Else
            hscHeightF(Index + 1).Value = hscHeightF(Index).Value
        End If
    End If
End Sub

Private Sub hscHeightF_Scroll(Index As Integer)
    lblHeightF(Index).Caption = hscHeightF(Index).Value / 2
    If Sync = True Then
        If Index = 1 Then
            hscHeightF(Index - 1).Value = hscHeightF(Index).Value
        Else
            hscHeightF(Index + 1).Value = hscHeightF(Index).Value
        End If
    End If
End Sub

Private Sub hscHeightR_Change(Index As Integer)
    lblHeightR(Index).Caption = hscHeightR(Index).Value / 2
    If Sync = True Then
        If Index = 1 Then
            hscHeightR(Index - 1).Value = hscHeightR(Index).Value
        Else
            hscHeightR(Index + 1).Value = hscHeightR(Index).Value
        End If
    End If
End Sub

Private Sub hscHeightR_Scroll(Index As Integer)
    lblHeightR(Index).Caption = hscHeightR(Index).Value / 2
    If Sync = True Then
        If Index = 1 Then
            hscHeightR(Index - 1).Value = hscHeightR(Index).Value
        Else
            hscHeightR(Index + 1).Value = hscHeightR(Index).Value
        End If
    End If
End Sub

Private Sub mnuClose_Click()
    On Error Resume Next
    Kill ProgramDir & "\file\tmpsetup.Rtm"
    Kill ProgramDir & "\file\tmpsetup.Qtm"
    Unload Me
End Sub

Private Sub mnuNew_Click()
    Reset
End Sub

Private Sub mnuOpen_Click()
    Read = oFile.ShowOpen("Gp2 Setup File (*.cs*)|*.cs*|All Files (*.*)|*.*|", Me.hWnd, , "Open CarSetup")
    If Read = "" Then Exit Sub
    OpenSetup Read
End Sub

Private Sub mnuQual_Click()
    SaveSetupFile ProgramDir & "\file\tmpsetup.Qtm"
    AddSetup frmMain.lstFile.SelectedItem.Key, ProgramDir & "\file\tmpsetup.Qtm", Qual
    WriteCheckSum frmMain.lstFile.SelectedItem.Key
    frmSetup.Caption = "Gp2 Setup File Editor - Qual Setup"
End Sub

Private Sub mnuRace_Click()
    SaveSetupFile ProgramDir & "\file\tmpsetup.Rtm"
    AddSetup frmMain.lstFile.SelectedItem.Key, ProgramDir & "\file\tmpsetup.Rtm", Race
    WriteCheckSum frmMain.lstFile.SelectedItem.Key
    frmSetup.Caption = "Gp2 Setup File Editor - Race Setup"
End Sub

Private Sub mnuReset_Click()
    Reset
End Sub

Private Sub mnuSaveRQ_Click()
    SaveSetupFile ProgramDir & "\file\tmpsetup.Rtm"
    AddSetup frmMain.lstFile.SelectedItem.Key, ProgramDir & "\file\tmpsetup.Rtm", Race
    SaveSetupFile ProgramDir & "\file\tmpsetup.Qtm"
    AddSetup frmMain.lstFile.SelectedItem.Key, ProgramDir & "\file\tmpsetup.Qtm", Qual
    WriteCheckSum frmMain.lstFile.SelectedItem.Key
End Sub

Private Sub mnuSync_Click()
    If mnuSync.Checked = True Then
        mnuSync.Checked = False
        Sync = False
    Else
        mnuSync.Checked = True
        Sync = True
    End If
End Sub

Private Sub TabStrip1_Click()
    If TabStrip1.Tabs(1).Selected = True Then
        fraMisc.Left = 200
        fraFrontSettings.Left = -10000
        fraRear.Left = -10000
    ElseIf TabStrip1.Tabs(2).Selected = True Then
        fraFrontSettings.Left = 200
        fraMisc.Left = -10000
        fraRear.Left = -10000
    ElseIf TabStrip1.Tabs(3).Selected = True Then
        fraRear.Left = 200
        fraFrontSettings.Left = -10000
        fraMisc.Left = -10000
    End If
End Sub

Private Sub txtFastBumpF_Change(Index As Integer)
    If txtFastBumpF(Index).Text <> "" Then
        If txtFastBumpF(Index).Text > 8 Then txtFastBumpF(Index).Text = 8
    End If
    If Sync = True Then
        If Index = 1 Then
            txtFastBumpF(Index - 1).Text = txtFastBumpF(Index).Text
        Else
            txtFastBumpF(Index + 1).Text = txtFastBumpF(Index).Text
        End If
    End If
End Sub

Private Sub txtFastBumpF_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtFastBumpF_LostFocus(Index As Integer)
    If txtFastBumpF(Index).Text = "" Then txtFastBumpF(Index).Text = "0"
End Sub

Private Sub txtFastBumpR_Change(Index As Integer)
    If txtFastBumpR(Index).Text <> "" Then
        If txtFastBumpR(Index).Text > 8 Then txtFastBumpR(Index).Text = 8
    End If
    If Sync = True Then
        If Index = 1 Then
            txtFastBumpR(Index - 1).Text = txtFastBumpR(Index).Text
        Else
            txtFastBumpR(Index + 1).Text = txtFastBumpR(Index).Text
        End If
    End If
End Sub

Private Sub txtFastBumpR_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtFastBumpR_LostFocus(Index As Integer)
    If txtFastBumpR(Index).Text = "" Then txtFastBumpR(Index).Text = "0"
End Sub

Private Sub txtFastReboundF_Change(Index As Integer)
    If txtFastReboundF(Index).Text <> "" Then
        If txtFastReboundF(Index).Text > 8 Then txtFastReboundF(Index).Text = 8
    End If
    If Sync = True Then
        If Index = 1 Then
            txtFastReboundF(Index - 1).Text = txtFastReboundF(Index).Text
        Else
            txtFastReboundF(Index + 1).Text = txtFastReboundF(Index).Text
        End If
    End If
End Sub

Private Sub txtFastReboundF_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtFastReboundF_LostFocus(Index As Integer)
    If txtFastReboundF(Index).Text = "" Then txtFastReboundF(Index).Text = "0"
End Sub

Private Sub txtFastReboundR_Change(Index As Integer)
    If txtFastReboundR(Index).Text <> "" Then
        If txtFastReboundR(Index).Text > 8 Then txtFastReboundR(Index).Text = 8
    End If
    If Sync = True Then
        If Index = 1 Then
            txtFastReboundR(Index - 1).Text = txtFastReboundR(Index).Text
        Else
            txtFastReboundR(Index + 1).Text = txtFastReboundR(Index).Text
        End If
    End If
End Sub

Private Sub txtFastReboundR_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtFastReboundR_LostFocus(Index As Integer)
    If txtFastReboundR(Index).Text = "" Then txtFastReboundR(Index).Text = "0"
End Sub

Private Sub txtFWing_Change()
    If txtFWing <> "" Then
        If txtFWing.Text > 20 Then txtFWing.Text = 20
        If txtFWing.Text < 1 Then txtFWing.Text = 1
    End If
End Sub

Private Sub txtFWing_GotFocus()
    TextSelected
End Sub

Private Sub txtFWing_LostFocus()
    If txtFWing.Text = "" Then txtFWing.Text = "1"
End Sub

Private Sub txtGear_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtGear_LostFocus(Index As Integer)
    If txtGear(Index).Text = "" Then txtGear(Index).Text = 16 + Index
    If txtGear(Index).Text > Index + 75 Then txtGear(Index).Text = Index + 75
    If txtGear(Index).Text < Index + 16 Then txtGear(Index).Text = Index + 16
    If Index <> 0 Then If txtGear(Index).Text <= txtGear(Index - 1) Then txtGear(Index) = txtGear(Index - 1) + 1
End Sub

Private Sub txtPacF_Change(Index As Integer)
    If txtPacF(Index).Text <> "" Then
        If txtPacF(Index).Text > 40 Then txtPacF(Index).Text = 40
        If txtPacF(Index).Text < 0 Then txtPacF(Index).Text = 0
    End If
    If Sync = True Then
        If Index = 1 Then
            txtPacF(Index - 1).Text = txtPacF(Index).Text
        Else
            txtPacF(Index + 1).Text = txtPacF(Index).Text
        End If
    End If
End Sub

Private Sub txtPacF_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtPacF_LostFocus(Index As Integer)
    If txtPacF(Index).Text = "" Then txtPacF(Index).Text = "0"
End Sub

Private Sub txtPacR_Change(Index As Integer)
    If txtPacR(Index).Text <> "" Then
        If txtPacR(Index).Text > 80 Then txtPacR(Index).Text = 80
        If txtPacR(Index).Text < 0 Then txtPacR(Index).Text = 0
    End If
    If Sync = True Then
        If Index = 1 Then
            txtPacR(Index - 1).Text = txtPacR(Index).Text
        Else
            txtPacR(Index + 1).Text = txtPacR(Index).Text
        End If
    End If
End Sub

Private Sub txtPacR_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtPacR_LostFocus(Index As Integer)
    If txtPacR(Index).Text = "" Then txtPacR(Index).Text = "0"
End Sub

Private Sub txtRWing_Change()
    If txtRWing <> "" Then
        If txtRWing.Text > 20 Then txtRWing.Text = 20
        If txtRWing.Text < 1 Then txtRWing.Text = 1
    End If
End Sub

Private Sub txtRWing_GotFocus()
    TextSelected
End Sub

Private Sub txtRWing_LostFocus()
    If txtRWing.Text = "" Then txtRWing.Text = "1"
End Sub

Private Sub txtSlowBumpF_Change(Index As Integer)
    If txtSlowBumpF(Index).Text <> "" Then
        If txtSlowBumpF(Index).Text > 24 Then txtSlowBumpF(Index).Text = 24
    End If
    If Sync = True Then
        If Index = 1 Then
            txtSlowBumpF(Index - 1).Text = txtSlowBumpF(Index).Text
        Else
            txtSlowBumpF(Index + 1).Text = txtSlowBumpF(Index).Text
        End If
    End If
End Sub

Private Sub txtSlowBumpF_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtSlowBumpF_LostFocus(Index As Integer)
    If txtSlowBumpF(Index).Text = "" Then txtSlowBumpF(Index).Text = "0"
End Sub

Private Sub txtSlowBumpR_Change(Index As Integer)
    If txtSlowBumpR(Index).Text <> "" Then
        If txtSlowBumpR(Index).Text > 24 Then txtSlowBumpR(Index).Text = 24
    End If
    If Sync = True Then
        If Index = 1 Then
            txtSlowBumpR(Index - 1).Text = txtSlowBumpR(Index).Text
        Else
            txtSlowBumpR(Index + 1).Text = txtSlowBumpR(Index).Text
        End If
    End If
End Sub

Private Sub txtSlowBumpR_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtSlowBumpR_LostFocus(Index As Integer)
    If txtSlowBumpR(Index).Text = "" Then txtSlowBumpR(Index).Text = "0"
End Sub

Private Sub txtSlowReboundF_Change(Index As Integer)
    If txtSlowReboundF(Index).Text <> "" Then
        If txtSlowReboundF(Index).Text > 24 Then txtSlowReboundF(Index).Text = 24
    End If
    If Sync = True Then
        If Index = 1 Then
            txtSlowReboundF(Index - 1).Text = txtSlowReboundF(Index).Text
        Else
            txtSlowReboundF(Index + 1).Text = txtSlowReboundF(Index).Text
        End If
    End If
End Sub

Private Sub txtSlowReboundF_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub txtSlowReboundF_LostFocus(Index As Integer)
    If txtSlowReboundF(Index).Text = "" Then txtSlowReboundF(Index).Text = "0"
End Sub

Private Sub txtSlowReboundR_Change(Index As Integer)
    If txtSlowReboundR(Index).Text <> "" Then
        If txtSlowReboundR(Index).Text > 24 Then txtSlowReboundR(Index).Text = 24
    End If
    If Sync = True Then
        If Index = 1 Then
            txtSlowReboundR(Index - 1).Text = txtSlowReboundR(Index).Text
        Else
            txtSlowReboundR(Index + 1).Text = txtSlowReboundR(Index).Text
        End If
    End If
End Sub

Private Sub txtSlowReboundR_GotFocus(Index As Integer)
    TextSelected
End Sub

Public Sub SetAllToNr()
Dim oCtl As Control
    For Each oCtl In frmSetup.Controls
        If TypeOf oCtl Is TextBox Then
            X = GetWindowLong(oCtl.hWnd, GWL_STYLE)
            X = X Or ES_NUMBER
            Call SetWindowLong(oCtl.hWnd, GWL_STYLE, X)
        End If
    Next
    Set oCtl = Nothing
End Sub

Private Sub txtSlowReboundR_LostFocus(Index As Integer)
    If txtSlowReboundR(Index).Text = "" Then txtSlowReboundR(Index).Text = "0"
End Sub
