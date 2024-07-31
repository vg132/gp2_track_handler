VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "GP2 Track Handler v1.4"
   ClientHeight    =   9225
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13890
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   MousePointer    =   7  'Size N S
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Import"
            Object.ToolTipText     =   "Import from GP2"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Export"
            Object.ToolTipText     =   "Export to GP2"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GP2Edit"
            Object.ToolTipText     =   "Add a GP2Edit EXE file"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Database"
            Object.ToolTipText     =   "View Lap Time Database"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   15
         Left            =   3600
         TabIndex        =   113
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox picTree 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8550
      Left            =   0
      ScaleHeight     =   8520
      ScaleWidth      =   3450
      TabIndex        =   3
      Top             =   420
      Width           =   3480
      Begin ComctlLib.TreeView TreeView1 
         Height          =   5895
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   10398
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlTree"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComctlLib.ImageList imlTree 
         Left            =   2640
         Top             =   5880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   4
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":085C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":0DAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIForm1.frx":0EC0
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picData 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8550
      Left            =   3630
      ScaleHeight     =   8520
      ScaleWidth      =   10230
      TabIndex        =   2
      Top             =   420
      Width           =   10260
      Begin VB.Frame frameTrackInfo 
         Caption         =   "Track Info"
         Height          =   3855
         Left            =   4560
         TabIndex        =   94
         Top             =   120
         Width           =   3975
         Begin VB.Label lblEvent 
            AutoSize        =   -1  'True
            Caption         =   "Event: "
            Height          =   195
            Left            =   120
            TabIndex        =   112
            Top             =   720
            Width           =   510
         End
         Begin VB.Label lblAuthor 
            AutoSize        =   -1  'True
            Caption         =   "Author: "
            Height          =   195
            Left            =   120
            TabIndex        =   111
            Top             =   2400
            Width           =   555
         End
         Begin VB.Label lblW 
            AutoSize        =   -1  'True
            Caption         =   "Tyre ware: "
            Height          =   195
            Left            =   120
            TabIndex        =   110
            Top             =   1680
            Width           =   795
         End
         Begin VB.Label lblLe 
            AutoSize        =   -1  'True
            Caption         =   "Length (m): "
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label lblL 
            AutoSize        =   -1  'True
            Caption         =   "Laps: "
            Height          =   195
            Left            =   120
            TabIndex        =   108
            Top             =   1200
            Width           =   435
         End
         Begin VB.Label lblC 
            AutoSize        =   -1  'True
            Caption         =   "Country:"
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   480
            Width           =   585
         End
         Begin VB.Label lblTrackName 
            AutoSize        =   -1  'True
            Caption         =   "Track Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   930
         End
         Begin VB.Label lblYear 
            AutoSize        =   -1  'True
            Caption         =   "Year: "
            Height          =   195
            Left            =   120
            TabIndex        =   105
            Top             =   960
            Width           =   420
         End
         Begin VB.Label lblMisc 
            Caption         =   "Misc Info: "
            Height          =   1035
            Left            =   120
            TabIndex        =   104
            Top             =   2640
            Width           =   3735
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblR 
            AutoSize        =   -1  'True
            Caption         =   "Race Lap Record:"
            Height          =   195
            Left            =   120
            TabIndex        =   103
            Top             =   2160
            Width           =   1320
         End
         Begin VB.Label lblQ 
            AutoSize        =   -1  'True
            Caption         =   "Qual Lap Record:"
            Height          =   195
            Left            =   120
            TabIndex        =   102
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label lblCountry 
            Height          =   195
            Left            =   720
            TabIndex        =   101
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblLaps 
            Height          =   195
            Left            =   600
            TabIndex        =   100
            Top             =   1200
            Width           =   3135
         End
         Begin VB.Label lblLength 
            Height          =   195
            Left            =   960
            TabIndex        =   99
            Top             =   1440
            Width           =   2895
         End
         Begin VB.Label lblWare 
            Height          =   195
            Left            =   960
            TabIndex        =   98
            Top             =   1680
            Width           =   2775
         End
         Begin VB.Label lblQLap 
            Height          =   195
            Left            =   1440
            TabIndex        =   97
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label lblRLap 
            Height          =   195
            Left            =   1560
            TabIndex        =   96
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblName 
            Height          =   195
            Left            =   1080
            TabIndex        =   95
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame frmaeTime 
         Caption         =   "Lap Time Data"
         Height          =   1935
         Left            =   120
         TabIndex        =   33
         Top             =   5280
         Width           =   5535
         Begin VB.TextBox txtQTime 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   44
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtQDriver 
            Height          =   285
            Left            =   1080
            MaxLength       =   23
            TabIndex        =   43
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtQTeam 
            Height          =   285
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   42
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtQDate 
            Height          =   285
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   41
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox txtRTime 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   40
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtRDriver 
            Height          =   285
            Left            =   1080
            MaxLength       =   23
            TabIndex        =   39
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtRTeam 
            Height          =   285
            Left            =   3120
            MaxLength       =   12
            TabIndex        =   38
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtRDate 
            Height          =   285
            Left            =   4320
            MaxLength       =   10
            TabIndex        =   37
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add Time"
            Height          =   325
            Left            =   4320
            TabIndex        =   36
            Top             =   1440
            Width           =   1000
         End
         Begin VB.CommandButton cmdGet 
            Caption         =   "&View Best Lap Database"
            Height          =   325
            Left            =   2160
            TabIndex        =   35
            Top             =   1440
            Width           =   1965
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Info"
            Height          =   325
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Race Time"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Date"
            Height          =   195
            Left            =   4320
            TabIndex        =   51
            Top             =   840
            Width           =   345
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Team"
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   50
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Driver"
            Height          =   195
            Index           =   0
            Left            =   1080
            TabIndex        =   49
            Top             =   840
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Qual Time"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Driver"
            Height          =   195
            Left            =   1080
            TabIndex        =   47
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Team"
            Height          =   195
            Left            =   3120
            TabIndex        =   46
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Date"
            Height          =   195
            Left            =   4320
            TabIndex        =   45
            Top             =   240
            Width           =   345
         End
      End
      Begin VB.TextBox txtFullPath 
         Height          =   285
         Left            =   5160
         TabIndex        =   28
         Top             =   9000
         Width           =   1215
         Visible         =   0   'False
      End
      Begin VB.TextBox txtFramedPath 
         Height          =   285
         Left            =   5160
         TabIndex        =   27
         Top             =   8640
         Width           =   1215
         Visible         =   0   'False
      End
      Begin VB.Frame framePlayer 
         Caption         =   "Player Car"
         Height          =   5175
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3135
         Visible         =   0   'False
         Begin VB.CheckBox chkSelectedTeam 
            Caption         =   "Use Selected Team Power"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CheckBox chkNoLimit 
            Caption         =   "No Speed Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   3840
            Width           =   2655
         End
         Begin VB.HScrollBar hscPitSpeed 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   201
            Min             =   1
            TabIndex        =   14
            Top             =   3480
            Value           =   50
            Width           =   2655
         End
         Begin VB.HScrollBar hscPGrip 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   1000
            TabIndex        =   13
            Top             =   2880
            Value           =   198
            Width           =   2655
         End
         Begin VB.HScrollBar hscWeight 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   4000
            Min             =   401
            TabIndex        =   12
            Top             =   2280
            Value           =   1313
            Width           =   2655
         End
         Begin VB.HScrollBar hscPQPower 
            Height          =   255
            LargeChange     =   20
            Left            =   120
            Max             =   1579
            TabIndex        =   11
            Top             =   1200
            Value           =   790
            Width           =   2655
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   120
            TabIndex        =   10
            Top             =   4200
            Width           =   2895
         End
         Begin VB.CommandButton cmdExportSettings 
            Caption         =   "&Export"
            Height          =   325
            Left            =   120
            TabIndex        =   9
            Top             =   4320
            Width           =   1000
         End
         Begin VB.CommandButton cmdImportSettings 
            Caption         =   "&Import"
            Height          =   325
            Left            =   120
            TabIndex        =   8
            Top             =   4680
            Width           =   1000
         End
         Begin VB.CommandButton cmdDefaultSettings 
            Caption         =   "GP2 Default"
            Height          =   325
            Left            =   1170
            TabIndex        =   7
            Top             =   4320
            Width           =   1100
         End
         Begin VB.HScrollBar hscPRPower 
            Height          =   255
            LargeChange     =   20
            Left            =   120
            Max             =   1579
            TabIndex        =   6
            Top             =   600
            Value           =   780
            Width           =   2655
         End
         Begin VB.Label lblPitSpeed 
            Alignment       =   1  'Right Justify
            Caption         =   "50mph (80km/h)"
            Height          =   195
            Left            =   1320
            TabIndex        =   26
            Top             =   3240
            Width           =   1440
         End
         Begin VB.Label lblPit 
            AutoSize        =   -1  'True
            Caption         =   "Pit Speed Limit"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   3240
            Width           =   1050
         End
         Begin VB.Label lblGrip 
            Alignment       =   1  'Right Justify
            Caption         =   "198"
            Height          =   195
            Left            =   2280
            TabIndex        =   24
            Top             =   2640
            Width           =   480
         End
         Begin VB.Label lblGrip2 
            AutoSize        =   -1  'True
            Caption         =   "Grip"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   2640
            Width           =   285
         End
         Begin VB.Label lblWeight2 
            Alignment       =   1  'Right Justify
            Caption         =   "1313lb (596Kg)"
            Height          =   195
            Left            =   1560
            TabIndex        =   22
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label lblWeight 
            AutoSize        =   -1  'True
            Caption         =   "Player Car Weight"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   2040
            Width           =   1275
         End
         Begin VB.Label lblPRPower 
            Alignment       =   1  'Right Justify
            Caption         =   "780"
            Height          =   195
            Left            =   2295
            TabIndex        =   20
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Power in Qual"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   990
         End
         Begin VB.Label lblPQPower 
            Alignment       =   1  'Right Justify
            Caption         =   "790"
            Height          =   195
            Left            =   2280
            TabIndex        =   18
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Power in Race"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1050
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4680
         Top             =   7320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame frameData 
         Caption         =   "Track Data"
         Height          =   2175
         Left            =   120
         TabIndex        =   53
         Top             =   3000
         Width           =   4215
         Begin VB.TextBox txtLaps 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   57
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtLength 
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   56
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtTire 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   55
            Top             =   1680
            Width           =   735
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   285
            Left            =   615
            Max             =   3
            Min             =   126
            TabIndex        =   54
            Top             =   480
            Value           =   3
            Width           =   200
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Laps (3-126)"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Track Length (0-9999 m)"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tyre ware (14848-37887)"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   58
            Top             =   1440
            Width           =   1785
         End
      End
      Begin VB.Frame frameInfo 
         Caption         =   "Track Info"
         Height          =   2775
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   4215
         Begin VB.TextBox txtPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   255
            TabIndex        =   66
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   120
            TabIndex        =   65
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtCountry 
            Height          =   285
            Left            =   120
            TabIndex        =   64
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox txtAdjectiv 
            Height          =   285
            Left            =   120
            TabIndex        =   63
            Top             =   2280
            Width           =   2775
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse"
            Height          =   285
            Left            =   3120
            TabIndex        =   62
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Adjective:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   70
            Top             =   2040
            Width           =   705
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Country:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   69
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Track Name:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Track Path"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame frameGlobal 
         Caption         =   "Global Settings"
         Height          =   5175
         Left            =   3360
         TabIndex        =   71
         Top             =   120
         Width           =   3135
         Visible         =   0   'False
         Begin VB.CheckBox chkSave 
            Caption         =   "Always save track record"
            Height          =   255
            Left            =   240
            TabIndex        =   76
            Top             =   3120
            Width           =   2175
         End
         Begin VB.CheckBox chk0as1 
            Caption         =   "Show car 1 as 0"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   2760
            Width           =   2055
         End
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   74
            Top             =   1200
            Value           =   5
            Width           =   2655
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   4000
            Min             =   401
            TabIndex        =   73
            Top             =   600
            Value           =   1313
            Width           =   2655
         End
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   120
            TabIndex        =   72
            Top             =   3600
            Width           =   2895
         End
         Begin ComctlLib.Slider Slider1 
            Height          =   375
            Left            =   240
            TabIndex        =   77
            Top             =   2280
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   327682
            LargeChange     =   10
            Min             =   1900
            Max             =   2099
            SelStart        =   1994
            TickFrequency   =   10
            Value           =   1994
         End
         Begin VB.Label lblYear2 
            AutoSize        =   -1  'True
            Caption         =   "1994"
            Height          =   195
            Left            =   1800
            TabIndex        =   89
            Top             =   1920
            Width           =   360
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Year for this Season:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   88
            Top             =   1920
            Width           =   1470
         End
         Begin VB.Label lblQuick 
            Alignment       =   1  'Right Justify
            Caption         =   "5%"
            Height          =   195
            Left            =   2400
            TabIndex        =   87
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Quick Race Length"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   86
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lblCWeight 
            Alignment       =   1  'Right Justify
            Caption         =   "1313lb (596kg)"
            Height          =   195
            Left            =   1800
            TabIndex        =   85
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Computer Car Weight"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   84
            Top             =   360
            Width           =   1515
         End
         Begin VB.Image S 
            Height          =   180
            Index           =   0
            Left            =   240
            Top             =   4320
            Width           =   195
         End
         Begin VB.Image S 
            Height          =   180
            Index           =   6
            Left            =   1410
            Top             =   4320
            Width           =   195
         End
         Begin VB.Image S 
            Height          =   180
            Index           =   5
            Left            =   1215
            Top             =   4320
            Width           =   195
         End
         Begin VB.Image S 
            Height          =   180
            Index           =   4
            Left            =   1020
            Top             =   4320
            Width           =   195
         End
         Begin VB.Image S 
            Height          =   180
            Index           =   3
            Left            =   825
            Top             =   4320
            Width           =   195
         End
         Begin VB.Image S 
            Height          =   180
            Index           =   2
            Left            =   630
            Top             =   4320
            Width           =   195
         End
         Begin VB.Image S 
            Height          =   180
            Index           =   1
            Left            =   435
            Top             =   4320
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   0
            Left            =   240
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   0
            Left            =   240
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   6
            Left            =   1410
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   6
            Left            =   1410
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   5
            Left            =   1215
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   5
            Left            =   1215
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   4
            Left            =   1020
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   4
            Left            =   1020
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   3
            Left            =   825
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   3
            Left            =   825
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   2
            Left            =   630
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   2
            Left            =   630
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   1
            Left            =   435
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   1
            Left            =   435
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image A 
            Height          =   180
            Index           =   6
            Left            =   1410
            Top             =   4080
            Width           =   195
         End
         Begin VB.Image A 
            Height          =   180
            Index           =   5
            Left            =   1215
            Top             =   4080
            Width           =   195
         End
         Begin VB.Image A 
            Height          =   180
            Index           =   4
            Left            =   1020
            Top             =   4080
            Width           =   195
         End
         Begin VB.Image A 
            Height          =   180
            Index           =   3
            Left            =   825
            Top             =   4080
            Width           =   195
         End
         Begin VB.Image A 
            Height          =   180
            Index           =   2
            Left            =   630
            Top             =   4080
            Width           =   195
         End
         Begin VB.Image A 
            Height          =   180
            Index           =   1
            Left            =   435
            Top             =   4080
            Width           =   195
         End
         Begin VB.Image A 
            Height          =   180
            Index           =   0
            Left            =   240
            Top             =   4080
            Width           =   195
         End
         Begin VB.Image R 
            Height          =   180
            Index           =   6
            Left            =   1410
            Top             =   3840
            Width           =   195
         End
         Begin VB.Image R 
            Height          =   180
            Index           =   5
            Left            =   1215
            Top             =   3840
            Width           =   195
         End
         Begin VB.Image R 
            Height          =   180
            Index           =   4
            Left            =   1020
            Top             =   3840
            Width           =   195
         End
         Begin VB.Image R 
            Height          =   180
            Index           =   3
            Left            =   825
            Top             =   3840
            Width           =   195
         End
         Begin VB.Image R 
            Height          =   180
            Index           =   2
            Left            =   630
            Top             =   3840
            Width           =   195
         End
         Begin VB.Image R 
            Height          =   180
            Index           =   1
            Left            =   435
            Top             =   3840
            Width           =   195
         End
         Begin VB.Image R 
            Height          =   180
            Index           =   0
            Left            =   240
            Top             =   3840
            Width           =   195
         End
         Begin VB.Label lblAce 
            AutoSize        =   -1  'True
            Caption         =   "Ace"
            Height          =   180
            Left            =   1800
            TabIndex        =   83
            Top             =   4800
            Width           =   285
         End
         Begin VB.Label lblRookie 
            AutoSize        =   -1  'True
            Caption         =   "Rookie"
            Height          =   180
            Index           =   3
            Left            =   1800
            TabIndex        =   82
            Top             =   3840
            Width           =   510
         End
         Begin VB.Label lblAmateur 
            AutoSize        =   -1  'True
            Caption         =   "Amateur"
            Height          =   180
            Left            =   1800
            TabIndex        =   81
            Top             =   4080
            Width           =   585
         End
         Begin VB.Label lblSemiPro 
            AutoSize        =   -1  'True
            Caption         =   "Semi-Pro"
            Height          =   180
            Left            =   1800
            TabIndex        =   80
            Top             =   4320
            Width           =   630
         End
         Begin VB.Label lblPro 
            AutoSize        =   -1  'True
            Caption         =   "Pro"
            Height          =   180
            Left            =   1800
            TabIndex        =   79
            Top             =   4560
            Width           =   240
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "<-- Click"
            Height          =   195
            Left            =   2280
            TabIndex        =   78
            Top             =   1920
            Width           =   570
         End
      End
      Begin VB.Frame frameFile 
         Caption         =   " Select Track/Menu Picture "
         Height          =   3855
         Left            =   6120
         TabIndex        =   29
         Top             =   4080
         Width           =   4215
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   2160
            TabIndex        =   32
            Top             =   3390
            Width           =   1815
         End
         Begin VB.DirListBox Dir1 
            Height          =   3015
            Left            =   2160
            TabIndex        =   31
            Top             =   300
            Width           =   1815
         End
         Begin VB.FileListBox File1 
            Height          =   3405
            Left            =   120
            Pattern         =   "*.dat;*.gif*;*.bmp"
            TabIndex        =   30
            Top             =   300
            Width           =   1935
         End
      End
      Begin VB.Label lblFull 
         AutoSize        =   -1  'True
         Caption         =   "Fullview Picture"
         Height          =   195
         Left            =   4560
         TabIndex        =   93
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label lblFramed 
         AutoSize        =   -1  'True
         Caption         =   "FrameView Picture"
         Height          =   195
         Left            =   4560
         TabIndex        =   92
         Top             =   2640
         Width           =   1320
         Visible         =   0   'False
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   240
         TabIndex        =   91
         Top             =   7800
         Width           =   855
      End
      Begin VB.Image imgPre 
         Height          =   1455
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Image imgSize 
         Height          =   495
         Left            =   6240
         Top             =   5760
         Width           =   1215
         Visible         =   0   'False
      End
      Begin VB.Image imgFull 
         Height          =   2100
         Left            =   4560
         Stretch         =   -1  'True
         Top             =   480
         Width           =   2700
      End
      Begin VB.Image imgFramed 
         Height          =   2100
         Left            =   4560
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   2700
      End
      Begin VB.Image Off 
         Height          =   180
         Index           =   1
         Left            =   7920
         Picture         =   "MDIForm1.frx":0FD2
         Top             =   5400
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image Off 
         Height          =   180
         Index           =   2
         Left            =   8160
         Picture         =   "MDIForm1.frx":1339
         Top             =   5400
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image Off 
         Height          =   180
         Index           =   3
         Left            =   8400
         Picture         =   "MDIForm1.frx":169A
         Top             =   5400
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image Off 
         Height          =   180
         Index           =   4
         Left            =   8640
         Picture         =   "MDIForm1.frx":19FC
         Top             =   5400
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image Off 
         Height          =   180
         Index           =   5
         Left            =   8880
         Picture         =   "MDIForm1.frx":1D5B
         Top             =   5400
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image Off 
         Height          =   180
         Index           =   6
         Left            =   9120
         Picture         =   "MDIForm1.frx":20BE
         Top             =   5400
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image On1 
         Height          =   180
         Index           =   3
         Left            =   8400
         Picture         =   "MDIForm1.frx":2410
         Top             =   5160
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image On1 
         Height          =   180
         Index           =   4
         Left            =   8640
         Picture         =   "MDIForm1.frx":2769
         Top             =   5160
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image On1 
         Height          =   180
         Index           =   5
         Left            =   8880
         Picture         =   "MDIForm1.frx":2AB9
         Top             =   5160
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image On1 
         Height          =   180
         Index           =   6
         Left            =   9120
         Picture         =   "MDIForm1.frx":2E12
         Top             =   5160
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image On1 
         Height          =   180
         Index           =   0
         Left            =   7680
         Picture         =   "MDIForm1.frx":3172
         Top             =   5160
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image Off 
         Height          =   180
         Index           =   0
         Left            =   7680
         Picture         =   "MDIForm1.frx":34D5
         Top             =   5400
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image On1 
         Height          =   180
         Index           =   1
         Left            =   7920
         Picture         =   "MDIForm1.frx":383E
         Top             =   5160
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Image On1 
         Height          =   180
         Index           =   2
         Left            =   8160
         Picture         =   "MDIForm1.frx":3BA5
         Top             =   5160
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.Label lblNote 
         Caption         =   $"MDIForm1.frx":3F00
         Height          =   615
         Left            =   120
         TabIndex        =   90
         Top             =   5400
         Width           =   5055
         Visible         =   0   'False
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8970
      Width           =   13890
      _ExtentX        =   24500
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "13:21"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "1999-05-19"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3FD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":40E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":41F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4307
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4621
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":493B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4F6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":5081
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":539B
            Key             =   ""
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
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Sav&e as..."
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "File &Converter..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGP2 
         Caption         =   "&Import from GP2"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuToGP2 
         Caption         =   "&Export to GP2"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen1 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpen2 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpen3 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuPoint 
         Caption         =   "&Point Editor..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuDosPath 
         Caption         =   "&Add GP2Edit Carset File (EXE)..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRnd 
         Caption         =   "&Random Season"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSet 
         Caption         =   "&Set all times to..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuDatabase 
         Caption         =   "&Best Lap Database..."
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Options"
      Begin VB.Menu mnuTrackDir 
         Caption         =   "&Track Directory..."
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&GP2 Location..."
      End
   End
   Begin VB.Menu mnuHelpMenu 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "Track Handler on the &Web"
         Begin VB.Menu mnuHomePage 
            Caption         =   "Track Handler HomePage"
         End
         Begin VB.Menu mnuReg 
            Caption         =   "Register as a user (free)"
         End
         Begin VB.Menu mnuBug 
            Caption         =   "Bug Report"
         End
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About GP2 Track Handler"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

' Return codes from Registration functions.
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Sub MDIForm_Load()
    Dim lStyle As Long
    Dim nodX As Node    ' Create variable.
    Set oMisc = New TrackHandler.Misc
    Set oData = New TrackHandler.Data

    'Lås så att bara siffror kan skrivas in i text box!
    lStyle = GetWindowLong(txtLength.hwnd, GWL_STYLE)
    lStyle = lStyle Or ES_NUMBER
    Call SetWindowLong(txtLength.hwnd, GWL_STYLE, lStyle)
    lStyle = GetWindowLong(txtTire.hwnd, GWL_STYLE)
    lStyle = lStyle Or ES_NUMBER
    Call SetWindowLong(txtTire.hwnd, GWL_STYLE, lStyle)
    lStyle = GetWindowLong(txtLaps.hwnd, GWL_STYLE)
    lStyle = lStyle Or ES_NUMBER
    Call SetWindowLong(txtLaps.hwnd, GWL_STYLE, lStyle)

    Set nodX = TreeView1.Nodes.Add(, , "q", "Hej")
    NewTree
    FileInfo.FileType = FileNew
    frameFile.Visible = True
    frameInfo.Visible = False
    frameData.Visible = False
    frmaeTime.Visible = False
    frameFile.Top = frameInfo.Top
    frameFile.Left = frameInfo.Left
    Read4 = ""
    MDIForm1.MousePointer = 11
    GP2TH = "GP2 Track Handler"
    TH = "Track Handler"
    On Error Resume Next
    SelectVersion = False
    CurrentRecord = 0
    Gp2Dir = GetSetting(GP2TH, "Settings", "GP2 Path")
    CountNr = Len(Gp2Dir)
    If Mid(Gp2Dir, CountNr, 1) = "\" Then Gp2Dir = Mid(Gp2Dir, 1, CountNr - 1)
    ProgramDir = "g:" 'App.Path
    If Len(ProgramDir) = 3 And Mid(ProgramDir, 3, 1) = "\" Then ProgramDir = Mid(ProgramDir, 1, 2)
    App.HelpFile = ProgramDir + "\Help.chm"
    Read = oMisc.File_Exists(ProgramDir + "\mall.lda")
    If Read = False Then
        MsgBox GP2TH + " was not able to find the file mall.lda, this file must be in the same directory as " + GP2TH + ". This program will now be terminated.", vbCritical, TH
        End
    End If
    CommonDialog1.InitDir = ProgramDir
    NoSupport = True
    If Gp2Dir <> "" Then GetGP2Version
    CurrentRecord = 0
    CurrentRecord2 = 0

    NewFile
    If Gp2Dir = "" Then frmOptions.Show , MDIForm1

    TrackPath = ""
    RegFileName

    MDIForm1.StatusBar1.Panels(1) = GP2TH & " © Viktor Gars 1998/99"
    MDIForm1.StatusBar1.Panels(2) = "GP2 Version: " + GP2Country
    MDIForm1.StatusBar1.Panels(3) = "GP2 Directory: " + Gp2Dir
    
    DefaultTrackPath = GetSetting(GP2TH, "Settings", "TrackPath")
    If DefaultTrackPath = "" Then DefaultTrackPath = ProgramDir
    File1.Path = DefaultTrackPath
    Drive1.Drive = Mid(DefaultTrackPath, 1, 3)
    Dir1.Path = DefaultTrackPath
    ShowRecent
    MDIForm1.MousePointer = 0
    If Command() <> "" Then
        Read = Command()
        OpenStartFile (Read)
    End If
Exit Sub

ErrorTrap:
    Select Case Err.Number
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
        MDIForm1.MousePointer = 0
    End Select
End Sub

Private Sub MDIForm_Resize()
    Resize_Form
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Count1 = 0
    mnuExit_Click
    If Count1 = 1 Then
        Cancel = -1
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, MDIForm1
End Sub

Private Sub mnuBug_Click()
    Read = "http://hem1.passagen.se/formula1/bugin.htm"
    OpenPage = ShellExecute(MDIForm1.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuConvert_Click()
    frmConverter.Show vbModal, MDIForm1
End Sub

Private Sub mnuDatabase_Click()
    frmBestLap.Show , MDIForm1
End Sub

Private Sub mnuDosPath_Click()
    Read = oMisc.ReadINI("Misc", "EXEPath", ProgramDir + "\WorkCopy.lda")
    If Read = "" Then
        CommonDialog1.CancelError = True
        On Error GoTo ErrorTrap
        CommonDialog1.DialogTitle = "Select GP2 Edit Carset File"
        CommonDialog1.Flags = cdlOFNHideReadOnly
        CommonDialog1.Filter = "GP2 Edit Carset file (*.exe)|*.exe|"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.ShowOpen
        FileNum = FreeFile
        Open CommonDialog1.FileName For Binary As FileNum
        Read = String(38, " ")
        Get #FileNum, 45419, Read
        Close FileNum
        If Read <> "Installer v1.81 - (c)1998 Steven Young" Then
            CountNr = Len(CommonDialog1.FileName)
            Read = Mid(CommonDialog1.FileName, CountNr - 2, 3)
            If Read = "gp2" Then
                Responce = MsgBox(GP2TH + " don't support *.gp2 files, " + GP2TH + " only suppots the GP2Edit Dos Patch files.", vbExclamation + vbRetryCancel, TH)
            Else
                Responce = MsgBox("This is not a GP2Edit Carset file (EXE).", vbExclamation + vbRetryCancel, TH)
            End If
            If Responce = vbRetry Then mnuDosPath_Click
            If Responce = vbCancel Then Exit Sub
        End If
        Read4 = oMisc.WriteINI("Misc", "EXEPath", CommonDialog1.FileName, ProgramDir + "\WorkCopy.lda")
    End If
    frmDosPath.Show , MDIForm1
    Exit Sub

ErrorTrap:
    Select Case Err.Number
    Case "32755"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Private Sub mnuExit_Click()
    On Error Resume Next
    Unload frmExport
    Unload frmImport
    Unload frmDosPath
    Unload frmAbout
    Unload frmConverter
    Unload frmSelect
    DeleteFile ProgramDir + "\WorkCopy.lda"
    DeleteFile Gp2Dir + "\_MenuPic.bat"
    DeleteFile Gp2Dir + "\$$Check$.bat"
    DeleteFile Gp2Dir + "\$$Check.exe"
    End
End Sub

Private Sub mnuGP2_Click()
    frmImport.Show , MDIForm1
End Sub

Private Sub mnuHelp_Click()
    SendKeys "{F1}"
End Sub

Private Sub mnuHomePage_Click()
    Read = "http://hem1.passagen.se/formula1/index.htm"
    OpenPage = ShellExecute(MDIForm1.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuNew_Click()
    If FileInfo.FileType = FileOpen Then
        Dim RetVal As Variant
        RetVal = oMisc.SaveFile(FileInfo.FilePath, ProgramDir, FileInfo.FileType)
        If RetVal = True Then
            Responce = MsgBox("Save changes?", vbYesNoCancel, "Save")
            If Responce = vbCancel Then Exit Sub
            If Responce = vbYes Then
                mnuSave_Click
            End If
        End If
    End If
    MDIForm1.MousePointer = 11
    Read = oMisc.File_Exists(ProgramDir + "\WorkCopy.lda")
    If Read = True Then DeleteFile ProgramDir + "\WorkCopy.lda"
    FileInfo.FileType = FileNew
    NewFile
    LastClick = ""
    CurrentRecord = 0
    MDIForm1.MousePointer = 0
End Sub

Private Sub mnuOpen_Click()
    On Error GoTo ErrorTrap
    MDIForm1.MousePointer = 11
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = GP2TH + " File (*.ths)|*.ths|"
    CommonDialog1.ShowOpen
    
    OpenThFile CommonDialog1.FileName
    oMisc.RecentFile 1, CommonDialog1.FileName, CommonDialog1.FileTitle, GP2TH, SaveNew
    ShowRecent
    MDIForm1.Caption = GP2TH & " v1.4 - " & CommonDialog1.FileTitle

    FileInfo.FileName = CommonDialog1.FileTitle
    FileInfo.FilePath = CommonDialog1.FileName
    FileInfo.FileType = FileOpen

    MDIForm1.MousePointer = 0
Exit Sub

ErrorTrap:
    Select Case Err.Number
    Case "32755"
        MDIForm1.MousePointer = 0
        Exit Sub
    Case "75"
        MDIForm1.MousePointer = 0
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
        MDIForm1.MousePointer = 0
    End Select
End Sub

Private Sub mnuOpen1_Click()
    oMisc.RecentFile 1, "", "", GP2TH, OpenRecent
    OpenThFile mnuOpen1.Tag
    MDIForm1.Caption = GP2TH & " v1.4 - " & mnuOpen1.Caption
    FileInfo.FileName = mnuOpen1.Caption
    FileInfo.FilePath = mnuOpen1.Tag
    FileInfo.FileType = FileOpen
End Sub

Private Sub mnuOpen2_Click()
    oMisc.RecentFile 2, "", "", GP2TH, OpenRecent
    OpenThFile mnuOpen2.Tag
    MDIForm1.Caption = GP2TH & " v1.4 - " & mnuOpen2.Caption
    FileInfo.FileName = mnuOpen2.Caption
    FileInfo.FilePath = mnuOpen2.Tag
    FileInfo.FileType = FileOpen
    ShowRecent
End Sub

Private Sub mnuOpen3_Click()
    oMisc.RecentFile 3, "", "", GP2TH, OpenRecent
    OpenThFile mnuOpen3.Tag
    MDIForm1.Caption = GP2TH & " v1.4 - " & mnuOpen3.Caption
    FileInfo.FileName = mnuOpen3.Caption
    FileInfo.FilePath = mnuOpen3.Tag
    FileInfo.FileType = FileOpen
    ShowRecent
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show , MDIForm1
End Sub

Private Sub mnuPoint_Click()
    frmPoint.Show vbModal, MDIForm1
End Sub

Private Sub mnuReg_Click()
    Read = "http://hem1.passagen.se/formula1/regin.htm"
    OpenPage = ShellExecute(MDIForm1.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuRnd_Click()
    If File1.ListCount > 15 Then
        Rand File1.ListCount
    Else
        MsgBox "You must have 16 track's or more.", vbInformation, TH
    End If
End Sub

Private Sub mnuSave_Click()
    MDIForm1.MousePointer = 11
    SaveLastClick
    If FileInfo.FileType = FileImport Then
        SaveImport
    End If
    If FileInfo.FileType = FileNew Then
        mnuSaveAs_Click
        Exit Sub
    End If
    SaveThFile
    FileInfo.FileType = FileOpen
    MDIForm1.MousePointer = 0
    Exit Sub
ErrorTrap:
    MDIForm1.MousePointer = 0
    Select Case Err.Number
    Case "32755"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Private Sub mnuSaveAs_Click()
    MDIForm1.MousePointer = 11
    SaveLastClick
    If FileInfo.FileType = FileImport Then
        SaveImport
    End If
    On Error GoTo ErrorTrap
    CommonDialog1.Filter = GP2TH + " File (*.ths)|*.ths|"
    CommonDialog1.ShowSave
    FileInfo.FileName = CommonDialog1.FileTitle
    FileInfo.FilePath = CommonDialog1.FileName
    SaveThFile
    ShowRecent
    FileInfo.FileType = FileOpen
    MDIForm1.MousePointer = 0
Exit Sub
ErrorTrap:
    MDIForm1.MousePointer = 0
    Select Case Err.Number
    Case "75"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Public Sub SaveImport()
    MsgBox "You have imported data in this file, please select a destination directory of the imported track's. The track files will have the same name as the track.", vbInformation, TH
    szTitle = "Select Destination directory"
    With tBrowseInfo
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Read = sBuffer
    End If
    If Len(Read) = 3 Then Read = Mid(Read, 1, 2)
    X = 1
    Do Until X > 16
        Read2 = oMisc.ReadINI("Track " + Trim(Str(X)), "TPath", ProgramDir + "\WorkCopy.lda")
        If X > 9 Then
            Read4 = Trim(Str(X))
        Else
            Read4 = "0" + Trim(Str(X))
        End If
        Read3 = Gp2Dir + "\Circuits\f1ct" + Read4 + ".dat"
        If UCase(Read2) = UCase(Read3) Then
            Read4 = oMisc.ReadINI("Track " + Trim(Str(X)), "Name", ProgramDir + "\WorkCopy.lda")
            Read4 = Read + "\" + Read4 + ".dat"
            SourceFile = Read2
            TargetFile = Read4
            FileCopy Read2, Read4
            Read2 = oMisc.WriteINI("Track " + Trim(Str(X)), "TPath", Read4, ProgramDir + "\WorkCopy.lda")
        End If
        X = X + 1
    Loop
    Exit Sub
ErrorTrap:
    Select Case Err.Number
    Case "32755"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Private Sub mnuSet_Click()
    Read = InputBox("Time (#,##,###)", "Set time")
    If Len(Read) = 8 Then
        X = 1
        Do Until X > 16
            Read2 = ProgramDir + "\WorkCopy.lda"
            Read3 = "Track " + Trim(Str(X))
            Read4 = oMisc.WriteINI(Read3, "QDate", Date, Read2)
            Read4 = oMisc.WriteINI(Read3, "RDate", Date, Read2)
            Read4 = oMisc.WriteINI(Read3, "QTeam", "No Team", Read2)
            Read4 = oMisc.WriteINI(Read3, "RTeam", "No Team", Read2)
            Read4 = oMisc.WriteINI(Read3, "QDriver", "No Driver", Read2)
            Read4 = oMisc.WriteINI(Read3, "RDriver", "No Driver", Read2)
            Read4 = oMisc.WriteINI(Read3, "QTime", Read, Read2)
            Read4 = oMisc.WriteINI(Read3, "RTime", Read, Read2)
            X = X + 1
        Loop
        txtQDate = Date
        txtRDate = Date
        txtQDriver = "No Driver"
        txtRDriver = "No Driver"
        txtQTeam = "No Team"
        txtRTeam = "No Team"
        txtQTime = Read
        txtRTime = Read
    Else
        Responce = MsgBox("You must write the time in the right format, #,##,###.", vbRetryCancel, TH)
        If Responce = vbRetry Then mnuSet_Click
    End If
End Sub

Private Sub mnuShowTip_Click()
    'Load frmTip.
    'frmTip.Show , MDIForm1
End Sub

Private Sub mnuToGP2_Click()
    frmExport.Show , MDIForm1
End Sub

Private Sub mnuTrackDir_Click()
    szTitle = "Select Default Track File Location"
    With tBrowseInfo
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        DefaultTrackPath = sBuffer
        SaveSetting GP2TH, "Settings", "TrackPath", DefaultTrackPath
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "Exit"
            mnuExit_Click
        Case "Open"
            mnuOpen_Click
        Case "New"
            mnuNew_Click
        Case "Save"
            mnuSave_Click
        Case "Export"
            frmExport.Show , MDIForm1
        Case "Import"
            frmImport.Show , MDIForm1
        Case "GP2Edit"
            mnuDosPath_Click
        Case "Help"
            mnuHelp_Click
        Case "Database"
            mnuDatabase_Click
        Case "Delete"
            DeleteTrack
    End Select
End Sub

Public Sub OpenStartFile(ByVal FileToOpen As String)
    Read = String(1, " ")
    Read2 = ""
    X = Len(FileToOpen)
    Do Until Read = "\"
        Read = Mid(FileToOpen, X, 1)
        If Read <> "\" Then Read2 = Read + Read2
        X = X - 1
    Loop
    OpenThFile FileToOpen
    oMisc.RecentFile 1, FileToOpen, Read, GP2TH, SaveNew
    ShowRecent
    FileInfo.FileName = Read
    FileInfo.FilePath = FileToOpen
    FileInfo.FileType = FileOpen
    Exit Sub
ErrorTrap:
    Select Case Err.Number
    Case "32755"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Public Sub RegFileName()
Dim sKeyName As String   'Holds Key Name in registry.
Dim sKeyValue As String  'Holds Key Value in registry.
Dim ret&           'Holds error status if any from API calls.
Dim lphKey&        'Holds created key handle from RegCreateKey.
    'This creates a Root entry called "MyApp".
    sKeyName = TH
    sKeyValue = GP2TH + " File"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
    'This creates a Root entry called .BAR associated with "MyApp".
    sKeyName = ".ths"
    sKeyValue = TH
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "", REG_SZ, sKeyValue, 0&)
    'This sets the command line for "MyApp".
    sKeyName = TH
    sKeyValue = ProgramDir + "\" + App.EXEName + ".exe %1"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "Shell\Open\Command", REG_SZ, sKeyValue, MAX_PATH)
    'This sets the Icon for "MyApp".
    sKeyName = TH
    sKeyValue = ProgramDir + "\File.ico"
    ret& = RegCreateKey&(HKEY_CLASSES_ROOT, sKeyName, lphKey&)
    ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH)
End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    On Error GoTo errorstop
    Dim nodX As Node
    
    CountNr = Len(File1.FileName)
    Read = UCase(Mid(File1.FileName, CountNr - 2, 3))
    If (Read = "DAT") And (Len(TreeView1.DropHighlight.Key) = 3) Then
        Count3 = Mid(TreeView1.DropHighlight.Key, 2, 2) - 10
        GetDataFromLabel
        TreeView1.DropHighlight.Text = Trim(Str(Count3)) + ". " + lblName
        Read = oMisc.ReadINI("Track " + Trim(Str(Count3)), "TPath", ProgramDir + "\WorkCopy.lda")
        Count2 = 90
        Read2 = "DAT"
        Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 100)), "Track File: " + Read, 4, 4)
        TreeView1.DropHighlight = Nothing
        Exit Sub
    End If
    If ((Read = "BMP") Or (Read = "GIF")) And (Len(TreeView1.DropHighlight.Key) = 3) Then
        If (PicX = 640) Then
            Count3 = Mid(TreeView1.DropHighlight.Key, 2, 2) - 10
            Read = "Track " + Trim(Str(Count3))
            If (Mid(Dir1.Path, 3, 1) = "\") And (Len(Dir1.Path) = 3) Then
                Read = oMisc.WriteINI(Read, "BPic", Dir1.Path + File1.FileName, ProgramDir + "\WorkCopy.lda")
            Else
                Read = oMisc.WriteINI(Read, "BPic", Dir1.Path + "\" + File1.FileName, ProgramDir + "\WorkCopy.lda")
            End If
            Read = oMisc.ReadINI("Track " + Trim(Str(Count3)), "BPic", ProgramDir + "\WorkCopy.lda")
            Count2 = 190
            Read2 = "BMP"
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 200)), "Full Pic: " + Read, 3, 3)
            TreeView1.DropHighlight = Nothing
            Exit Sub
        Else
            Count3 = Mid(TreeView1.DropHighlight.Key, 2, 2) - 10
            Read = "Track " + Trim(Str(Count3))
            If (Mid(Dir1.Path, 3, 1) = "\") And (Len(Dir1.Path) = 3) Then
                Read = oMisc.WriteINI(Read, "SPic", Dir1.Path + File1.FileName, ProgramDir + "\WorkCopy.lda")
            Else
                Read = oMisc.WriteINI(Read, "SPic", Dir1.Path + "\" + File1.FileName, ProgramDir + "\WorkCopy.lda")
            End If
            Read = oMisc.ReadINI("Track " + Trim(Str(Count3)), "SPic", ProgramDir + "\WorkCopy.lda")
            Count2 = 290
            Read2 = "BMP"
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 300)), "Framed Pic: " + Read, 3, 3)
            TreeView1.DropHighlight = Nothing
            Exit Sub
        End If
    End If
    MsgBox "You can't place this object on this place!", vbInformation, TH
    TreeView1.DropHighlight = Nothing
    Exit Sub
errorstop:
    Select Case Err.Number
    Case "35602"
        CountNr = Mid(TreeView1.DropHighlight.Child.Key, 2, 3) - Count2
        If CountNr = Mid(TreeView1.DropHighlight.Key, 2, 2) Then
            X = TreeView1.DropHighlight.Child.Index
        Else
            CountNr = Mid(TreeView1.DropHighlight.Child.Next.Key, 2, 3) - Count2
            If CountNr = Mid(TreeView1.DropHighlight.Key, 2, 2) Then
                X = TreeView1.DropHighlight.Child.Next.Index
            Else
                CountNr = Mid(TreeView1.DropHighlight.Child.Next.Next.Key, 2, 3) - Count2
                If CountNr = Mid(TreeView1.DropHighlight.Key, 2, 2) Then
                    X = TreeView1.DropHighlight.Child.Next.Next.Index
                End If
            End If
        End If
        TreeView1.Nodes.Remove (X)
        If Read2 = "DAT" Then
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 100)), "Track File: " + Read, 4, 4)
        End If
        If Count2 = 290 Then
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 300)), "Framed Pic: " + Read, 3, 3)
        End If
        If Count2 = 190 Then
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 200)), "Full Pic: " + Read, 3, 3)
        End If
        TreeView1.DropHighlight = Nothing
    Case "91"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
        Set TreeView1.DropHighlight = Nothing
        Exit Sub
    End Select
End Sub


Private Sub TreeView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If InDrag = True Then
        ' Set DropHighlight to the mouse's coordinates.
        Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
    End If
End Sub

Private Sub TreeView1_GotFocus()
    MDIForm1.Toolbar1.Buttons.Item(5).Enabled = True
End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        DeleteTrack
    End If
End Sub

Private Sub TreeView1_LostFocus()
    MDIForm1.Toolbar1.Buttons.Item(5).Enabled = False
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
    If (TreeView1.SelectedItem.Text = "GP2 Track's") Or (TreeView1.SelectedItem.Key = "q") Then
        SaveLastClick
        frameFile.Visible = True
        frameTrackInfo.Visible = True
        frameInfo.Visible = False
        frameData.Visible = False
        frmaeTime.Visible = False
        frameFile.Top = frameInfo.Top
        frameFile.Left = frameInfo.Left
        imgFramed.Picture = Nothing
        imgFull.Picture = Nothing
        frameGlobal.Visible = False
        framePlayer.Visible = False
        lblNote.Visible = False
        LastClick = "GP2 Track's"
        CurrentRecord = 0
        Exit Sub
    End If
    imgPre.Picture = Nothing
    If Mid(TreeView1.SelectedItem.Key, 1, 1) = "t" Then
        SaveLastClick
        frameFile.Visible = False
        frameTrackInfo.Visible = False
        frameInfo.Visible = True
        frameData.Visible = True
        frmaeTime.Visible = True
        frameGlobal.Visible = False
        framePlayer.Visible = False
        lblNote.Visible = False

        Count4 = Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
        If (Count4 > 10) And (Count4 < 100) Then Count4 = Count4 - 10
        If (Count4 > 100) And (Count4 < 200) Then Count4 = Count4 - 100
        If (Count4 > 200) And (Count4 < 300) Then Count4 = Count4 - 200
        If Count4 > 300 Then Count4 = Count4 - 300
        
        If Count4 = CurrentRecord Then Exit Sub
        'If CurrentRecord <> 0 Then SaveFileData
        GetFileData Count4
        CurrentRecord = Count4
        LastClick = "Track"
        Exit Sub
    End If
    If TreeView1.SelectedItem.Key = "e" Then
        SaveLastClick
        frameFile.Visible = False
        frameTrackInfo.Visible = False
        frameInfo.Visible = False
        frameData.Visible = False
        frmaeTime.Visible = False
        frameGlobal.Visible = True
        framePlayer.Visible = True
        GetPlayerData
        LastClick = "Player"
        Read = oMisc.File_Exists(Gp2Dir + "\f1gstate.sav")
        If Read = False Then
            HScroll2.Enabled = False
            lblNote.Visible = True
        Else
            HScroll2.Enabled = True
            lblNote.Visible = False
        End If
    End If
    Exit Sub
ErrorTrap:
    Select Case Err.Number
        Case Else
            MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
            Set TreeView1.DropHighlight = Nothing
            Exit Sub
    End Select
    Exit Sub
End Sub

Private Sub A_Click(Index As Integer)
    If A(Index).Tag = "On" Then
        A(Index).Picture = On1(Index).Picture
        A(Index).Tag = "Off"
    Else
        A(Index).Picture = Off(Index).Picture
        A(Index).Tag = "On"
    End If
    
End Sub
Private Sub AC_Click(Index As Integer)
    If AC(Index).Tag = "On" Then
        AC(Index).Picture = On1(Index).Picture
        AC(Index).Tag = "Off"
    Else
        AC(Index).Picture = Off(Index).Picture
        AC(Index).Tag = "On"
    End If
    
End Sub

Private Sub cmdDefaultSettings_Click()
    hscPitSpeed.Value = 50
    hscPQPower.Value = 790
    hscPRPower.Value = 780
    hscWeight.Value = 1313
    HScroll1.Value = 1313
    HScroll2.Value = 5
    hscPGrip.Value = 198
    GP2AidsSet
    chkNoLimit.Value = 0
    chkSave.Value = 0
    chk0as1.Value = 1
    Slider1.Value = 1994
    chkSelectedTeam.Value = 0
    
End Sub

Private Sub cmdExportSettings_Click()
    SavePlayerData
    If HScroll2.Enabled = True Then
        F1SaveFileNum = FreeFile
        Open Gp2Dir + "\F1gstate.sav" For Binary As F1SaveFileNum
            ExportQuickRace
        Close F1SaveFileNum
    End If
    GP2FileNum = FreeFile
    Open Gp2Dir + "\GP2.exe" For Binary As GP2FileNum
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
    Close GP2FileNum
    If HScroll2.Enabled = True Then
        DeleteFile Gp2Dir + "\$$Check$.bat"
        SourceFile = ProgramDir + "\gp2utils\check.exe"
        TargetFile = Gp2Dir + "\$$check.exe"
        FileCopy SourceFile, TargetFile
        FileNum = FreeFile
        Open Gp2Dir + "\$$Check$.bat" For Append As FileNum
        Print #FileNum, Mid(Gp2Dir, 1, 2)
        Print #FileNum, "cd " + Gp2Dir
        Print #FileNum, Gp2Dir + "\$$Check f1gstate.sav"
        Print #FileNum, "del $$Check.exe"
        Close FileNum
        Read = oMisc.File_Exists("c:\command.com")
        Dim RetVal
        If Read = True Then
            ChDir Gp2Dir
            RetVal = Shell("c:\command.com /c " + Gp2Dir + "\$$Check$.bat", vbNormalFocus)
        Else
            ChDir Gp2Dir
            RetVal = Shell(Gp2Dir + "\$$Check$.bat", vbNormalFocus)
        End If
    End If
End Sub

Private Sub cmdImportSettings_Click()
    SavePlayerData
    GP2FileNum = FreeFile
    Open Gp2Dir + "\GP2.exe" For Binary As GP2FileNum
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
    Close GP2FileNum
    If HScroll2.Enabled = True Then
        F1SaveFileNum = FreeFile
        Open Gp2Dir + "\F1gstate.sav" For Binary As F1SaveFileNum
            ImportQuick
        Close F1SaveFileNum
    End If
    
    GetPlayerData
End Sub

Private Sub Command1_Click()
    MsgBox "The time must be writen in this formate #,##,###  and the date must be writen like this, 1999-01-24. You my not enter a date 'lower' then 1978-01-01 and not a time higher the 9,59,999.", vbInformation, TH
End Sub

Private Sub lblYear2_Click()
    On Error GoTo ErrorTrap
    Read = InputBox("Year (1900-2099):", "Select Year")
    If Read = "" Then Exit Sub
    If (Read > 1899) And (Read < 3000) Then Slider1.Value = Read
ErrorTrap:
    Exit Sub
End Sub

Private Sub P_Click(Index As Integer)
    If P(Index).Tag = "On" Then
        P(Index).Picture = On1(Index).Picture
        P(Index).Tag = "Off"
    Else
        P(Index).Picture = Off(Index).Picture
        P(Index).Tag = "On"
    End If
    
End Sub

Private Sub R_Click(Index As Integer)
    If R(Index).Tag = "On" Then
        R(Index).Picture = On1(Index).Picture
        R(Index).Tag = "Off"
    Else
        R(Index).Picture = Off(Index).Picture
        R(Index).Tag = "On"
    End If
    
End Sub

Private Sub S_Click(Index As Integer)
    If S(Index).Tag = "On" Then
        S(Index).Picture = On1(Index).Picture
        S(Index).Tag = "Off"
    Else
        S(Index).Picture = Off(Index).Picture
        S(Index).Tag = "On"
    End If
    
End Sub

Private Sub chkNoLimit_Click()
    If chkNoLimit.Value = 1 Then
        hscPitSpeed.Enabled = False
    Else
        hscPitSpeed.Enabled = True
    End If
    
End Sub

Private Sub chkSelectedTeam_Click()
    If chkSelectedTeam.Value = 1 Then
        hscPQPower.Enabled = False
        hscPRPower.Enabled = False
    Else
        hscPQPower.Enabled = True
        hscPRPower.Enabled = True
    End If
    
End Sub

Private Sub cmdAdd_Click()
    DataBaseFileNum = FreeFile
    RecordLen = Len(TimeBase)
    Open ProgramDir + "\database.tdb" For Random As DataBaseFileNum Len = RecordLen
    LastRecord2 = FileLen(ProgramDir + "\database.tdb") / RecordLen
    LastRecord2 = LastRecord2 + 1
    TimeBase.TName = txtName
    TimeBase.QTime = txtQTime
    TimeBase.RTime = txtRTime
    TimeBase.QDriver = txtQDriver
    TimeBase.RDriver = txtRDriver
    TimeBase.QTeam = txtQTeam
    TimeBase.RTeam = txtRTeam
    TimeBase.QDate = txtQDate
    TimeBase.RDate = txtRDate
    TimeBase.TName = txtName
    Put #DataBaseFileNum, LastRecord2, TimeBase
    Close DataBaseFileNum
End Sub

Private Sub cmdBrowse_Click()
    X17 = 32000
    On Error GoTo ErrorTrap
    If TrackPath <> "" Then
        CommonDialog1.InitDir = TrackPath
    Else
        CommonDialog1.InitDir = DefaultTrackPath
    End If
    CommonDialog1.Filter = "GP2 Track Files (*.dat)|*.dat|All Files (*.*)|*.*|"
    CommonDialog1.ShowOpen
    Read3 = CommonDialog1.FileName
    TrackPath = Read3
    ReadTrackFile (CommonDialog1.FileName)
    If NoSupport = True Then Exit Sub

    Read = lblCountry
    GetAdjectiv
    txtAdjectiv = Read
    txtCountry = lblCountry
    txtLaps = lblLaps
    txtLength = lblLength
    txtName = lblName
    txtPath = CommonDialog1.FileName
    txtTire = lblWare
    txtQTime = lblQLap
    txtRTime = lblRLap

    X17 = 1

    X = Mid(TreeView1.SelectedItem.Key, 2, 2) - 10
    TreeView1.SelectedItem.Text = Trim(Str(X)) + ". " + txtName

    
    Dim nodX As Node    ' Create variable.
    Read = TreeView1.SelectedItem.Key
    If TreeView1.SelectedItem.Children > 0 Then
        If TreeView1.SelectedItem.Child.Key = "t" + Trim(Str(X + 100)) Then
            TreeView1.Nodes.Remove (TreeView1.SelectedItem.Child.Index)
        End If
        If TreeView1.SelectedItem.Children = 2 Then
            If TreeView1.SelectedItem.Child.Next.Key = "t" + Trim(Str(X + 100)) Then
                TreeView1.Nodes.Remove (TreeView1.SelectedItem.Child.Next.Index)
            End If
        End If
        If TreeView1.SelectedItem.Children = 3 Then
            If TreeView1.SelectedItem.Child.Next.Next.Key = "t" + Trim(Str(X + 100)) Then
                TreeView1.Nodes.Remove (TreeView1.SelectedItem.Child.Next.Next.Index)
            End If
        End If
    End If
    Set nodX = TreeView1.Nodes.Add(Read, tvwChild, "t" + Trim(Str(X + 100)), "Track File: " + CommonDialog1.FileName, 4, 4)
    
    
    Exit Sub
ErrorTrap:
    Select Case Err.Number
    Case "32755"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Private Sub cmdGet_Click()
    frmBestLap.Show , MDIForm1
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrorTrap
    Dir1.Path = Drive1.Drive
Exit Sub
ErrorTrap:
    Select Case Err.Number
    Case 68
        MsgBox Err.Description, vbCritical
        Drive1.Drive = "c:"
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Private Sub File1_Click()
    lblAuthor.Caption = "Author:"
    lblEvent.Caption = "Event:"
    lblYear.Caption = "Year:"
    lblMisc.Caption = "Misc Info:"
    lblRLap.Caption = ""
    lblQLap.Caption = ""
    lblWare.Caption = ""
    lblLength.Caption = ""
    lblLaps.Caption = ""
    lblCountry.Caption = ""
    lblName.Caption = ""
    CountNr = Len(File1.FileName)
    Read = UCase(Mid(File1.FileName, CountNr - 2, 3))
    If (UCase(Read) = UCase("bmp")) Or (UCase(Read) = UCase("gif")) Then
        Read = Dir1.Path + "\" + File1.FileName
        Set imgSize.Picture = LoadPicture(Read)
        PicY = imgSize.Height / 15
        PicX = imgSize.Width / 15
        If ((PicX = 640) And (PicY = 480)) Then
            imgPre.Height = (PicY * 15) / 2
            imgPre.Width = (PicX * 15) / 2
            Set imgPre.Picture = LoadPicture(Read)
            NoSupport = False
            Exit Sub
        End If
        If ((PicX = 440) And (PicY = 330)) Then
            imgPre.Height = (PicY * 15) / 2
            imgPre.Width = (PicX * 15) / 2
            Set imgPre.Picture = LoadPicture(Read)
            NoSupport = False
            Exit Sub
        End If
        MsgBox "This picture is not supported by " + TH + ".", vbInformation, TH
        Exit Sub
    End If
    Set imgPre = Nothing
    If (Mid(Dir1.Path, 3, 1) = "\") And (Len(Dir1.Path) = 3) Then
        ReadTrackFile (Dir1.Path + File1.FileName)
    Else
        ReadTrackFile (Dir1.Path + "\" + File1.FileName)
    End If
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If NoSupport = False Then
        Dim DY  As Long ' Declare variable.
        DY = 200 'TextHeight("A")    ' Get height of one line.
        Label1.Move File1.Left + 30, File1.Top + Y + DY / 3, File1.Width - 30, DY
        Label1.Drag ' Drag label outline.
        InDrag = True
    End If
End Sub

Private Sub File1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Source.Tag = Dir1.Path + "\" + File1.FileName
End Sub

Private Sub hscPGrip_Change()
    lblGrip.Caption = hscPGrip.Value
    
End Sub

Private Sub hscPGrip_Scroll()
    lblGrip.Caption = hscPGrip.Value
End Sub

Private Sub hscPitSpeed_Change()
    X = hscPitSpeed.Value
    X = X * 1.5966
    lblPitSpeed.Caption = Str(hscPitSpeed.Value) + "mph (" + Trim(Str(X)) + "km/h)"
    
End Sub

Private Sub hscPitSpeed_Scroll()
    X = hscPitSpeed.Value
    X = X * 1.5966
    lblPitSpeed.Caption = Str(hscPitSpeed.Value) + "mph (" + Trim(Str(X)) + "km/h)"
End Sub

Private Sub hscPQPower_Change()
    lblPQPower.Caption = hscPQPower.Value
    
End Sub

Private Sub hscPQPower_Scroll()
    lblPQPower.Caption = hscPQPower.Value
End Sub

Private Sub hscPRPower_Change()
    lblPRPower.Caption = hscPRPower.Value
    
End Sub

Private Sub hscPRPower_Scroll()
    lblPRPower.Caption = hscPRPower.Value
End Sub

Private Sub HScroll1_Change()
    X = HScroll1.Value
    X = X / 2.203020134
    lblCWeight.Caption = Str(HScroll1.Value) + "lb (" + Trim(Str(X)) + "kg)"
    
End Sub

Private Sub HScroll1_Scroll()
    X = HScroll1.Value
    X = X / 2.203020134
    lblCWeight.Caption = Str(HScroll1.Value) + "lb (" + Trim(Str(X)) + "kg)"
End Sub

Private Sub HScroll2_Change()
    lblQuick.Caption = Str(HScroll2.Value) + "%"
    
End Sub

Private Sub HScroll2_Scroll()
    lblQuick.Caption = Str(HScroll2.Value) + "%"
End Sub

Private Sub hscWeight_Change()
    X = hscWeight.Value
    X = X / 2.203020134
    lblWeight2.Caption = Str(hscWeight.Value) + "lb (" + Trim(Str(X)) + "kg)"
    
End Sub

Private Sub hscWeight_Scroll()
    X = hscWeight.Value
    X = X / 2.203020134
    lblWeight2.Caption = Str(hscWeight.Value) + "lb (" + Trim(Str(X)) + "kg)"
End Sub

Private Sub Slider1_Change()
    lblYear2.Caption = Slider1.Value
    
End Sub

Private Sub Slider1_Scroll()
    lblYear2.Caption = Slider1.Value
End Sub

Private Sub txtAdjectiv_GotFocus()
    txtAdjectiv.SelStart = 0
    txtAdjectiv.SelLength = Len(txtAdjectiv)
End Sub

Private Sub txtCountry_GotFocus()
    txtCountry.SelStart = 0
    txtCountry.SelLength = Len(txtCountry)
End Sub

Private Sub txtLaps_Change()
    If (txtLaps <> "") And (txtLaps <> "0") Then
        VScroll1.Value = txtLaps.Text
    End If
End Sub

Private Sub txtLaps_GotFocus()
    txtLaps.SelStart = 0
    txtLaps.SelLength = Len(txtLaps)
End Sub

Private Sub txtLength_GotFocus()
    txtLength.SelStart = 0
    txtLength.SelLength = Len(txtLength)
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_LostFocus()
    
    If (Len(txtName) > 0) Then
        X = Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
        If (X > 10) And (X < 100) Then X = X - 10
        If (X > 100) And (X < 200) Then X = X - 100
        If (X > 200) And (X < 300) Then X = X - 200
        If X > 300 Then X = X - 300
        
        TreeView1.SelectedItem.Text = Trim(Str(X)) + ". " + txtName.Text
    Else
        X = Mid(TreeView1.SelectedItem.Key, 2, 2) - 10
        TreeView1.SelectedItem.Text = "Track " + Trim(Str(X))
    End If
End Sub

Private Sub txtQDate_GotFocus()
    txtQDate.SelStart = 0
    txtQDate.SelLength = Len(txtQDate)
End Sub

Private Sub txtQDriver_GotFocus()
    txtQDriver.SelStart = 0
    txtQDriver.SelLength = Len(txtQDriver)
End Sub

Private Sub txtQTeam_GotFocus()
    txtQTeam.SelStart = 0
    txtQTeam.SelLength = Len(txtQTeam)
End Sub

Private Sub txtQTime_GotFocus()
    txtQTime.SelStart = 0
    txtQTime.SelLength = Len(txtQTime)
End Sub

Private Sub txtRDate_GotFocus()
    txtRDate.SelStart = 0
    txtRDate.SelLength = Len(txtRDate)
End Sub

Private Sub txtRDriver_GotFocus()
    txtRDriver.SelStart = 0
    txtRDriver.SelLength = Len(txtRDriver)
End Sub

Private Sub txtRTeam_GotFocus()
    txtRTeam.SelStart = 0
    txtRTeam.SelLength = Len(txtRTeam)
End Sub

Private Sub txtRTime_GotFocus()
    txtRTime.SelStart = 0
    txtRTime.SelLength = Len(txtRTime)
End Sub

Private Sub txtTire_GotFocus()
    txtTire.SelStart = 0
    txtTire.SelLength = Len(txtTire)
End Sub

Private Sub VScroll1_Change()
    txtLaps = VScroll1.Value
End Sub
