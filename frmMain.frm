VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GP2 Track Handler v1.4"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgToolBar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   1
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
            Key             =   "Import"
            Object.ToolTipText     =   "Import data from GP2"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Export"
            Object.ToolTipText     =   "Export data to GP2"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GP2Exe"
            Object.ToolTipText     =   "Add a GP2Edit file"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GP2"
            Object.ToolTipText     =   "Run GP2"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   0
      TabIndex        =   111
      Top             =   420
      Width           =   9255
   End
   Begin ComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   108
      Top             =   6750
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5203
            MinWidth        =   5203
            Text            =   "GP2 Track Handler © Viktor Gars 98-99"
            TextSave        =   "GP2 Track Handler © Viktor Gars 98-99"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5115
            MinWidth        =   5115
            Text            =   "GP2 Version:"
            TextSave        =   "GP2 Version:"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5998
            MinWidth        =   5998
            Text            =   "GP2 Directory:"
            TextSave        =   "GP2 Directory:"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   6255
      Left            =   3240
      TabIndex        =   90
      Top             =   480
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "File Manager"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frameNorm"
      Tab(0).Control(1)=   "fraFileInfo"
      Tab(0).Control(2)=   "fraMenuPic"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Track Data"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frameData"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frameInfo"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraQual"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraRace"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdJamCheck"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Menu Pics"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabPic"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Performence Edit"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "framePlayer"
      Tab(3).Control(1)=   "frameGlobal"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Lap Times"
      TabPicture(4)   =   "frmMain.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblTimeDrag"
      Tab(4).Control(1)=   "lblTextLen"
      Tab(4).Control(2)=   "cmdDelete"
      Tab(4).Control(3)=   "cmdAdd"
      Tab(4).Control(4)=   "lstTime"
      Tab(4).ControlCount=   5
      Begin ComctlLib.ListView lstTime 
         DragIcon        =   "frmMain.frx":0396
         Height          =   5175
         Left            =   -74880
         TabIndex        =   123
         Top             =   480
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDragMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Pos."
            Object.Width           =   441
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Track"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Driver"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Time"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   847
         EndProperty
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         Height          =   315
         Left            =   -73680
         TabIndex        =   117
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   -74880
         TabIndex        =   116
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdJamCheck 
         Caption         =   "&Jam Check"
         Height          =   375
         Left            =   120
         TabIndex        =   88
         Top             =   5640
         Width           =   1095
      End
      Begin TabDlg.SSTab tabPic 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   56
         Top             =   480
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   9975
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Big Menu Picture"
         TabPicture(0)   =   "frmMain.frx":06A0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "imgBPic"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtBPic"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Smal Menu Picture"
         TabPicture(1)   =   "frmMain.frx":06BC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtSPic"
         Tab(1).Control(1)=   "imgSPic"
         Tab(1).ControlCount=   2
         Begin VB.TextBox txtSPic 
            Height          =   285
            Left            =   -74760
            TabIndex        =   110
            Top             =   3120
            Width           =   2055
            Visible         =   0   'False
         End
         Begin VB.TextBox txtBPic 
            Height          =   285
            Left            =   240
            TabIndex        =   109
            Top             =   4200
            Width           =   1815
            Visible         =   0   'False
         End
         Begin VB.Image imgSPic 
            BorderStyle     =   1  'Fixed Single
            Height          =   2475
            Left            =   -74760
            Stretch         =   -1  'True
            Top             =   480
            Width           =   3300
         End
         Begin VB.Image imgBPic 
            BorderStyle     =   1  'Fixed Single
            Height          =   3600
            Left            =   240
            Stretch         =   -1  'True
            Top             =   480
            Width           =   4800
         End
      End
      Begin VB.Frame frameGlobal 
         Caption         =   "Global Settings"
         Height          =   5175
         Left            =   -72000
         TabIndex        =   101
         Top             =   480
         Width           =   2775
         Begin VB.CheckBox chkSave 
            Caption         =   "Always save track record"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   3120
            Width           =   2175
         End
         Begin VB.CheckBox chk0as1 
            Caption         =   "Show car 1 as 0"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   2760
            Width           =   2055
         End
         Begin VB.HScrollBar hscQRace 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   79
            Top             =   1200
            Value           =   5
            Width           =   2415
         End
         Begin VB.HScrollBar hscCWeight 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   4000
            Min             =   401
            TabIndex        =   76
            Top             =   600
            Value           =   1313
            Width           =   2415
         End
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   120
            TabIndex        =   102
            Top             =   3600
            Width           =   2415
         End
         Begin ComctlLib.Slider Slider1 
            Height          =   375
            Left            =   240
            TabIndex        =   82
            Top             =   2280
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            _Version        =   327682
            LargeChange     =   10
            Min             =   1900
            Max             =   2099
            SelStart        =   1994
            TickFrequency   =   10
            Value           =   1994
         End
         Begin VB.Label lblYear 
            AutoSize        =   -1  'True
            Caption         =   "1994"
            Height          =   195
            Left            =   1800
            TabIndex        =   81
            Top             =   1920
            Width           =   360
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Year for this Season:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   80
            Top             =   1920
            Width           =   1470
         End
         Begin VB.Label lblQuick 
            Alignment       =   1  'Right Justify
            Caption         =   "5%"
            Height          =   195
            Left            =   2175
            TabIndex        =   78
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Quick Race Length"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   77
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lblCWeight 
            Alignment       =   1  'Right Justify
            Caption         =   "1313lb (596kg)"
            Height          =   195
            Left            =   1455
            TabIndex        =   75
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Car Weight"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   74
            Top             =   360
            Width           =   795
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
            TabIndex        =   107
            Top             =   4800
            Width           =   285
         End
         Begin VB.Label lblRookie 
            AutoSize        =   -1  'True
            Caption         =   "Rookie"
            Height          =   180
            Index           =   3
            Left            =   1800
            TabIndex        =   106
            Top             =   3840
            Width           =   510
         End
         Begin VB.Label lblAmateur 
            AutoSize        =   -1  'True
            Caption         =   "Amateur"
            Height          =   180
            Left            =   1800
            TabIndex        =   105
            Top             =   4080
            Width           =   585
         End
         Begin VB.Label lblSemiPro 
            AutoSize        =   -1  'True
            Caption         =   "Semi-Pro"
            Height          =   180
            Left            =   1800
            TabIndex        =   104
            Top             =   4320
            Width           =   630
         End
         Begin VB.Label lblPro 
            AutoSize        =   -1  'True
            Caption         =   "Pro"
            Height          =   180
            Left            =   1800
            TabIndex        =   103
            Top             =   4560
            Width           =   240
         End
      End
      Begin VB.Frame framePlayer 
         Caption         =   "Player Car"
         Height          =   5175
         Left            =   -74880
         TabIndex        =   99
         Top             =   480
         Width           =   2775
         Begin VB.CommandButton cmdDefaultSettings 
            Caption         =   "GP2 Default"
            Height          =   315
            Left            =   1500
            TabIndex        =   87
            Top             =   4320
            Width           =   1035
         End
         Begin VB.CommandButton cmdImportSettings 
            Caption         =   "&Import"
            Height          =   315
            Left            =   120
            TabIndex        =   86
            Top             =   4700
            Width           =   1035
         End
         Begin VB.CommandButton cmdExportSettings 
            Caption         =   "&Export"
            Height          =   315
            Left            =   120
            TabIndex        =   85
            Top             =   4320
            Width           =   1035
         End
         Begin VB.CheckBox chkUPower 
            Caption         =   "Use Selected Team Power"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1560
            Width           =   2500
         End
         Begin VB.CheckBox chkNoLimit 
            Caption         =   "No Speed Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   3840
            Width           =   2500
         End
         Begin VB.HScrollBar hscPitSpeed 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   201
            Min             =   1
            TabIndex        =   72
            Top             =   3480
            Value           =   50
            Width           =   2415
         End
         Begin VB.HScrollBar hscPGrip 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   1000
            TabIndex        =   69
            Top             =   2880
            Value           =   198
            Width           =   2415
         End
         Begin VB.HScrollBar hscWeight 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   4000
            Min             =   401
            TabIndex        =   66
            Top             =   2280
            Value           =   1313
            Width           =   2415
         End
         Begin VB.HScrollBar hscPQPower 
            Height          =   255
            LargeChange     =   20
            Left            =   120
            Max             =   1579
            TabIndex        =   62
            Top             =   1200
            Value           =   790
            Width           =   2415
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   120
            TabIndex        =   100
            Top             =   4200
            Width           =   2415
         End
         Begin VB.HScrollBar hscPRPower 
            Height          =   255
            LargeChange     =   20
            Left            =   120
            Max             =   1579
            TabIndex        =   59
            Top             =   600
            Value           =   780
            Width           =   2415
         End
         Begin VB.Label lblPit 
            AutoSize        =   -1  'True
            Caption         =   "Pit Speed Limit"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   3240
            Width           =   1050
         End
         Begin VB.Label lblPitSpeed 
            Alignment       =   1  'Right Justify
            Caption         =   "50mph (80km/h)"
            Height          =   195
            Left            =   1095
            TabIndex        =   71
            Top             =   3240
            Width           =   1440
         End
         Begin VB.Label lblGrip 
            Alignment       =   1  'Right Justify
            Caption         =   "198"
            Height          =   195
            Left            =   2055
            TabIndex        =   68
            Top             =   2640
            Width           =   480
         End
         Begin VB.Label lblGrip2 
            AutoSize        =   -1  'True
            Caption         =   "Grip"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   2640
            Width           =   285
         End
         Begin VB.Label lblWeight2 
            Alignment       =   1  'Right Justify
            Caption         =   "1313lb (596Kg)"
            Height          =   195
            Left            =   1335
            TabIndex        =   65
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label lblWeight 
            AutoSize        =   -1  'True
            Caption         =   "Car Weight"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label lblPRPower 
            Alignment       =   1  'Right Justify
            Caption         =   "780"
            Height          =   195
            Left            =   2055
            TabIndex        =   58
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Power in Qual"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   990
         End
         Begin VB.Label lblPQPower 
            Alignment       =   1  'Right Justify
            Caption         =   "790"
            Height          =   195
            Left            =   2055
            TabIndex        =   61
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Power in Race"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   1050
         End
      End
      Begin VB.Frame fraRace 
         Caption         =   "Lap Time Data - Race"
         Height          =   2775
         Left            =   3480
         TabIndex        =   97
         Top             =   3360
         Width           =   2175
         Begin VB.CommandButton cmdSaveRace 
            Height          =   300
            Left            =   1800
            Picture         =   "frmMain.frx":06D8
            Style           =   1  'Graphical
            TabIndex        =   119
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtRDate 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   55
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtRTeam 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   53
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox txtRDriver 
            Height          =   285
            Left            =   120
            MaxLength       =   23
            TabIndex        =   51
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtRTime 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   49
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Driver"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   840
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Team"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Date (e.g. 1999-06-19)"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   2040
            Width           =   1605
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Time (e.g. 1:24.145)"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.Frame fraQual 
         Caption         =   "Lap Time Data - Qual"
         Height          =   2775
         Left            =   3480
         TabIndex        =   96
         Top             =   480
         Width           =   2175
         Begin VB.CommandButton cmdSaveQual 
            Height          =   300
            Left            =   1800
            Picture         =   "frmMain.frx":07DA
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtQDate 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   47
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtQTeam 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   45
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox txtQDriver 
            Height          =   285
            Left            =   120
            MaxLength       =   23
            TabIndex        =   43
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtQTime 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   41
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Date (e.g 1999-06-19)"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   2040
            Width           =   1560
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Team"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   1440
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Driver"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Time (e.g. 1:24.145)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.Frame frameInfo 
         Caption         =   "Track Info"
         Height          =   2775
         Left            =   120
         TabIndex        =   94
         Top             =   480
         Width           =   3135
         Begin VB.TextBox txtAdjectiv 
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox txtCountry 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   255
            TabIndex        =   115
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Track Path"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Track Name:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Country:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Adjective:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   2040
            Width           =   705
         End
      End
      Begin VB.Frame frameData 
         Caption         =   "Track Data"
         Height          =   2175
         Left            =   120
         TabIndex        =   93
         Top             =   3360
         Width           =   3135
         Begin VB.VScrollBar vscLaps 
            Height          =   285
            Left            =   600
            Max             =   0
            Min             =   126
            TabIndex        =   35
            Top             =   480
            Value           =   3
            Width           =   200
         End
         Begin VB.TextBox txtTire 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   39
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtLength 
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   37
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtLaps 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   34
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tyre &Ware (14000-40000)"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   1830
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tra&ck Length (0-9999 m)"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Laps (3-126)"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame frameNorm 
         Caption         =   " Select Track/Menu Picture "
         Height          =   3550
         Left            =   -74880
         TabIndex        =   92
         Top             =   480
         Width           =   5655
         Begin ComctlLib.ListView lstFile 
            Height          =   3210
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   5662
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            OLEDragMode     =   1
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            OLEDragMode     =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "File Name"
               Object.Width           =   3616
            EndProperty
         End
         Begin VB.CommandButton cmdJam 
            Caption         =   "&Jam Check"
            Height          =   435
            Left            =   2880
            TabIndex        =   26
            ToolTipText     =   "Check if all jamfiles are installed"
            Top             =   3015
            Width           =   2655
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   2880
            TabIndex        =   0
            Top             =   2640
            Width           =   2655
         End
         Begin VB.DirListBox Dir1 
            Height          =   2340
            Left            =   2880
            TabIndex        =   1
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label1 
            Height          =   135
            Left            =   2280
            TabIndex        =   98
            Top             =   4080
            Width           =   1335
         End
      End
      Begin VB.Frame fraFileInfo 
         Caption         =   "Track File Info"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   112
         Top             =   4080
         Width           =   5655
         Begin VB.TextBox lblRace 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1250
            MaxLength       =   8
            TabIndex        =   14
            ToolTipText     =   "Click to edit"
            Top             =   1665
            Width           =   1700
         End
         Begin VB.TextBox lblQual 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1250
            MaxLength       =   8
            TabIndex        =   12
            ToolTipText     =   "Click to edit"
            Top             =   1425
            Width           =   1700
         End
         Begin VB.TextBox lblWare 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1250
            MaxLength       =   5
            TabIndex        =   10
            ToolTipText     =   "Click to edit"
            Top             =   1185
            Width           =   1700
         End
         Begin VB.TextBox lblLen 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1250
            MaxLength       =   4
            TabIndex        =   8
            ToolTipText     =   "Click to edit"
            Top             =   945
            Width           =   1700
         End
         Begin VB.TextBox lblLaps 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1250
            MaxLength       =   3
            TabIndex        =   6
            ToolTipText     =   "Click to edit"
            Top             =   720
            Width           =   1700
         End
         Begin VB.TextBox lblTrackName 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1250
            TabIndex        =   4
            ToolTipText     =   "Click to edit"
            Top             =   465
            Width           =   1700
         End
         Begin VB.TextBox lblEvent 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3600
            TabIndex        =   18
            ToolTipText     =   "Click to edit"
            Top             =   495
            Width           =   1935
         End
         Begin VB.TextBox lblInfoYear 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3600
            MaxLength       =   4
            TabIndex        =   16
            ToolTipText     =   "Click to edit"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox lblSlot 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   20
            ToolTipText     =   "Click to edit"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox lblMisc 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   3000
            MultiLine       =   -1  'True
            TabIndex        =   24
            ToolTipText     =   "Click to edit"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.CommandButton cmdSaveGP2Info 
            Enabled         =   0   'False
            Height          =   360
            Left            =   5160
            Picture         =   "frmMain.frx":08DC
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Save GP2Info"
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox lblCountry 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1250
            TabIndex        =   3
            ToolTipText     =   "Click to edit"
            Top             =   225
            Width           =   1700
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#113"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   122
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblAuthor 
            Height          =   225
            Left            =   3600
            TabIndex        =   22
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#118"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#120"
            Height          =   255
            Index           =   8
            Left            =   3000
            TabIndex        =   17
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#119"
            Height          =   255
            Index           =   7
            Left            =   3000
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#117"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#116"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#115"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#114"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#112"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#121"
            Height          =   255
            Index           =   9
            Left            =   3000
            TabIndex        =   19
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#122"
            Height          =   255
            Index           =   10
            Left            =   3000
            TabIndex        =   21
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblInfoText 
            Caption         =   "#123"
            Height          =   255
            Index           =   11
            Left            =   3000
            TabIndex        =   23
            Top             =   1200
            Width           =   615
         End
      End
      Begin VB.Frame fraMenuPic 
         Caption         =   "Menu Picture"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   113
         Top             =   4080
         Width           =   5655
         Begin VB.Image imgPre 
            Height          =   1560
            Left            =   120
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2040
         End
         Begin VB.Label lblPicInfo 
            Height          =   255
            Left            =   2280
            TabIndex        =   114
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Label lblTextLen 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   -70800
         TabIndex        =   121
         Top             =   6000
         Width           =   45
         Visible         =   0   'False
      End
      Begin VB.Label lblTimeDrag 
         Height          =   375
         Left            =   -71040
         TabIndex        =   118
         Top             =   5520
         Width           =   1695
      End
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   6247
      Left            =   0
      TabIndex        =   89
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   11007
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "TreeViewImages"
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin ComctlLib.ImageList TreeViewImages 
      Left            =   0
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":09DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0CA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0F62
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgRealSize 
      Height          =   735
      Left            =   2760
      Top             =   8040
      Width           =   975
      Visible         =   0   'False
   End
   Begin ComctlLib.ImageList imgToolBar 
      Left            =   1920
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":131E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1430
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1542
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1654
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":182E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":203C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2356
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
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "File &Converter..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import from GP2..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export to GP2..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSep5 
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
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuGP2Path 
         Caption         =   "Select GP2 Path..."
      End
      Begin VB.Menu mnuTrackPath 
         Caption         =   "Select Default Track Path..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatusbar 
         Caption         =   "&Statusbar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuRand 
         Caption         =   "&Random Season"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuPoint 
         Caption         =   "&Point Editor..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuReset 
         Caption         =   "R&eset all records..."
      End
      Begin VB.Menu mnuTrackSettings 
         Caption         =   "&Advanced Track Settings..."
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuTopHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help..."
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "On the &Web"
         Begin VB.Menu mnuTHHome 
            Caption         =   "Track Handler HomePage"
         End
         Begin VB.Menu mnuSep13 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGrandPrix1 
            Caption         =   "GrandPrix1.com"
         End
         Begin VB.Menu mnuGrandPrix2 
            Caption         =   "GrandPrix2.com"
         End
         Begin VB.Menu mnuGP3 
            Caption         =   "GP3.org"
         End
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuCCCarSetup 
         Caption         =   "&Advanced Track Settings"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFile As String
Dim ChangeLap As Boolean
Dim Dott As Boolean
Dim RetVal
Dim CheckBatFile As String
Dim Support As Boolean
Dim File As String

Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Enum PatFile
    Dat = 0
    Bmp = 1
    Gif = 2
    All = 3
End Enum

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Sub A_Click(Index As Integer)
    If A(Index).Tag = "On" Then
        A(Index).Tag = "Off"
        A(Index).Picture = LoadResPicture(108 + Index, 0)
    Else
        A(Index).Tag = "On"
        A(Index).Picture = LoadResPicture(101 + Index, 0)
    End If
End Sub

Private Sub AC_Click(Index As Integer)
    If AC(Index).Tag = "On" Then
        AC(Index).Tag = "Off"
        AC(Index).Picture = LoadResPicture(108 + Index, 0)
    Else
        AC(Index).Tag = "On"
        AC(Index).Picture = LoadResPicture(101 + Index, 0)
    End If
End Sub

Private Sub chkNoLimit_Click()
    If chkNoLimit.Value = 1 Then
        lblPit.Enabled = False
        lblPitSpeed.Enabled = False
        hscPitSpeed.Enabled = False
    Else
        lblPit.Enabled = True
        lblPitSpeed.Enabled = True
        hscPitSpeed.Enabled = True
    End If
End Sub

Private Sub chkUPower_Click()
    If chkUPower.Value = 1 Then
        hscPQPower.Enabled = False
        hscPRPower.Enabled = False
        lblPRPower.Enabled = False
        lblPQPower.Enabled = False
        Label2(2).Enabled = False
        Label10.Enabled = False
    Else
        hscPQPower.Enabled = True
        hscPRPower.Enabled = True
        lblPRPower.Enabled = True
        lblPQPower.Enabled = True
        Label2(2).Enabled = True
        Label10.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    frmAddTime.Show vbModal, frmMain
    LoadTimeData
End Sub

Private Sub cmdDefaultSettings_Click()
    Slider1.Value = 1994
    hscPGrip.Value = 198
    hscPQPower.Value = 790
    hscPRPower.Value = 780
    chkSave.Value = 0
    chkNoLimit.Value = 0
    chk0as1.Value = 1
    hscQRace.Value = 5
    hscCWeight = 1313
    hscWeight.Value = 1313
    hscPitSpeed.Value = 50
    chkUPower.Value = 0
    DriveHelpDefault
End Sub

Private Sub cmdDelete_Click()
Dim TempNr
    frmMain.MousePointer = 11
    TempNr = lstTime.SelectedItem.Index
    DeleteTime
    If lstTime.ListItems.Count > 0 Then
        If lstTime.ListItems.Count < TempNr Then TempNr = TempNr - 1
        lstTime.ListItems(TempNr).Selected = True
    End If
    frmMain.MousePointer = 0
End Sub

Private Sub cmdExportSettings_Click()
    GP2FileNum = FreeFile
    Open GP2Dir & "\gp2.exe" For Binary As GP2FileNum
    SaveMisc
    ExportCarHelp
    ExportCWeight
    ExportLevel
    ExportNullAsOne
    ExportSaveLap
    ExportPGrip
    ExportPQPower
    ExportPRPower
    ExportPWeight
    ExportUseTeam
    ExportSpeed
    Read = oMisc.File_Exists(GP2Dir & "\f1gstate.sav")
    If Read = True Then
        F1SaveFileNum = FreeFile
        Open GP2Dir & "\f1gstate.sav" For Binary As F1SaveFileNum
        ExportQuickRace
        Close F1SaveFileNum
    End If
    Close GP2FileNum
    GetMisc
End Sub

Private Sub cmdImportSettings_Click()
    GP2FileNum = FreeFile
    Open GP2Dir & "\gp2.exe" For Binary As GP2FileNum
    ImportCWeight
    ImportGameSettings
    ImportLevel
    ImportNullAsOne
    ImportPGrip
    ImportPQPower
    ImportPRPower
    ImportPWeight
    ImportSaveLap
    ImportSpeed
    ImportUseTeam
    Read = oMisc.File_Exists(GP2Dir & "\f1gstate.sav")
    If Read = True Then
        F1SaveFileNum = FreeFile
        Open GP2Dir & "\f1gstate.sav" For Binary As F1SaveFileNum
        ImportQuick
        Close F1SaveFileNum
    End If
    Close GP2FileNum
    GetMisc
End Sub

Private Sub cmdJam_Click()
    On Error GoTo ErrHandler
    lstFile.SetFocus
    If (lstFile.Tag = ".dat") And (lstFile.SelectedItem.Text <> "") Then
        CheckJam lstFile.SelectedItem.Key
    End If
Exit Sub
ErrHandler:
End Sub

Private Sub cmdJamCheck_Click()
    CheckJam txtPath.Text
End Sub

Public Sub cmdSaveGP2Info_Click()
Dim Ret As Boolean
Dim Path As String
    frmMain.MousePointer = 11
    On Error Resume Next
    MakeText
    Kill (ProgramDir & "\Bat\gp2info.bat")
    Path = lstFile.SelectedItem.Key
    Path = oMisc.GetShortName(Path)
    For X = Len(Path) To 0 Step -1
        If Mid(Path, X, 1) = "\" Then Exit For
    Next
    Read = Mid(Path, X + 1)
    Path = Mid(Path, 1, X - 1)
    
    FileNum = FreeFile
    Open ProgramDir & "\Bat\gp2info.bat" For Append As FileNum
    Print #FileNum, "@echo off"
    Print #FileNum, "cd " & Path
    Print #FileNum, Mid(Dir1.Path, 1, 2)
    Print #FileNum, "thcheck " & Read
    Ret = oMisc.File_Exists(Path & "\thcheck.exe")
    If Ret = False Then
        FileCopy ProgramDir & "\gp2utils\thcheck.exe", Path & "\thcheck.exe"
        Print #FileNum, "del thcheck.exe"
    End If
    Print #FileNum, "cls"
    Print #FileNum, "echo You can now close this window."
    Close FileNum
    RetVal = ShellExecute(frmMain.hwnd, "open", ProgramDir & "\Bat\gp2info.bat", vbNullString, vbNullString, 1)
    frmMain.MousePointer = 0
End Sub

Private Sub cmdSaveQual_Click()
    frmMain.MousePointer = 11
    Read = txtQTime.Text & ";" & txtQDate.Text & ";Qual;" & txtQDriver.Text & ";" & txtQTeam.Text & ";" & txtName
    oDB.SaveNew dbFile, Read
    LoadTimeData
    frmMain.MousePointer = 0
End Sub

Private Sub cmdSaveRace_Click()
    frmMain.MousePointer = 11
    Read = txtRTime.Text & ";" & txtRDate.Text & ";Race;" & txtRDriver.Text & ";" & txtRTeam.Text & ";" & txtName
    oDB.SaveNew dbFile, Read
    LoadTimeData
    frmMain.MousePointer = 0
End Sub

Private Sub Dir1_Change()
    lstFile.Tag = ""
    LoadFiles All
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrHandler
    Dir1.Path = Drive1.Drive
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 68
        MsgBox LoadResString(109), vbExclamation, TH
        Drive1.Drive = "C:"
    Case Else
        Print #Log, Date & " " & Time & " Drive1_Change, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub Form_Load()
    frmMain.MousePointer = 11
    On Error Resume Next
    NewTree
    tabMain.TabEnabled(1) = False
    tabMain.TabEnabled(2) = False
    tabMain.Tab = 0
    ProgramDir = App.Path '"d:"
    Set oMisc = New Misc
    Set oData = New GP2Info
    Set oReg = New oReg
    Set oDB = New clsDB
    MkDir ProgramDir & "\File"
    MkDir ProgramDir & "\Bat"
    dbFile = ProgramDir & "\Time.lda"
    GetResText
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "Nr")
    X = X + 1
    oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "Nr", Trim(Str(X))
    Read = ""
    Read = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "TrackPath")
    If Read <> "" Then
        Drive1.Drive = Mid(Read, 1, 2)
        Dir1.Path = Read
    End If

    'Check if Toolbar is on or off, the same with status bar, if off the hide
    X = 0
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "Toolbar")
    If X = 1 Then
        mnuToolbar.Checked = True
        mnuToolbar_Click
    End If
    X = 0
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "Statusbar")
    If X = 1 Then
        mnuStatusbar.Checked = True
        mnuStatusbar_Click
    End If
    GP2Dir = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "GP2Path")
    If GP2Dir = "" Then
        Read = SetGP2Dir(False)
        If Read = False Then mnuExit_Click
    Else
        stbMain.Panels(3).Text = stbMain.Panels(3).Text & " " & GP2Dir
    End If
    RegFileName
    GetGP2Version
    'App.HelpFile = ProgramDir & "\Help.hlp"

    X = GetWindowLong(txtLaps.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtLaps.hwnd, GWL_STYLE, X)

    X = GetWindowLong(txtLength.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtLength.hwnd, GWL_STYLE, X)

    X = GetWindowLong(txtTire.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtTire.hwnd, GWL_STYLE, X)

    X = GetWindowLong(lblLen.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(lblLen.hwnd, GWL_STYLE, X)
    
    X = GetWindowLong(lblWare.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(lblWare.hwnd, GWL_STYLE, X)

    X = GetWindowLong(lblLaps.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(lblLaps.hwnd, GWL_STYLE, X)
    
    X = GetWindowLong(txtQTime.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtQTime.hwnd, GWL_STYLE, X)

    X = GetWindowLong(txtRTime.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtRTime.hwnd, GWL_STYLE, X)

    X = GetWindowLong(txtQDate.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtQDate.hwnd, GWL_STYLE, X)

    X = GetWindowLong(txtRDate.hwnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtRDate.hwnd, GWL_STYLE, X)

    Call SendMessageLong(lstTime.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, -1)
    Call SendMessageLong(lstFile.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, -1)

    fraFileInfo.Enabled = False
    LoadTimeData

    Log = FreeFile
    Open ProgramDir & "\THLog.txt" For Append As FileNum
    
    X = InStr(1, Command(), ".ths")
    If (Command() <> "") And (X > 0) Then
        OpenCommandFile
    Else
        'Normal start
        NewFile
        LoadRecent
        DriveHelpDefault
    End If
    frmMain.MousePointer = 0
Exit Sub

ErrHandler:
    Print #Log, Date & " " & Time & " frmMain_Load, Error Number: " & Err.Number & ", Error Description: " & Err.Description
    frmMain.MousePointer = 0
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.Description & vbCrLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Read = CheckIfSave
    If Read = "Cancel" Then
        Cancel = 1
        Exit Sub
    End If
    On Error Resume Next
    Set oMisc = Nothing
    Set oData = Nothing
    Kill (CheckBatFile)
    Kill (ProgramDir & "\File\*.*")
    Kill (TempFile)
    Close Log
    End
End Sub

Private Sub hscCWeight_Change()
    X = hscCWeight.Value
    X = X / 2.203020134
    lblCWeight.Caption = Str(hscCWeight.Value) + "lb (" + Trim(Str(X)) + "kg)"
End Sub

Private Sub hscCWeight_Scroll()
    X = hscCWeight.Value
    X = X / 2.203020134
    lblCWeight.Caption = Str(hscCWeight.Value) + "lb (" + Trim(Str(X)) + "kg)"
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

Private Sub hscQRace_Change()
    lblQuick.Caption = hscQRace.Value & "%"
End Sub

Private Sub hscQRace_Scroll()
    lblQuick.Caption = hscQRace.Value & "%"
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

Private Sub lblAuthor_Click()
    MsgBox LoadResString(101), vbInformation, TH
End Sub

Private Sub lblCountry_GotFocus()
    TextSelected
End Sub

Private Sub lblEvent_GotFocus()
    TextSelected
End Sub

Private Sub lblInfoYear_GotFocus()
    TextSelected
End Sub

Private Sub lblLaps_GotFocus()
    TextSelected
End Sub

Private Sub lblLen_GotFocus()
    TextSelected
End Sub

Private Sub lblMisc_GotFocus()
    TextSelected
End Sub

Private Sub lblQual_GotFocus()
    TextSelected
End Sub

Private Sub lblRace_GotFocus()
    TextSelected
End Sub

Private Sub lblSlot_GotFocus()
    TextSelected
End Sub

Private Sub lblTrackName_GotFocus()
    TextSelected
End Sub

Private Sub lblWare_GotFocus()
    TextSelected
End Sub

Private Sub lblYear_Click()
    Read = InputBox("Year for this Season", "Select Year", lblYear.Caption)
    If Read <> "" Then Slider1.Value = Read
End Sub

Private Sub lstFile_Click()
Dim PicY As Long
Dim PicX As Long
    On Error GoTo ErrHandler
    If lstFile.SelectedItem.Text <> "" Then
        Read = lstFile.SelectedItem.Key
        If GetExt(Read) = ".dat" Then
            fraMenuPic.Visible = False
            fraFileInfo.Visible = True
            lstFile.Tag = ".dat"
            ClearInfo
            Support = ReadGP2Info(Read)
            If Support = True Then
                fraFileInfo.Enabled = True
                cmdSaveGP2Info.Enabled = True
            Else
                fraFileInfo.Enabled = False
                cmdSaveGP2Info.Enabled = False
            End If
        ElseIf GetExt(Read) = ".bmp" Then
            fraFileInfo.Enabled = False
            fraFileInfo.Visible = False
            fraMenuPic.Visible = True
            imgRealSize.Picture = LoadPicture(Read)
            PicY = imgRealSize.Height / 15
            PicX = imgRealSize.Width / 15
            If ((PicX = 640) And (PicY = 480)) Then
                lblPicInfo.Caption = "Large Menu Picture"
                Set imgPre.Picture = LoadPicture(Read)
                lstFile.Tag = "big"
                Support = True
            ElseIf ((PicX = 440) And (PicY = 330)) Then
                lblPicInfo.Caption = "Smal Menu Picture"
                Set imgPre.Picture = LoadPicture(Read)
                lstFile.Tag = "smal"
                Support = True
            Else
                MsgBox LoadResString(110), vbInformation, TH
                Support = False
                Exit Sub
            End If
        End If
    End If
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 91
        Exit Sub
    Case Else
        Print #Log, Date & " " & Time & " lstFile_Click, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub lstFile_ItemClick(ByVal Item As ComctlLib.ListItem)
    If File <> lstFile.SelectedItem.Text Then
        lstFile_Click
        File = lstFile.SelectedItem.Text
    End If
End Sub

Private Sub lstFile_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 40) Or (KeyCode = 38) Then lstFile_Click
End Sub

Private Sub lstFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub lstFile_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
    If Support = True Then Data.SetData lstFile.SelectedItem.Text, 1
End Sub

Private Sub mnuAbout_Click()
    frmMain.MousePointer = 11
    frmAbout.Show vbModal, frmMain
    frmMain.MousePointer = 0
End Sub

Private Sub mnuCCCarSetup_Click()
    On Error Resume Next
    frmCCSetup.Show vbModal, frmMain
End Sub

Private Sub mnuConvert_Click()
    frmConvert.Show vbModal, frmMain
End Sub

Private Sub mnuExit_Click()
    Read = CheckIfSave
    If Read = "Cancel" Then Exit Sub
    On Error Resume Next
    Set oMisc = Nothing
    Set oData = Nothing
    Kill (ProgramDir & "\Bat\*.*")
    Kill (ProgramDir & "\File\*.*")
    Kill (TempFile)
    Close Log
    End
End Sub

Private Sub mnuExport_Click()
    SaveTrackData TreeNr
    SaveMisc
    frmExport.Show vbModal, frmMain
End Sub

Private Sub mnuGP2Path_Click()
    SetGP2Dir True
    GetGP2Version
End Sub

Private Sub mnuGP3_Click()
    Read = "http://www.gp3.org/"
    RetVal = ShellExecute(frmMain.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuGrandPrix1_Click()
    Read = "http://www.grandprix1.com/"
    RetVal = ShellExecute(frmMain.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuGrandPrix2_Click()
    Read = "http://www.grandprix2.com/"
    RetVal = ShellExecute(frmMain.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuHelp_Click()
    Read = ProgramDir & "\Help\Index.htm"
    RetVal = ShellExecute(frmMain.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuImport_Click()
    SaveTrackData TreeNr
    SaveMisc
    Read = CheckIfSave
    If Read = "Cancel" Then Exit Sub
    frmImport.Show vbModal, frmMain
End Sub

Private Sub mnuNew_Click()
    Read = CheckIfSave
    If Read = "Cancel" Then Exit Sub
    TreeView1.Nodes.Item(1).Selected = True
    TreeView1_NodeClick TreeView1.Nodes(1)
    NewFile
    LoadFile
    frmMain.Caption = "GP2 Track Handler v1.4"
End Sub

Private Sub mnuOpen_Click()
    OpenFile
End Sub

Private Sub mnuOpen1_Click()
    On Error GoTo ErrHandler
    Randomize
    X = Int((500) * Rnd)
    Kill (TempFile)
    TempFile = ProgramDir & "\File\th14" & Trim(Str(X)) & ".lda"
    FileCopy mnuOpen1.Tag, TempFile
    
    FileInfo.Name = mnuOpen1.Caption
    FileInfo.Path = mnuOpen1.Tag
    FileInfo.Saved = True
    FileInfo.Import = False
    frmMain.Caption = "GP2 Track Handler v1.4 [" & Trim(FileInfo.Name) & "]"
    LoadFile
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 53
        MsgBox LoadResString(111), vbExclamation, TH
    Case Else
        Print #Log, Date & " " & Time & " mnuOpen1_Click, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub mnuOpen2_Click()
    On Error GoTo ErrHandler
    Randomize
    X = Int((500) * Rnd)
    Kill (TempFile)
    TempFile = ProgramDir & "\File\th14" & Trim(Str(X)) & ".lda"
    FileCopy mnuOpen2.Tag, TempFile
    
    FileInfo.Name = mnuOpen2.Caption
    FileInfo.Path = mnuOpen2.Tag
    FileInfo.Saved = True
    FileInfo.Import = False
    LoadFile
    Read = oMisc.RecentFile(OpenRecent, , , 2)
    LoadRecent
    frmMain.Caption = "GP2 Track Handler v1.4 [" & Trim(FileInfo.Name) & "]"
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 53
        MsgBox LoadResString(111), vbExclamation, TH
    Case Else
        Print #Log, Date & " " & Time & " mnuOpen2_Click, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub mnuOpen3_Click()
    On Error GoTo ErrHandler
    Randomize
    X = Int((500) * Rnd)
    Kill (TempFile)
    TempFile = ProgramDir & "\File\th14" & Trim(Str(X)) & ".lda"
    FileCopy mnuOpen3.Tag, TempFile
    
    FileInfo.Name = mnuOpen3.Caption
    FileInfo.Path = mnuOpen3.Tag
    FileInfo.Saved = True
    FileInfo.Import = False
    LoadFile
    Read = oMisc.RecentFile(OpenRecent, , , 3)
    LoadRecent
    frmMain.Caption = "GP2 Track Handler v1.4 [" & Trim(FileInfo.Name) & "]"
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 53
        MsgBox LoadResString(111), vbExclamation, TH
    Case Else
        Print #Log, Date & " " & Time & " mnuOpen3_Click, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub mnuPoint_Click()
    frmPoint.Show vbModal, frmMain
End Sub

Private Sub mnuRand_Click()
    frmMain.MousePointer = 11
    LoadFiles Dat
    If lstFile.ListItems.Count > 15 Then
        TreeView1_NodeClick TreeView1.Nodes(1)
        TreeView1.Nodes(1).Selected = True
        RandomTracks lstFile.ListItems.Count
        FileInfo.Changes = True
    Else
        MsgBox "You need to have 16 track's or more in this directory to make a random season.", vbInformation, TH
    End If
    LoadFiles All
    frmMain.MousePointer = 0
End Sub

Private Sub mnuReset_Click()
    SaveTrackData TreeNr
    Read = InputBox("Set all records to:", "Reset Records", "3:00.000")
    ImportTime True, Read
    GetTrackData TreeNr
End Sub

Private Sub mnuSave_Click()
    SaveTrackData TreeNr
    SaveMisc
    SaveFile
End Sub

Private Sub mnuSaveas_Click()
    SaveTrackData TreeNr
    SaveMisc
    SaveFileAs
End Sub

Private Sub mnuStatusbar_Click()
    If mnuStatusbar.Checked = True Then
        stbMain.Visible = False
        mnuStatusbar.Checked = False
        frmMain.Height = frmMain.Height - 255
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "Statusbar", , "1"
    Else
        stbMain.Visible = True
        mnuStatusbar.Checked = True
        frmMain.Height = frmMain.Height + 255
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "Statusbar", , "0"
    End If
End Sub

Private Sub mnuTHHome_Click()
    Read = "http://hem1.passagen.se/formula1/"
    RetVal = ShellExecute(frmMain.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub mnuToolbar_Click()
    If mnuToolbar.Checked = True Then
        Toolbar1.Visible = False
        mnuToolbar.Checked = False
        TreeView1.Top = 60
        tabMain.Top = 60
        frmMain.Height = frmMain.Height - 420
        Frame3.Visible = False
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "Toolbar", , 1
    Else
        Toolbar1.Visible = True
        mnuToolbar.Checked = True
        TreeView1.Top = 480
        tabMain.Top = 480
        frmMain.Height = frmMain.Height + 420
        Frame3.Visible = True
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "Toolbar", , 0
    End If
End Sub

Private Sub mnuTrackPath_Click()
    szTitle = "Select Track Directory"
    With tBrowseInfo
        .hwndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "TrackPath", sBuffer
    End If
    Drive1.Drive = Mid(sBuffer, 1, 2)
    Dir1.Path = sBuffer
End Sub

Private Sub mnuTrackSettings_Click()
    On Error Resume Next
    frmCCSetup.Show vbModal, frmMain
End Sub

Private Sub P_Click(Index As Integer)
    If P(Index).Tag = "On" Then
        P(Index).Tag = "Off"
        P(Index).Picture = LoadResPicture(108 + Index, 0)
    Else
        P(Index).Tag = "On"
        P(Index).Picture = LoadResPicture(101 + Index, 0)
    End If
End Sub

Private Sub R_Click(Index As Integer)
    If R(Index).Tag = "On" Then
        R(Index).Tag = "Off"
        R(Index).Picture = LoadResPicture(108 + Index, 0)
    Else
        R(Index).Tag = "On"
        R(Index).Picture = LoadResPicture(101 + Index, 0)
    End If
End Sub

Private Sub S_Click(Index As Integer)
    If S(Index).Tag = "On" Then
        S(Index).Tag = "Off"
        S(Index).Picture = LoadResPicture(108 + Index, 0)
    Else
        S(Index).Tag = "On"
        S(Index).Picture = LoadResPicture(101 + Index, 0)
    End If
End Sub

Private Sub Slider1_Change()
    lblYear.Caption = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
    lblYear.Caption = Slider1.Value
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tab = 0 Then
        mnuRand.Enabled = True
    Else
        mnuRand.Enabled = False
    End If
    SaveMisc
    GetMisc
End Sub

Private Sub tabPic_GotFocus()
    If tabMain.Tab = 1 Then
        txtName.SetFocus
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    Case "Exit"
        mnuExit_Click
    Case "Save"
        mnuSave_Click
    Case "Open"
        mnuOpen_Click
    Case "New"
        mnuNew_Click
    Case "Import"
        mnuImport_Click
    Case "Export"
        mnuExport_Click
    Case "GP2"
        RetVal = ShellExecute(frmMain.hwnd, "open", "gp2.exe", vbNullString, GP2Dir, 1)
    Case "GP2Exe"
        AddExe
    End Select
End Sub

Public Sub DriveHelpDefault()
    For X = 0 To 6
        R(X).Picture = LoadResPicture(101 + X, 0)
        R(X).Tag = "On"
    Next
    A(0).Picture = LoadResPicture(108, 0)
    A(0).Tag = "Off"
    For X = 1 To 6
        A(X).Picture = LoadResPicture(101 + X, 0)
        A(X).Tag = "On"
    Next
    
    S(0).Picture = LoadResPicture(108, 0)
    S(0).Tag = "Off"
    For X = 1 To 6
        S(X).Picture = LoadResPicture(101 + X, 0)
        S(X).Tag = "On"
    Next

    P(0).Picture = LoadResPicture(108, 0)
    P(0).Tag = "Off"
    P(1).Picture = LoadResPicture(102, 0)
    P(1).Tag = "On"
    P(2).Picture = LoadResPicture(110, 0)
    P(2).Tag = "Off"
    P(3).Picture = LoadResPicture(111, 0)
    P(3).Tag = "Off"
    For X = 4 To 6
        P(X).Picture = LoadResPicture(101 + X, 0)
        P(X).Tag = "On"
    Next

    AC(0).Picture = LoadResPicture(108, 0)
    AC(0).Tag = "Off"
    AC(1).Picture = LoadResPicture(102, 0)
    AC(1).Tag = "On"
    AC(6).Picture = LoadResPicture(107, 0)
    AC(6).Tag = "On"
    For X = 2 To 5
        AC(X).Picture = LoadResPicture(108 + X, 0)
        AC(X).Tag = "Off"
    Next
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandler
    If (KeyCode = 46) And (TreeView1.SelectedItem.Key <> "r") Then
        If Len(TreeView1.SelectedItem.Key) <> 3 Then
            TreeView1.Nodes.Item(TreeView1.SelectedItem.Parent.Index).Selected = True
            Do Until TreeView1.SelectedItem.Children = 0
                TreeView1.Nodes.Remove (TreeView1.SelectedItem.Child.Index)
            Loop
            TreeView1.SelectedItem.Text = "Track " & Mid(TreeView1.SelectedItem.Key, 2, 2) - 10
        Else
            TreeView1.SelectedItem.Text = "Track " & Mid(TreeView1.SelectedItem.Key, 2, 2) - 10
            Do Until TreeView1.SelectedItem.Children = 0
                TreeView1.Nodes.Remove (TreeView1.SelectedItem.Child.Index)
            Loop
        End If
        With frmMain
            .txtAdjectiv = ""
            .txtName = ""
            .txtCountry = ""
            .txtLaps = ""
            .txtLength = ""
            .txtPath = ""
            .txtTire = ""
            .txtQDate = ""
            .txtRDate = ""
            .txtQTime = ""
            .txtRTime = ""
            .txtQTeam = ""
            .txtRTeam = ""
            .txtQDriver = ""
            .txtRDriver = ""
        End With
        SaveTrackData TreeNr
    End If
Exit Sub
ErrHandler:
    Exit Sub
End Sub

Public Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
    If Node.Key = "r" Then
        SaveTrackData TreeNr
        tabMain.TabEnabled(1) = False
        tabMain.TabEnabled(2) = False
        tabMain.Tab = 0
        TreeNr = 0
    Else
        If TreeNr <> 0 Then SaveTrackData TreeNr
        tabMain.TabEnabled(1) = True
        tabMain.TabEnabled(2) = True

        TreeNr = Mid(TreeView1.SelectedItem.Key, 2, 2)
        TreeNr = TreeNr - 10
        GetTrackData TreeNr

        If InStr(1, TreeView1.SelectedItem.Key, "BPic") Then
            tabMain.Tab = 2
            tabPic.Tab = 0
        ElseIf InStr(1, TreeView1.SelectedItem.Key, "SPic") Then
            tabMain.Tab = 2
            tabPic.Tab = 1
        Else
            tabMain.Tab = 1
        End If
        If (TreeView1.SelectedItem.Children = 0) And (Len(TreeView1.SelectedItem.Key) = 3) Then
            txtName.Enabled = False
            txtCountry.Enabled = False
            txtAdjectiv.Enabled = False
            txtLaps.Enabled = False
            txtLength.Enabled = False
            txtQDate.Enabled = False
            txtQDriver.Enabled = False
            txtQTeam.Enabled = False
            txtQTime.Enabled = False
            txtRDate.Enabled = False
            txtRDriver.Enabled = False
            txtRTeam.Enabled = False
            txtRTime.Enabled = False
            txtTire.Enabled = False
            cmdJamCheck.Enabled = False
            vscLaps.Enabled = False
            cmdSaveQual.Enabled = False
            cmdSaveRace.Enabled = False
        Else
            txtName.Enabled = True
            txtCountry.Enabled = True
            txtAdjectiv.Enabled = True
            txtLaps.Enabled = True
            txtLength.Enabled = True
            txtQDate.Enabled = True
            txtQDriver.Enabled = True
            txtQTeam.Enabled = True
            txtQTime.Enabled = True
            txtRDate.Enabled = True
            txtRDriver.Enabled = True
            txtRTeam.Enabled = True
            txtRTime.Enabled = True
            txtTire.Enabled = True
            cmdJamCheck.Enabled = True
            vscLaps.Enabled = True
            cmdSaveQual.Enabled = True
            cmdSaveRace.Enabled = True
        End If
    End If
End Sub

Private Sub TreeView1_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tabMain.Tab = 4 Then
        DropTime Data, Effect, Button, Shift, X, Y
    ElseIf tabMain.Tab = 0 Then
        DropTrack Data, Effect, Button, Shift, X, Y
    End If
End Sub

Private Sub TreeView1_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
End Sub

Private Sub txtAdjectiv_GotFocus()
    TextSelected
End Sub

Private Sub txtCarShape_GotFocus()
    TextSelected
End Sub

Private Sub txtCountry_GotFocus()
    TextSelected
End Sub

Private Sub txtLaps_Change()
    On Error Resume Next
    ChangeLap = False
    If txtLaps = "" Then Exit Sub
    If txtLaps < 126 Then
        vscLaps.Value = txtLaps.Text
    Else
        vscLaps.Value = 126
    End If
    ChangeLap = True
End Sub

Private Sub txtLaps_GotFocus()
    TextSelected
End Sub

Private Sub txtLaps_LostFocus()
    If tabMain.Tab = 1 Then
        If txtLaps > 126 Then txtLaps = 126
        If txtLaps < 3 Then txtLaps = 3
    End If
End Sub

Private Sub txtLength_GotFocus()
    TextSelected
End Sub

Private Sub txtName_GotFocus()
    If tabMain.Tab = 0 Then
        Drive1.SetFocus
    ElseIf tabMain.Tab = 1 Then
        TextSelected
    End If
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo ErrHandler
    If TreeView1.SelectedItem.Key <> "r" Then
        If TreeView1.SelectedItem.Children > 0 Then
            TreeNr = Mid(TreeView1.SelectedItem.Key, 2, 2)
            TreeNr = TreeNr - 10
            TreeView1.SelectedItem.Text = TreeNr & ". " & txtName.Text
        Else
            TreeNr = Mid(TreeView1.SelectedItem.Parent.Key, 2, 2)
            TreeNr = TreeNr - 10
            TreeView1.SelectedItem.Parent.Text = TreeNr & ". " & txtName.Text
        End If
    End If
Exit Sub

ErrHandler:
    Select Case Err.Number
        Case 91
            Exit Sub
        Case Else
        Print #Log, Date & " " & Time & " txtName_LostFocus, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub txtQDate_GotFocus()
    TextSelected
End Sub

Private Sub txtQDriver_GotFocus()
    TextSelected
End Sub

Private Sub txtQTeam_GotFocus()
    TextSelected
End Sub


Private Sub txtQTime_GotFocus()
    TextSelected
End Sub

Private Sub txtRDate_GotFocus()
    TextSelected
End Sub

Private Sub txtRDriver_GotFocus()
    TextSelected
End Sub

Private Sub txtRTeam_GotFocus()
    TextSelected
End Sub

Private Sub txtRTime_GotFocus()
    TextSelected
End Sub

Private Sub txtTire_GotFocus()
    TextSelected
End Sub

Public Sub SaveDropData(ByVal Path As String)
Dim TreeNr As Integer
    TreeNr = Mid(TreeView1.SelectedItem.Key, 2, 2)
    TreeNr = TreeNr - 10
    txtPath.Text = Path
    txtCountry.Text = lblCountry.Text
    txtLaps.Text = lblLaps.Text
    txtName.Text = lblTrackName.Text
    txtTire.Text = lblWare.Text
    txtLength.Text = lblLen.Text
    txtAdjectiv.Text = GetAdjectiv(Trim(txtCountry.Text))
    txtRTime.Text = lblRace.Text
    txtQTime.Text = lblQual.Text

    SaveTrackData TreeNr
    TreeView1.SelectedItem.Text = TreeNr & ". " & lblTrackName.Text
End Sub

Public Function SetGP2Dir(Optional Change As Boolean) As Boolean
    szTitle = "Select GP2 Location"
    With tBrowseInfo
        .hwndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        GP2Dir = sBuffer
    End If
    If GP2Dir = "" Then
        SetGP2Dir = False
        Exit Function
    End If
    If Len(GP2Dir) = 3 Then GP2Dir = Mid(GP2Dir, 1, 2)
    Read = oMisc.File_Exists(GP2Dir & "\gp2.exe")
    If Read = False Then
        Responce = MsgBox(LoadResString(105), vbRetryCancel, TH)
        If Responce = vbCancel Then
            If Change = True Then
                Exit Function
            Else
                End
            End If
        Else
            SetGP2Dir
        End If
    Else
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "GP2Path", GP2Dir
        stbMain.Panels(3).Text = "GP2 Directory: " & GP2Dir
    End If
End Function

Public Sub LoadRecent()
Dim Name1 As String
Dim Name2 As String
Dim Name3 As String
Dim Path1 As String
Dim Path2 As String
Dim Path3 As String

    Name1 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Name1")
    Name2 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Name2")
    Name3 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Name3")
    
    Path1 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Path1")
    Path2 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Path2")
    Path3 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Path3")
    If Name1 <> "" Then
        mnuOpen1.Visible = True
        mnuOpen1.Caption = Name1
        mnuOpen1.Tag = Path1
        mnuSep3.Visible = True
    End If
    If Name2 <> "" Then
        mnuOpen2.Visible = True
        mnuOpen2.Caption = Name2
        mnuOpen2.Tag = Path2
    End If
    If Name3 <> "" Then
        mnuOpen3.Visible = True
        mnuOpen3.Caption = Name3
        mnuOpen3.Tag = Path3
    End If
End Sub

Public Function CheckIfSave() As String
    If Trim(FileInfo.Path) <> "" Then
        FileNum = FreeFile
        Open Trim(FileInfo.Path) For Binary As FileNum
        Read = String(FileLen(Trim(FileInfo.Path)), " ")
        Get #FileNum, 1, Read
        Close FileNum
        FileNum = FreeFile
        Open TempFile For Binary As FileNum
        Read2 = String(FileLen(TempFile), " ")
        Get #FileNum, 1, Read2
        Close FileNum
        If Read = Read2 Then
            CheckIfSave = ""
            Exit Function
        End If
        Responce = MsgBox(LoadResString(106), vbYesNoCancel, TH)
        If Responce = vbNo Then
            CheckIfSave = ""
        ElseIf Responce = vbCancel Then
            CheckIfSave = "Cancel"
        Else
            CheckIfSave = ""
            mnuSave_Click
        End If
        Exit Function
    End If
    If FileInfo.Changes = True Then
        Responce = MsgBox(LoadResString(106), vbYesNoCancel, TH)
        If Responce = vbNo Then
            CheckIfSave = ""
        ElseIf Responce = vbCancel Then
            CheckIfSave = "Cancel"
        Else
            CheckIfSave = ""
            mnuSave_Click
        End If
    End If
End Function

Private Sub vscLaps_Change()
    If ChangeLap = True Then
        If vscLaps.Value > 2 Then
            txtLaps.Text = vscLaps.Value
        Else
            vscLaps.Value = vscLaps.Value + 1
        End If
        txtLaps.SetFocus
    End If
End Sub

Private Sub txtQTime_Change()
    If Dott = True Then
        If Len(txtQTime.Text) = 1 Then
            txtQTime.Text = txtQTime.Text & ":"
        ElseIf Len(txtQTime.Text) = 4 Then
            txtQTime.Text = txtQTime.Text & "."
        End If
        SendKeys ("^{END}")
    End If
End Sub

Private Sub txtQTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode <> 46) And (KeyCode <> 8) Then
        Dott = True
    Else
        Dott = False
    End If
End Sub

Private Sub txtRTime_Change()
    If Dott = True Then
        If Len(txtRTime.Text) = 1 Then
            txtRTime.Text = txtRTime.Text & ":"
        ElseIf Len(txtRTime.Text) = 4 Then
            txtRTime.Text = txtRTime.Text & "."
        End If
        SendKeys ("^{END}")
    End If
End Sub

Private Sub txtRTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode <> 46) And (KeyCode <> 8) Then
        Dott = True
    Else
        Dott = False
    End If
End Sub

Private Sub txtRDate_Change()
    If Dott = True Then
        If Len(txtRDate.Text) = 4 Then
            txtRDate.Text = txtRDate.Text & "-"
        ElseIf Len(txtRDate.Text) = 7 Then
            txtRDate.Text = txtRDate.Text & "-"
        End If
        SendKeys ("^{END}")
    End If
End Sub

Private Sub txtRDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode <> 46) And (KeyCode <> 8) Then
        Dott = True
    Else
        Dott = False
    End If
End Sub

Private Sub txtQDate_Change()
    If Dott = True Then
        If Len(txtQDate.Text) = 4 Then
            txtQDate.Text = txtQDate.Text & "-"
        ElseIf Len(txtQDate.Text) = 7 Then
            txtQDate.Text = txtQDate.Text & "-"
        End If
        SendKeys ("^{END}")
    End If
End Sub

Private Sub txtQDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode <> 46) And (KeyCode <> 8) Then
        Dott = True
    Else
        Dott = False
    End If
End Sub

Public Sub MakeText()
    Read = ""
    Read2 = ""
    Read = "#GP2INFO|Name|" & lblTrackName & "|Country|" & lblCountry & "|Created|Created by Track Editor written by Paul Hoad see (License.txt about distributing this track)|Author|" & lblAuthor & _
        "|Year|" & lblInfoYear & "|Event|" & lblEvent & "|Desc|" & lblMisc & _
        "|Laps|" & lblLaps & "|Slot|" & lblSlot & "|Tyre|" & lblWare & "|LengthMeters|" & lblLen
    If lblRace = "" Then
        Read = Read & "|LapRecord|None Entered"
    Else
        Read = Read & "|LapRecord|" & lblRace & "|"
    End If
    If lblQual = "" Then
        Read = Read & "|LapRecordQualify|None Entered|"
    Else
        Read = Read & "|LapRecordQualify|" & lblQual & "|"
    End If
    Read2 = String(3900 - Len(Read), Chr(0))
    Read = Read & Read2

    FileNum = FreeFile
    SetAttr lstFile.SelectedItem.Key, vbNormal
    Open lstFile.SelectedItem.Key For Binary As FileNum
    Put #FileNum, 1, Read
    Close FileNum
End Sub

Private Sub OpenCommandFile()
Dim GetOpen As String
    On Error GoTo ErrHandler

    For X = Len(Command()) To 0 Step -1
        If Mid(Command(), X, 1) = "\" Then Exit For
    Next
    Read2 = ""
    Read2 = Mid(Command(), X + 1)
    
    FileInfo.Path = Command()
    FileInfo.Name = Read2
    FileInfo.Changes = False
    FileInfo.Saved = True
    FileInfo.Import = False
    Randomize
    X = Int((500) * Rnd)
    TempFile = ProgramDir & "\File\th14" & Trim(Str(X)) & ".lda"
    FileCopy FileInfo.Path, TempFile
    LoadFile
    Read = oMisc.RecentFile(SaveNew, Trim(FileInfo.Path), Trim(FileInfo.Name))
    frmMain.LoadRecent
    frmMain.Caption = "GP2 Track Handler v1.4 [" & Trim(FileInfo.Name) & "]"
Exit Sub

ErrHandler:
    Print #Log, Date & " " & Time & " OpenCommandFile, Error Number: " & Err.Number & ", Error Description: " & Err.Description
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
End Sub

Public Sub LoadTimeData()
Dim LTime As String
Dim Driver As String
Dim Track As String
Dim TType As String
Dim ItemX
Dim Count5 As Long
    
    lstTime.ListItems.Clear
    Count5 = oDB.RecCount(dbFile)
    For X = 0 To Count5 - 1
        Read = oDB.GetRecord(dbFile, X)
        LTime = Mid(Read, 1, 8)
        Read2 = Mid(Read, 10, 10)
        TType = Mid(Read, 21, 4)
        Count1 = InStr(26, Read, ";")
        Driver = Mid(Read, 26, Count1 - 26)
        Count1 = Count1 + 1
        Count2 = InStr(Count1, Read, ";")
        Read2 = Mid(Read, Count1, Count2 - Count1)
        Track = Mid(Read, Count2 + 1)
        
        Set ItemX = lstTime.ListItems.Add(, "k" & X, X + 1)
        With ItemX
            .SubItems(1) = Track
            .SubItems(2) = Driver
            .SubItems(3) = LTime
            .SubItems(4) = TType
        End With
    Next
End Sub

Public Sub DeleteTime()
Dim ListNr
    ListNr = lstTime.SelectedItem.Index - 1
    Read = ProgramDir & "\tmpTime.lda"
    oDB.DeleteRecord dbFile, Read, lstTime.SelectedItem.Index - 1
    LoadTimeData
End Sub

Public Sub GetResText()
    For X = 0 To 11
        lblInfoText(X).Caption = LoadResString(X + 112)
    Next
End Sub

Public Function SetDir(Optional Change As Boolean) As String
    szTitle = "Select Track Directory"
    With tBrowseInfo
        .hwndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        SetDir = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
End Function

Public Sub AddExe()
    Read = oMisc.ReadINI("Misc", "ExePath", TempFile)
    If Read <> "" Then
        frmGP2Edit.Show vbModal, frmMain
    Else
        Read = ""
        Read = comExe
        If Read <> "" Then
            FileNum = FreeFile
            Open Read For Binary As FileNum
            Read2 = String(12, " ")
            Get #FileNum, 45445, Read2
            Close FileNum
            If Read2 = "Steven Young" Then
                oMisc.WriteINI "Misc", "ExePath", Read, TempFile
                frmGP2Edit.Show vbModal, frmMain
            Else
                MsgBox "This is not a valid GP2Edit Dos Pathch file.", vbInformation, TH
            End If
        Else
            Exit Sub
        End If
    End If
End Sub

Public Sub DropTime(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OldNr As Integer
    On Error GoTo ErrHandler
    If TreeNr <> 0 Then
        SaveTrackData TreeNr
    End If
    OldNr = TreeNr
    TreeNr = Mid(TreeView1.HitTest(X, Y).Key, 2, 2)
    TreeNr = TreeNr - 10

    GetTrackData TreeNr

    Count1 = Data.GetData(1) - 1
    Read = oDB.GetRecord(dbFile, Count1)
    Read2 = Mid(Read, 21, 4) 'Qual/race
    If Read2 = "Qual" Then
        txtQTime.Text = Mid(Read, 1, 8)  'time
        txtQDate.Text = Mid(Read, 10, 10) 'date
        Count1 = InStr(26, Read, ";")
        txtQDriver.Text = Mid(Read, 26, Count1 - 26)
        Count1 = Count1 + 1
        Count2 = InStr(Count1, Read, ";")
        txtQTeam.Text = Mid(Read, Count1, Count2 - Count1)
    Else
        txtRTime.Text = Mid(Read, 1, 8)  'time
        txtRDate.Text = Mid(Read, 10, 10) 'date
        Count1 = InStr(26, Read, ";")
        txtRDriver.Text = Mid(Read, 26, Count1 - 26)
        Count1 = Count1 + 1
        Count2 = InStr(Count1, Read, ";")
        txtRTeam.Text = Mid(Read, Count1, Count2 - Count1)
    End If
    SaveTrackData TreeNr
    If (TreeNr <> OldNr) And (OldNr <> 0) Then
        GetTrackData OldNr
    End If
    TreeNr = OldNr
    Set TreeView1.DropHighlight = Nothing
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 91
        MsgBox "You can't drop this here"
    Case Else
        Print #Log, Date & " " & Time & " DropTime, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
    Set TreeView1.DropHighlight = Nothing
End Sub

Public Sub DropTrack(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Path As String
Dim FileType As String

    On Error GoTo ErrHandler
    On Error Resume Next
    Read = ""
    Read = Data.Files(1)
    If Read <> "" Then
        FileType = LCase(Mid(Data.Files(1), Len(Data.Files(1)) - 3, 4))
        If FileType <> ".dat" Then
            Set TreeView1.DropHighlight = Nothing
            MsgBox "You can only add track files this way, if you want to add a menu pic you " & vbLf & "have to use the included file manager.", vbExclamation, TH
            Exit Sub
        End If
        Read = ReadGP2Info(Data.Files(1))
        If Read = False Then
            Set TreeView1.DropHighlight = Nothing
            Exit Sub
        End If
        Path = Data.Files(1)
    Else
        Path = lstFile.SelectedItem.Key
        FileType = lstFile.Tag
    End If
    On Error GoTo ErrHandler
    FileInfo.Changes = True

    If TreeView1.DropHighlight.Key <> "r" Then
        TreeNr = Mid(TreeView1.DropHighlight.Key, 2, 2)
        TreeNr = TreeNr - 10
    End If

Drop:
    If FileType = ".dat" Then
        TreeView1.Nodes(TreeView1.DropHighlight.Key).Selected = True
        TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-Track", "Track File: " & Path, 3, 3
        lstFile.Tag = ""
        Set TreeView1.DropHighlight = Nothing
        SaveDropData Path
    ElseIf FileType = "big" Then
        TreeView1.Nodes(TreeView1.DropHighlight.Key).Selected = True
        TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-BPic", "Big Pic: " & Path, 4, 4
        lstFile.Tag = ""
        GetTrackData TreeNr
        Set imgBPic.Picture = LoadPicture(Path)
        txtBPic.Text = Path
        SaveTrackData TreeNr
        Set TreeView1.DropHighlight = Nothing
    ElseIf FileType = "smal" Then
        TreeView1.Nodes(TreeView1.DropHighlight.Key).Selected = True
        TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-SPic", "Smal Pic: " & Path, 4, 4
        lstFile.Tag = ""
        GetTrackData TreeNr
        Set imgSPic.Picture = LoadPicture(Path)
        txtSPic.Text = Path
        SaveTrackData TreeNr
        Set TreeView1.DropHighlight = Nothing
    End If

Exit Sub
ErrHandler:
    Select Case Err.Number
    Case "35602"
        If FileType = ".dat" Then
            For X = 1 To TreeView1.Nodes.Count
                If TreeView1.Nodes(X).Key = TreeView1.SelectedItem.Key & "-Track" Then
                    TreeView1.Nodes(X).Parent.Text = "Track " & Mid(TreeView1.SelectedItem.Key, 2, 2) - 10
                    TreeView1.Nodes.Remove (X)
                    Exit For
                End If
            Next
        ElseIf FileType = "big" Then
            For X = 1 To TreeView1.Nodes.Count
                If TreeView1.Nodes(X).Key = TreeView1.SelectedItem.Key & "-BPic" Then
                    TreeView1.Nodes.Remove (X)
                    Exit For
                End If
            Next
        ElseIf FileType = "smal" Then
            For X = 1 To TreeView1.Nodes.Count
                If TreeView1.Nodes(X).Key = TreeView1.SelectedItem.Key & "-SPic" Then
                    TreeView1.Nodes.Remove (X)
                    Exit For
                End If
            Next
        End If
        GoTo Drop
    Case 91
        MsgBox LoadResString(104), vbInformation, TH
        Set TreeView1.DropHighlight = Nothing
    Case Else
        Print #Log, Date & " " & Time & " DropTrack, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub LoadFiles(ByVal Show As PatFile)
Dim WFD As WIN32_FIND_DATA
Dim hFile As Long
Dim Path As String
    If Len(Dir1.Path) = 3 Then
        Path = Dir1.Path
    Else
        Path = Dir1.Path & "\"
    End If
    lstFile.ListItems.Clear
    If (Show = Dat) Or (Show = All) Then
        hFile = FindFirstFile(Path & "*.dat", WFD)
        If hFile <> -1 Then
            lstFile.ListItems.Add , Path & Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
            While FindNextFile(hFile, WFD)
                lstFile.ListItems.Add , Path & Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
            Wend
        End If
    End If
    If (Show = Bmp) Or (Show = All) Then
        hFile = FindFirstFile(Path & "*.bmp", WFD)
        If hFile <> -1 Then
            lstFile.ListItems.Add , Path & Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
            While FindNextFile(hFile, WFD)
                lstFile.ListItems.Add , Path & Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
            Wend
        End If
    End If
    If (Show = Gif) Or (Show = All) Then
        hFile = FindFirstFile(Path & "*.gif", WFD)
        If hFile <> -1 Then
            lstFile.ListItems.Add , Path & Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
            While FindNextFile(hFile, WFD)
                lstFile.ListItems.Add , Path & Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1), Left(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1)
            Wend
        End If
    End If
    If lstFile.ListItems.Count > 0 Then
        lstFile.ListItems(1).Selected = True
    End If
End Sub

Public Sub ClearInfo()
    lblMisc = ""
    lblEvent = ""
    lblSlot = ""
    lblQual = ""
    lblRace = ""
    lblLen = ""
    lblWare = ""
    lblLaps = ""
    lblAuthor = ""
    lblInfoYear = ""
    lblTrackName = ""
    lblCountry = ""
End Sub

Public Function GetExt(File As String) As String
    GetExt = LCase(Mid(File, Len(File) - 3))
End Function
