VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{1DD137C0-F513-11D2-AFB9-C0B82D509E49}#1.0#0"; "JAD2JAM.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GP2 Track Handler v1.5"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9120
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   68
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgMisc"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   22
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Import"
            Object.ToolTipText     =   "Import data from GP2"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Export"
            Object.ToolTipText     =   "Export data to GP2"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GP2Edit"
            Object.ToolTipText     =   "Add GP2Edit file"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "JamCheck"
            Object.ToolTipText     =   "Jam Check"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Setup"
            Object.ToolTipText     =   "CC Car Settings"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Backup"
            Object.ToolTipText     =   "Backup a Track"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Home"
            Object.ToolTipText     =   "Goto default track directory"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Down"
            Object.ToolTipText     =   "Move track down one level"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Up"
            Object.ToolTipText     =   "Move track up one level"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GP2"
            Object.ToolTipText     =   "Run GP2"
            Object.Tag             =   ""
            ImageIndex      =   17
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Help"
            Object.ToolTipText     =   "Track Handler Help"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
      EndProperty
   End
   Begin JAD2JAMLib.Jad2Jam Jam2Jad 
      Left            =   1800
      Top             =   7080
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   0
   End
   Begin ComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   85
      Top             =   6705
      Width           =   9120
      _ExtentX        =   16087
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
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5115
            MinWidth        =   5115
            Text            =   "GP2 Version:"
            TextSave        =   "GP2 Version:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5998
            MinWidth        =   5998
            Text            =   "GP2 Directory:"
            TextSave        =   "GP2 Directory:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   6255
      Left            =   3195
      TabIndex        =   67
      Top             =   420
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "File Manager"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraNoSupport"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraMenuPic"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraFileInfo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frameNorm"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Track Data"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdJamCheck"
      Tab(1).Control(1)=   "fraRace"
      Tab(1).Control(2)=   "fraQual"
      Tab(1).Control(3)=   "frameInfo"
      Tab(1).Control(4)=   "frameData"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Menu Pics"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tabPic"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Misc Settings"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frameGlobal"
      Tab(3).Control(1)=   "framePlayer"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Lap Time Database"
      TabPicture(4)   =   "frmMain.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstTime"
      Tab(4).Control(1)=   "cmdAdd"
      Tab(4).Control(2)=   "cmdDelete"
      Tab(4).Control(3)=   "lblTextLen"
      Tab(4).Control(4)=   "lblTimeDrag"
      Tab(4).ControlCount=   5
      Begin ComctlLib.ListView lstTime 
         DragIcon        =   "frmMain.frx":0396
         Height          =   5175
         Left            =   -74880
         TabIndex        =   98
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
         Caption         =   "&Add Time"
         Height          =   315
         Left            =   -73680
         TabIndex        =   93
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   -74880
         TabIndex        =   92
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton cmdJamCheck 
         Caption         =   "&Jam Check"
         Height          =   375
         Left            =   -74880
         TabIndex        =   65
         Top             =   5640
         Width           =   1095
      End
      Begin TabDlg.SSTab tabPic 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   34
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
         TabCaption(1)   =   "Small Menu Picture"
         TabPicture(1)   =   "frmMain.frx":06BC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtSPic"
         Tab(1).Control(1)=   "imgSPic"
         Tab(1).ControlCount=   2
         Begin VB.TextBox txtSPic 
            Height          =   285
            Left            =   -74760
            TabIndex        =   87
            Top             =   3120
            Width           =   2055
            Visible         =   0   'False
         End
         Begin VB.TextBox txtBPic 
            Height          =   285
            Left            =   240
            TabIndex        =   86
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
         TabIndex        =   78
         Top             =   480
         Width           =   2775
         Begin VB.HScrollBar Slider1 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   2099
            Min             =   1900
            TabIndex        =   121
            Top             =   1800
            Value           =   1994
            Width           =   2415
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "Always save track record"
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   3120
            Width           =   2175
         End
         Begin VB.CheckBox chk0as1 
            Caption         =   "Show car 1 as 0"
            Height          =   255
            Left            =   240
            TabIndex        =   60
            Top             =   2760
            Width           =   2055
         End
         Begin VB.HScrollBar hscQRace 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   57
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
            TabIndex        =   54
            Top             =   600
            Value           =   1313
            Width           =   2415
         End
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   120
            TabIndex        =   79
            Top             =   3600
            Width           =   2415
         End
         Begin VB.Label lblYear 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1994"
            Height          =   195
            Left            =   1695
            TabIndex        =   59
            Top             =   1560
            Width           =   960
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Year for this Season"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   58
            Top             =   1560
            Width           =   1425
         End
         Begin VB.Label lblQuick 
            Alignment       =   1  'Right Justify
            Caption         =   "5%"
            Height          =   195
            Left            =   2175
            TabIndex        =   56
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Quick Race Length"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lblCWeight 
            Alignment       =   1  'Right Justify
            Caption         =   "1313lb (596kg)"
            Height          =   195
            Left            =   1455
            TabIndex        =   53
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Car Weight"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   52
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
            TabIndex        =   84
            Top             =   4800
            Width           =   285
         End
         Begin VB.Label lblRookie 
            AutoSize        =   -1  'True
            Caption         =   "Rookie"
            Height          =   180
            Index           =   3
            Left            =   1800
            TabIndex        =   83
            Top             =   3840
            Width           =   510
         End
         Begin VB.Label lblAmateur 
            AutoSize        =   -1  'True
            Caption         =   "Amateur"
            Height          =   180
            Left            =   1800
            TabIndex        =   82
            Top             =   4080
            Width           =   585
         End
         Begin VB.Label lblSemiPro 
            AutoSize        =   -1  'True
            Caption         =   "Semi-Pro"
            Height          =   180
            Left            =   1800
            TabIndex        =   81
            Top             =   4320
            Width           =   630
         End
         Begin VB.Label lblPro 
            AutoSize        =   -1  'True
            Caption         =   "Pro"
            Height          =   180
            Left            =   1800
            TabIndex        =   80
            Top             =   4560
            Width           =   240
         End
      End
      Begin VB.Frame framePlayer 
         Caption         =   "Player Car"
         Height          =   5175
         Left            =   -74880
         TabIndex        =   76
         Top             =   480
         Width           =   2775
         Begin VB.CommandButton cmdDefaultSettings 
            Caption         =   "GP2 Default"
            Height          =   315
            Left            =   1500
            TabIndex        =   64
            Top             =   4320
            Width           =   1035
         End
         Begin VB.CommandButton cmdImportSettings 
            Caption         =   "&Import"
            Height          =   315
            Left            =   120
            TabIndex        =   63
            Top             =   4700
            Width           =   1035
         End
         Begin VB.CommandButton cmdExportSettings 
            Caption         =   "&Export"
            Height          =   315
            Left            =   120
            TabIndex        =   62
            Top             =   4320
            Width           =   1035
         End
         Begin VB.CheckBox chkUPower 
            Caption         =   "Use Selected Team Power"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1560
            Width           =   2500
         End
         Begin VB.CheckBox chkNoLimit 
            Caption         =   "No Speed Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   3840
            Width           =   2500
         End
         Begin VB.HScrollBar hscPitSpeed 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   201
            Min             =   1
            TabIndex        =   50
            Top             =   3480
            Value           =   50
            Width           =   2415
         End
         Begin VB.HScrollBar hscPGrip 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   1000
            TabIndex        =   47
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
            TabIndex        =   44
            Top             =   2280
            Value           =   1313
            Width           =   2415
         End
         Begin VB.HScrollBar hscPQPower 
            Height          =   255
            LargeChange     =   20
            Left            =   120
            Max             =   1579
            TabIndex        =   40
            Top             =   1200
            Value           =   790
            Width           =   2415
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   120
            TabIndex        =   77
            Top             =   4200
            Width           =   2415
         End
         Begin VB.HScrollBar hscPRPower 
            Height          =   255
            LargeChange     =   20
            Left            =   120
            Max             =   1579
            TabIndex        =   37
            Top             =   600
            Value           =   780
            Width           =   2415
         End
         Begin VB.Label lblPit 
            AutoSize        =   -1  'True
            Caption         =   "Pit Speed Limit"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   3240
            Width           =   1050
         End
         Begin VB.Label lblPitSpeed 
            Alignment       =   1  'Right Justify
            Caption         =   "50mph (80km/h)"
            Height          =   195
            Left            =   1095
            TabIndex        =   49
            Top             =   3240
            Width           =   1440
         End
         Begin VB.Label lblGrip 
            Alignment       =   1  'Right Justify
            Caption         =   "198"
            Height          =   195
            Left            =   2055
            TabIndex        =   46
            Top             =   2640
            Width           =   480
         End
         Begin VB.Label lblGrip2 
            AutoSize        =   -1  'True
            Caption         =   "Grip"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   2640
            Width           =   285
         End
         Begin VB.Label lblWeight2 
            Alignment       =   1  'Right Justify
            Caption         =   "1313lb (596Kg)"
            Height          =   195
            Left            =   1335
            TabIndex        =   43
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label lblWeight 
            AutoSize        =   -1  'True
            Caption         =   "Car Weight"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label lblPRPower 
            Alignment       =   1  'Right Justify
            Caption         =   "780"
            Height          =   195
            Left            =   2055
            TabIndex        =   36
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Power in Qual"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   960
            Width           =   990
         End
         Begin VB.Label lblPQPower 
            Alignment       =   1  'Right Justify
            Caption         =   "790"
            Height          =   195
            Left            =   2055
            TabIndex        =   39
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Power in Race"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1050
         End
      End
      Begin VB.Frame fraRace 
         Caption         =   "Lap Time Data - Race"
         Height          =   2775
         Left            =   -71520
         TabIndex        =   74
         Top             =   3360
         Width           =   2175
         Begin VB.TextBox txtRTime 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   124
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CommandButton cmdSaveRace 
            Height          =   300
            Left            =   1800
            Picture         =   "frmMain.frx":06D8
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtRDate 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   33
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtRTeam 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   31
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtRDriver 
            Height          =   285
            Left            =   120
            MaxLength       =   23
            TabIndex        =   29
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Time (e.g. 1:24.145)"
            Height          =   195
            Left            =   120
            TabIndex        =   125
            Top             =   1440
            Width           =   1425
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Driver"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Team"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Date (e.g. 1999-06-19)"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   2040
            Width           =   1605
         End
      End
      Begin VB.Frame fraQual 
         Caption         =   "Lap Time Data - Qual"
         Height          =   2775
         Left            =   -71520
         TabIndex        =   73
         Top             =   480
         Width           =   2175
         Begin VB.TextBox txtQTime 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   122
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CommandButton cmdSaveQual 
            Height          =   300
            Left            =   1800
            Picture         =   "frmMain.frx":07DA
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtQDate 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   27
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtQTeam 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   25
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtQDriver 
            Height          =   285
            Left            =   120
            MaxLength       =   23
            TabIndex        =   23
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Time (e.g. 1:24.145)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   123
            Top             =   1440
            Width           =   1425
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Date (e.g 1999-06-19)"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   1560
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Team"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Driver"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.Frame frameInfo 
         Caption         =   "Track Info"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   71
         Top             =   480
         Width           =   3135
         Begin VB.TextBox txtAdjectiv 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox txtCountry 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   1680
            Width           =   2775
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   255
            TabIndex        =   91
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Track Path"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Track Name:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Country:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Adjective:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   2040
            Width           =   705
         End
      End
      Begin VB.Frame frameData 
         Caption         =   "Track Data"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   70
         Top             =   3360
         Width           =   3135
         Begin VB.VScrollBar vscLaps 
            Height          =   285
            Left            =   600
            Max             =   0
            Min             =   126
            TabIndex        =   17
            Top             =   480
            Value           =   3
            Width           =   200
         End
         Begin VB.TextBox txtTire 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   21
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtLength 
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   19
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtLaps 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   16
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tyre &Ware (14000-40000)"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   20
            Top             =   1440
            Width           =   1830
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tra&ck Length (0-9999 m)"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Laps (3-126)"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame frameNorm 
         Caption         =   " Select Track/Menu Picture "
         Height          =   3675
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   5655
         Begin ComctlLib.Toolbar Toolbar2 
            Height          =   390
            Left            =   3960
            TabIndex        =   101
            Top             =   195
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   688
            ButtonWidth     =   635
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            ImageList       =   "imgMisc"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   5
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Up"
                  Object.ToolTipText     =   "Up One Level"
                  Object.Tag             =   ""
                  ImageIndex      =   24
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   3
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "List"
                  Object.ToolTipText     =   "List"
                  Object.Tag             =   ""
                  ImageIndex      =   21
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Details"
                  Object.ToolTipText     =   "Details"
                  Object.Tag             =   ""
                  ImageIndex      =   22
                  Style           =   2
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Large"
                  Object.ToolTipText     =   "Large Icons"
                  Object.Tag             =   ""
                  ImageIndex      =   23
                  Style           =   2
               EndProperty
            EndProperty
         End
         Begin ComctlLib.ListView lstFile 
            Height          =   2750
            Left            =   120
            TabIndex        =   100
            Top             =   600
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   4842
            View            =   2
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            OLEDragMode     =   1
            _Version        =   327682
            Icons           =   "imgBig"
            SmallIcons      =   "imgSmall"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDragMode     =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Size"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Modified"
               Object.Width           =   1235
            EndProperty
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   99
            Top             =   240
            Width           =   3735
         End
         Begin VB.Label lblMyPath 
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   3360
            Width           =   5415
         End
         Begin VB.Label Label1 
            Height          =   135
            Left            =   2280
            TabIndex        =   75
            Top             =   4080
            Width           =   1335
         End
      End
      Begin VB.Frame fraFileInfo 
         Caption         =   "Track File Info"
         Height          =   2000
         Left            =   120
         TabIndex        =   88
         Top             =   4150
         Width           =   5655
         Begin VB.TextBox lblSlot 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   107
            ToolTipText     =   "Click to edit"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox lblMisc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   3000
            MultiLine       =   -1  'True
            TabIndex        =   106
            ToolTipText     =   "Click to Edit"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox lblTrackName 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            TabIndex        =   102
            ToolTipText     =   "Click to edit"
            Top             =   480
            Width           =   1700
         End
         Begin VB.TextBox lblRace 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   8
            TabIndex        =   5
            ToolTipText     =   "Click to edit"
            Top             =   1680
            Width           =   1700
         End
         Begin VB.TextBox lblQual 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   8
            TabIndex        =   4
            ToolTipText     =   "Click to edit"
            Top             =   1440
            Width           =   1700
         End
         Begin VB.TextBox lblWare 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   5
            TabIndex        =   3
            ToolTipText     =   "Click to edit"
            Top             =   1200
            Width           =   1700
         End
         Begin VB.TextBox lblLen 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   4
            TabIndex        =   2
            ToolTipText     =   "Click to edit"
            Top             =   960
            Width           =   1700
         End
         Begin VB.TextBox lblLaps 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   3
            TabIndex        =   1
            ToolTipText     =   "Click to edit"
            Top             =   720
            Width           =   1700
         End
         Begin VB.CommandButton cmdSaveGP2Info 
            Enabled         =   0   'False
            Height          =   345
            Left            =   5160
            Picture         =   "frmMain.frx":08DC
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Save GP2Info"
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox lblCountry 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            TabIndex        =   0
            ToolTipText     =   "Click to edit"
            Top             =   240
            Width           =   1700
         End
         Begin VB.TextBox lblEvent 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   3840
            TabIndex        =   6
            ToolTipText     =   "Click to edit"
            Top             =   735
            Width           =   1695
         End
         Begin VB.TextBox lblInfoYear 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   3840
            MaxLength       =   4
            TabIndex        =   103
            ToolTipText     =   "Click to edit"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblInfoText 
            AutoSize        =   -1  'True
            Caption         =   "&Slot:"
            Height          =   200
            Index           =   9
            Left            =   4440
            TabIndex        =   120
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Author:"
            Height          =   200
            Index           =   10
            Left            =   3000
            TabIndex        =   119
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Country:"
            Height          =   200
            Index           =   0
            Left            =   120
            TabIndex        =   118
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Laps:"
            Height          =   200
            Index           =   2
            Left            =   120
            TabIndex        =   117
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "L&ength (m):"
            Height          =   200
            Index           =   3
            Left            =   120
            TabIndex        =   116
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "Tyre &Ware:"
            Height          =   200
            Index           =   4
            Left            =   120
            TabIndex        =   115
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "Best Lap &Qual:"
            Height          =   200
            Index           =   5
            Left            =   120
            TabIndex        =   114
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "E&vent:"
            Height          =   200
            Index           =   8
            Left            =   3000
            TabIndex        =   113
            Top             =   735
            Width           =   735
         End
         Begin VB.Label lblInfoText 
            Caption         =   "Best Lap &Race:"
            Height          =   200
            Index           =   6
            Left            =   120
            TabIndex        =   112
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Track Name:"
            Height          =   200
            Index           =   1
            Left            =   120
            TabIndex        =   111
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Year:"
            Height          =   200
            Index           =   7
            Left            =   3000
            TabIndex        =   110
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblInfoText 
            AutoSize        =   -1  'True
            Caption         =   "Car &Setup:"
            Height          =   195
            Index           =   12
            Left            =   3000
            TabIndex        =   109
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Misc:"
            Height          =   200
            Left            =   3000
            TabIndex        =   108
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblAuthor 
            Height          =   200
            Left            =   3840
            TabIndex        =   7
            Top             =   960
            Width           =   1695
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblSetup 
            Height          =   200
            Left            =   3840
            TabIndex        =   104
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraMenuPic 
         Caption         =   "Menu Picture"
         Height          =   2000
         Left            =   120
         TabIndex        =   89
         Top             =   4150
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
            TabIndex        =   90
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame fraNoSupport 
         Height          =   2000
         Left            =   120
         TabIndex        =   126
         Top             =   4150
         Width           =   5655
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "File Not Supported!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   600
            TabIndex        =   127
            Top             =   720
            Width           =   4455
         End
      End
      Begin VB.Label lblTextLen 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   -70800
         TabIndex        =   97
         Top             =   6000
         Width           =   45
         Visible         =   0   'False
      End
      Begin VB.Label lblTimeDrag 
         Height          =   375
         Left            =   -71040
         TabIndex        =   94
         Top             =   5520
         Width           =   1695
      End
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   6247
      Left            =   0
      TabIndex        =   66
      Top             =   420
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   11007
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgMisc"
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin ComctlLib.ImageList imgBig 
      Left            =   600
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":09DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":109C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":132E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":15C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1852
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1AE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgSmall 
      Left            =   1200
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":22B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":23C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":24D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":25EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":26FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgMisc 
      Left            =   0
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   24
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":280E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2920
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2F8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":309E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":31B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":34CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":35DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":38F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3C10
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4260
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":45B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":48CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":49DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4D14
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
      Begin VB.Menu mnuImp 
         Caption         =   "&Import"
         Begin VB.Menu mnuImport 
            Caption         =   "Data &from GP2..."
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuImpRec 
            Caption         =   "&Lap Time Data from rec file..."
         End
      End
      Begin VB.Menu mnuExp 
         Caption         =   "&Export"
         Begin VB.Menu mnuExport 
            Caption         =   "&Data to GP2..."
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuExpRec 
            Caption         =   "&Lap Time Date to rec file..."
         End
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortOpen 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShortOpen 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShortOpen 
         Caption         =   ""
         Index           =   2
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
         Caption         =   "&GP2 Path..."
      End
      Begin VB.Menu mnuTrackPath 
         Caption         =   "&Default Track Path..."
      End
      Begin VB.Menu mnuTrackEditPath 
         Caption         =   "&Track Editor Path..."
      End
      Begin VB.Menu mnuSep31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunGP2 
         Caption         =   "Settings for ""Run GP2"" Button..."
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
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuList 
         Caption         =   "&List"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "&Details"
      End
      Begin VB.Menu mnuLargeIcons 
         Caption         =   "Lar&ge Icons"
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
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJamCheck2 
         Caption         =   "&Jam Check..."
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuTrackSettings 
         Caption         =   "&CC Car Settings..."
      End
      Begin VB.Menu mnuPCS 
         Caption         =   "Player Car Setup"
         Begin VB.Menu mnuSetupFile2 
            Caption         =   "Add Setup from File..."
         End
         Begin VB.Menu mnuNewSetup2 
            Caption         =   "Create New Setup..."
         End
         Begin VB.Menu mnuSep29 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditQSetup2 
            Caption         =   "Edit Qual Setup..."
         End
         Begin VB.Menu mnuEditRSetup2 
            Caption         =   "Edit Race Setup..."
         End
         Begin VB.Menu mnuRemove2 
            Caption         =   "Remove Setup"
         End
      End
      Begin VB.Menu mnuSep26 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup Track..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuSep22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGP2Edit 
         Caption         =   "A&dd GP2Edit File..."
      End
   End
   Begin VB.Menu mnuTopHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "On the &Web"
         Begin VB.Menu mnuTHHome 
            Caption         =   "GP2 Track Handler WebSite"
         End
         Begin VB.Menu mnuSep13 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGrandPrix1 
            Caption         =   "Unofficial GP1 WebSite"
         End
         Begin VB.Menu mnuGrandPrix2 
            Caption         =   "Unofficial GP2 WebSite"
         End
         Begin VB.Menu mnuUnGP3 
            Caption         =   "Unofficial GP3 WebSite"
         End
         Begin VB.Menu mnuSep30 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGP3 
            Caption         =   "Official GP3 WebSite"
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
      Begin VB.Menu mnuJamCheck 
         Caption         =   "&Jam Check..."
      End
      Begin VB.Menu mnuCCCarSetup 
         Caption         =   "&CC Car Settings..."
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "Player Car &Setup"
         Begin VB.Menu mnuSetupFile 
            Caption         =   "Add Setup from File..."
         End
         Begin VB.Menu mnuNewSetup 
            Caption         =   "Create New Setup..."
         End
         Begin VB.Menu mnuSep25 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditQSetup 
            Caption         =   "Edit Qual Setup..."
         End
         Begin VB.Menu mnuEditRSetup 
            Caption         =   "Edit Race Setup..."
         End
         Begin VB.Menu mnuRemove 
            Caption         =   "Remove Setup"
         End
      End
      Begin VB.Menu mnuSep24 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup2 
         Caption         =   "&Backup Track..."
      End
      Begin VB.Menu mnuCheckSum 
         Caption         =   "&Write Checksum"
      End
      Begin VB.Menu mnuSep27 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Edit Track"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep28 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename File"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete File"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function WinHelp Lib "user32.dll" Alias "WinHelpA" (ByVal hWndMain As Long, ByVal lpHelpFile As String, ByVal uCommand As Long, dwData As Any) As Long

Private MyPath As String
Private sFile As String
Private ChangeLap As Boolean
Private Dott As Boolean
Private RetVal
Private CheckBatFile As String
Private Support As Boolean
Private File As String
Private RunGP2 As String

'Left or Right mouse button
Private bButton As Byte

'WinHelp Const
Const HELP_CONTENTS = &H3

Const DRIVE_CDROM = 5
Const DRIVE_FIXED = 3
Const DRIVE_RAMDISK = 6
Const DRIVE_REMOTE = 4
Const DRIVE_REMOVABLE = 2

Private Enum PatFile
    Dat = 0
    Bmp = 1
    Gif = 2
    All = 3
End Enum

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
    Exp.GP2FileNum = FreeFile
    Open GP2Dir & "\gp2.exe" For Binary As Exp.GP2FileNum
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
    Close Exp.GP2FileNum
    Read = oMisc.File_Exists(GP2Dir & "\f1gstate.sav")
    If Read = True Then
        Exp.F1FileNum = FreeFile
        Open GP2Dir & "\f1gstate.sav" For Binary As Exp.F1FileNum
        ExportQuickRace
        Close Exp.F1FileNum
        Read = oMisc.GetShortName(GP2Dir & "\f1gstate.sav")
        RetVal = ShellExecute(frmMain.hwnd, "open", ProgramDir & "\gp2utils\thcheck.exe", Read, vbNullString, 0)
        RetVal = oMisc.CloseDosPrompt("thcheck")
    End If
    GetMisc
End Sub

Private Sub cmdImportSettings_Click()
    Exp.GP2FileNum = FreeFile
    Open GP2Dir & "\gp2.exe" For Binary As Exp.GP2FileNum
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
        Exp.F1FileNum = FreeFile
        Open GP2Dir & "\f1gstate.sav" For Binary As Exp.F1FileNum
        ImportQuick
        Close Exp.F1FileNum
    End If
    Close Exp.GP2FileNum
    GetMisc
End Sub

Private Sub cmdJamCheck_Click()
    JamCheck
End Sub

Public Sub cmdSaveGP2Info_Click()
    frmMain.MousePointer = 11
    MakeText
    Read = oMisc.GetShortName(lstFile.SelectedItem.Key)
    RetVal = ShellExecute(frmMain.hwnd, "open", ProgramDir & "\gp2utils\thcheck.exe", Read, vbNullString, 1)
    frmMain.MousePointer = 0
End Sub

Private Sub cmdSaveQual_Click()
    frmMain.MousePointer = 11
    If txtQTime.Text <> "" Then
        Read = txtQTime.Text & ";" & txtQDate.Text & ";Qual;" & txtQDriver.Text & ";" & txtQTeam.Text & ";" & txtName
        oDB.SaveNew dbFile, Read
        LoadTimeData
    Else
        MsgBox "You must have a time to save a time.", vbInformation, TH
    End If
    frmMain.MousePointer = 0
End Sub

Private Sub cmdSaveRace_Click()
    frmMain.MousePointer = 11
    Read = txtRTime.Text & ";" & txtRDate.Text & ";Race;" & txtRDriver.Text & ";" & txtRTeam.Text & ";" & txtName
    oDB.SaveNew dbFile, Read
    LoadTimeData
    frmMain.MousePointer = 0
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrHandler
    MyPath = Mid(Drive1.Drive, 1, 2)
    If Read2 <> "LoadTrackPathNow-Flag" Then ListFiles All
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 68
        MsgBox LoadResString(109), vbExclamation, TH
        Drive1.Drive = "C:"
    Case Else
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
            Resume Next
    End Select
End Sub

Private Sub Form_Load()
    frmMain.MousePointer = 11
    On Error Resume Next
    NewTree
    tabMain.TabEnabled(1) = False
    tabMain.TabEnabled(2) = False
    Toolbar1.Buttons("Up").Enabled = False
    Toolbar1.Buttons("Down").Enabled = False
    tabMain.Tab = 0
    frmMain.Show

    ProgramDir = App.Path
    If Right(ProgramDir, 1) = "\" Then ProgramDir = Mid(ProgramDir, 1, Len(ProgramDir) - 1)
    'ProgramDir = "G:\Mina Program\Visual Basic\GP2 Track Handler v15\TestCenter"

    oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "Path", ProgramDir

    FlatToolbar Toolbar1
    FlatToolbar Toolbar2

    Set oMisc = New Misc
    Set oData = New GP2Info
    Set oReg = New oReg
    Set oDB = New clsDB
    MkDir ProgramDir & "\File"
    MkDir ProgramDir & "\Bat"
    dbFile = ProgramDir & "\Time.lda"
    App.HelpFile = ProgramDir & "\Help.hlp"

    GetRegValue
    SetTextProp

    fraFileInfo.Enabled = False
    LoadTimeData

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

Private Sub lblCountry_GotFocus()
    lblCountry.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblCountry_LostFocus()
    lblCountry.BackColor = "&H8000000F"
End Sub

Private Sub lblEvent_GotFocus()
    lblEvent.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblEvent_LostFocus()
    lblEvent.BackColor = "&H8000000F"
End Sub

Private Sub lblInfoYear_GotFocus()
    lblInfoYear.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblInfoYear_LostFocus()
    lblInfoYear.BackColor = "&H8000000F"
End Sub

Private Sub lblLaps_GotFocus()
    lblLaps.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblLaps_LostFocus()
    lblLaps.BackColor = "&H8000000F"
End Sub

Private Sub lblLen_GotFocus()
    lblLen.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblLen_LostFocus()
    lblLen.BackColor = "&H8000000F"
End Sub

Private Sub lblMisc_LostFocus()
    lblMisc.BackColor = "&H8000000F"
End Sub

Private Sub lblQual_GotFocus()
    lblQual.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblMisc_GotFocus()
    lblMisc.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblQual_LostFocus()
    lblQual.BackColor = "&H8000000F"
End Sub

Private Sub lblRace_GotFocus()
    lblRace.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblRace_LostFocus()
    lblRace.BackColor = "&H8000000F"
End Sub

Private Sub lblSlot_GotFocus()
    lblSlot.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblSlot_LostFocus()
    lblSlot.BackColor = "&H8000000F"
End Sub

Private Sub lblTrackName_GotFocus()
    lblTrackName.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblTrackName_LostFocus()
    lblTrackName.BackColor = "&H8000000F"
End Sub

Private Sub lblWare_GotFocus()
    lblWare.BackColor = "&H80000005"
    TextSelected
End Sub

Private Sub lblWare_LostFocus()
    lblWare.BackColor = "&H8000000F"
End Sub

Private Sub lblYear_Click()
    Read = InputBox("Year for this Season", "Select Year", lblYear.Caption)
    If Read <> "" Then Slider1.Value = Read
End Sub

Private Sub lstFile_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim TempInt As Integer
    Read = MyPath
    If Right(Read, 1) <> "\" Then Read = Read & "\"
    Read = Read & NewString
    If LCase(lstFile.SelectedItem.Key) <> LCase(Read) Then
        FileCopy lstFile.SelectedItem.Key, Read
        Kill lstFile.SelectedItem.Key
        TempInt = lstFile.SelectedItem.Index
        ListFiles All
        lstFile.ListItems(TempInt).Selected = True
    End If
End Sub

Private Sub lstFile_Click()
Dim PicY As Long
Dim PicX As Long
    On Error GoTo ErrHandler
    If LCase(Mid(lstFile.SelectedItem.Key, 1, 3)) = "dir" Or LCase(Mid(lstFile.SelectedItem.Key, 1, 5)) = "drive" Then
        ClearInfo
        Exit Sub
    End If
    If lstFile.SelectedItem.Text <> "" Then
        Read = lstFile.SelectedItem.Key
        If GetExt(Read) = ".dat" Then
            lstFile.Tag = ".dat"
            ClearInfo
            Support = ReadGP2Info(Read)
            If Support = True Then
                FileNum = FreeFile
                Open lstFile.SelectedItem.Key For Binary As FileNum
                Get #FileNum, 3997, Var.iInt1
                If Var.iInt1 = 12345 Then
                    Get #FileNum, 3999, Var.iInt1
                    If Var.iInt1 = 1 Then
                        lblSetup.Caption = "Yes"
                    Else
                        lblSetup.Caption = "No"
                    End If
                Else
                    lblSetup.Caption = "No"
                End If
                Close FileNum
                fraFileInfo.Enabled = True
                cmdSaveGP2Info.Enabled = True
                fraMenuPic.Visible = False
                fraFileInfo.Visible = True
                fraNoSupport.Visible = False
            Else
                fraMenuPic.Visible = False
                fraFileInfo.Visible = False
                fraNoSupport.Visible = True
                fraFileInfo.Enabled = False
                cmdSaveGP2Info.Enabled = False
            End If
        ElseIf (GetExt(Read) = ".bmp") Or (GetExt(Read) = ".gif") Then
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
                lblPicInfo.Caption = "Small Menu Picture"
                Set imgPre.Picture = LoadPicture(Read)
                lstFile.Tag = "small"
                Support = True
            Else
                Support = False
                lstFile.OLEDragMode = ccOLEDragAutomatic
                lstFile.OLEDragMode = ccOLEDragManual
                fraNoSupport.Visible = True
                fraMenuPic.Visible = False
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
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub lstFile_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    If ColumnHeader.Text = "Size" Then
        lstFile.SortKey = 1
        lstFile.Sorted = True
        If lstFile.SortOrder = lvwAscending Then
            lstFile.SortOrder = lvwDescending
        Else
            lstFile.SortOrder = lvwAscending
        End If
    ElseIf ColumnHeader.Text = "Name" Then
        lstFile.SortKey = 0
        lstFile.Sorted = True
        If lstFile.SortOrder = lvwAscending Then
            lstFile.SortOrder = lvwDescending
        Else
            lstFile.SortOrder = lvwAscending
        End If
    ElseIf ColumnHeader.Text = "Modified" Then
        lstFile.SortKey = 2
        lstFile.Sorted = True
        If lstFile.SortOrder = lvwAscending Then
            lstFile.SortOrder = lvwDescending
        Else
            lstFile.SortOrder = lvwAscending
        End If
    End If
End Sub

Private Sub lstFile_DblClick()
    If bButton = 2 Then Exit Sub
    On Error GoTo ErrHandler
    If Mid(lstFile.SelectedItem.Key, 1, 3) = "dir" Then
        MyPath = Mid(lstFile.SelectedItem.Key, 4)
        ListFiles All
    End If
    If LCase(Mid(lstFile.SelectedItem.Key, 1, 5)) = "drive" Then
        Toolbar2.Buttons(1).Enabled = True
        Drive1.Drive = Mid(lstFile.SelectedItem.Key, 6)
        ListFiles All
        Exit Sub
    End If
ErrHandler:
End Sub

Private Sub lstFile_ItemClick(ByVal Item As ComctlLib.ListItem)
    If File <> lstFile.SelectedItem.Text Then
        lstFile_Click
        File = lstFile.SelectedItem.Text
    End If
End Sub

Private Sub lstFile_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 40) Or (KeyCode = 38) Then lstFile_Click
    If KeyCode = 13 Then lstFile_DblClick
    If KeyCode = 8 Then
        UpOneLevel
    End If
End Sub

Private Sub lstFile_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    bButton = Button
End Sub

Private Sub lstFile_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then
        If LCase(Mid(lstFile.SelectedItem.Key, 1, 3)) <> "dir" And LCase(Right(lstFile.SelectedItem.Key, 3)) = "dat" Then
            If LCase(lblSetup.Caption) = "no" Then
                mnuEditQSetup.Enabled = False
                mnuEditRSetup.Enabled = False
                mnuRemove.Enabled = False
            Else
                mnuEditQSetup.Enabled = True
                mnuEditRSetup.Enabled = True
                mnuRemove.Enabled = True
            End If
            PopupMenu mnuPopup
        End If
    End If
End Sub

Private Sub lstFile_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
    If Support = True Then
        Data.SetData lstFile.SelectedItem.Text, 1
    End If
End Sub

Private Sub mnuBackup_Click()
    If tabMain.Tab = 0 Then
        If Mid(lstFile.SelectedItem.Key, Len(lstFile.SelectedItem.Key) - 3, 4) <> ".dat" Then Exit Sub
        Var.sString1 = Mid(lstFile.SelectedItem.Text, 1, Len(lstFile.SelectedItem.Text) - 3) & "zip"
        Read = oMisc.ShowSave("Zip Files (*.Zip)|*.zip|", "zip", Me.hwnd, ProgramDir, "Create Zip File", Var.sString1)
        If Read <> "" Then
            Me.MousePointer = 11
            DoEvents
            BackupTrack lstFile.SelectedItem.Key, Read
            Me.MousePointer = 0
        End If
    ElseIf tabMain.Tab = 1 Then
        If txtPath.Text = "" Then Exit Sub
        Var.sString1 = GetFileName(txtPath.Text)
        Var.sString1 = Mid(Var.sString1, 1, Len(Var.sString1) - 3) & "zip"
        Read = oMisc.ShowSave("Zip Files (*.Zip)|*.Zip|", "zip", Me.hwnd, ProgramDir, "Create Zip File", Var.sString1)
        If Read <> "" Then
            Me.MousePointer = 11
            DoEvents
            BackupTrack txtPath.Text, Read
            Me.MousePointer = 0
        End If
    End If
End Sub

Private Sub mnuBackup2_Click()
    mnuBackup_Click
End Sub

Private Sub mnuCheckSum_Click()
    WriteCheckSum lstFile.SelectedItem.Key
End Sub

Private Sub mnuDelete_Click()
Dim TempInt As Integer
    Var.iInt1 = MsgBox("Are you sure you want to delete this file?", vbYesNo, "Confirm File Delete")
    If Var.iInt1 = vbYes Then
        Kill lstFile.SelectedItem.Key
        TempInt = lstFile.SelectedItem.Index - 1
        ListFiles All
        lstFile.ListItems(TempInt).Selected = True
    End If
End Sub

Private Sub mnuDetails_Click()
    mnuList.Checked = False
    mnuDetails.Checked = True
    mnuLargeIcons.Checked = False
    lstFile.View = lvwReport
    Toolbar2.Buttons(4).Value = tbrPressed
    oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "ToolBar", , 4
End Sub

Private Sub mnuEditQSetup_Click()
    FileNum = FreeFile
    Open lstFile.SelectedItem.Key For Binary As FileNum
    Read = String(80, " ")
    Get #FileNum, 3969, Read
    Close FileNum
    FileNum = FreeFile
    Open ProgramDir & "\file\impQSet.tmp" For Binary As FileNum
    Put #FileNum, 1, Read
    Close FileNum
    frmMain.Tag = " - Qual Setup"
    Load frmSetup
    OpenSetup ProgramDir & "\file\impQSet.tmp"
    frmSetup.Show vbModal, frmMain
    On Error Resume Next
    Kill ProgramDir & "\file\impQSet.tmp"
    frmMain.Tag = ""
End Sub

Private Sub mnuEditQSetup2_Click()
    mnuEditQSetup_Click
End Sub

Private Sub mnuEditRSetup_Click()
    FileNum = FreeFile
    Open lstFile.SelectedItem.Key For Binary As FileNum
    Read = String(80, " ")
    Get #FileNum, 4017, Read
    Close FileNum
    FileNum = FreeFile
    Open ProgramDir & "\file\impRSet.tmp" For Binary As FileNum
    Put #FileNum, 1, Read
    Close FileNum
    frmMain.Tag = " - Race Setup"
    Load frmSetup
    OpenSetup ProgramDir & "\file\impRSet.tmp"
    frmSetup.Show vbModal, frmMain
    On Error Resume Next
    Kill ProgramDir & "\file\impRSet.tmp"
    frmMain.Tag = ""
End Sub

Private Sub mnuEditRSetup2_Click()
    mnuEditRSetup_Click
End Sub

Private Sub mnuExpRec_Click()
Dim sFileName As String
    sFileName = oMisc.ShowSave("Lap Time Data File (*.rec)|*.rec|All files (*.*)|*.*|", "rec", Me.hwnd, GP2Dir, "Save Record File")
    If sFileName = "" Then Exit Sub
    FileNum = FreeFile
    Open GP2Dir & "\f1gstate.sav" For Binary As FileNum
    Read2 = String(2820, " ")
    Get #FileNum, 650, Read2
    Close FileNum
    Exp.F1FileNum = FreeFile
    Open sFileName For Binary As Exp.F1FileNum
    Read4 = String(27, Chr(0))
    Read3 = Chr(190) & Chr(161) & Chr(61) & Chr(133) & Chr(1) & Read4
    Put #Exp.F1FileNum, 1, Read3
    Read3 = ""
    Read4 = ""
    Put #Exp.F1FileNum, 33, Read2
    For Exp.TrackNr = 0 To 15
        ImportQDate RecFile
        ImportRDate RecFile
        ExportQName RecFile
        ExportRName RecFile
        ExportQTeam RecFile
        ExportRTeam RecFile
        ExportTime Qual, RecFile
        ExportTime Race, RecFile
    Next
    Close Exp.F1FileNum
    WriteCheckSum sFileName
End Sub

Private Sub mnuGP2Edit_Click()
    AddExe
End Sub

Private Sub mnuImpRec_Click()
    Read = oMisc.ShowOpen("Lap Time Data File (*.rec)|*.rec|All files (*.*)|*.*|", Me.hwnd, GP2Dir, "Select a Record File")
    If Read = "" Then Exit Sub
    Exp.F1FileNum = FreeFile
    Open Read For Binary As Exp.F1FileNum
    For Exp.TrackNr = 0 To 15
        ImportQDate RecFile
        ImportRDate RecFile
        ImportQName RecFile
        ImportRName RecFile
        ImportQTeam RecFile
        ImportRTeam RecFile
        ImportTime Qual, RecFile
        ImportTime Race, RecFile
    Next
    Close Exp.F1FileNum
End Sub

Private Sub mnuJamCheck_Click()
    JamCheck
End Sub

Private Sub mnuJamCheck2_Click()
    JamCheck
End Sub

Private Sub mnuLargeIcons_Click()
    mnuList.Checked = False
    mnuDetails.Checked = False
    mnuLargeIcons.Checked = True
    lstFile.View = lvwIcon
    Toolbar2.Buttons(5).Value = tbrPressed
    oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "ToolBar", , 5
End Sub

Private Sub mnuList_Click()
    mnuList.Checked = True
    mnuDetails.Checked = False
    mnuLargeIcons.Checked = False
    lstFile.View = lvwList
    Toolbar2.Buttons(3).Value = tbrPressed
    oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "ToolBar", , 3
End Sub

Private Sub mnuNewSetup_Click()
    frmSetup.Show vbModal, frmMain
End Sub

Private Sub mnuNewSetup2_Click()
    mnuNewSetup_Click
End Sub

Private Sub mnuOpenFile_Click()
    If mnuOpenFile.Tag = "" Then
        RetVal = ShellExecute(Me.hwnd, "open", lstFile.SelectedItem.Key, vbNullString, vbNullString, 1)
    Else
        Read = oMisc.GetShortName(lstFile.SelectedItem.Key)
        Read2 = ""
        For X = Len(mnuOpenFile.Tag) To 1 Step -1
            If Mid(mnuOpenFile.Tag, X, 1) = "\" Then Exit For
        Next
        Read2 = Mid(mnuOpenFile.Tag, 1, X - 1)
        RetVal = ShellExecute(Me.hwnd, "open", mnuOpenFile.Tag, Read, Read2, 1)
    End If
End Sub

Private Sub mnuRemove_Click()
    DeteteSetup lstFile.SelectedItem.Key
    RetVal = ShellExecute(Me.hwnd, "open", ProgramDir & "\gp2utils\thcheck.exe", oMisc.GetShortName(frmMain.lstFile.SelectedItem.Key), vbNullString, 1)
End Sub

Private Sub mnuRemove2_Click()
    mnuRemove_Click
End Sub

Private Sub mnuRename_Click()
    lstFile.StartLabelEdit
End Sub

Private Sub mnuRunGP2_Click()
    Read = oMisc.ShowOpen("All Application Files (*.exe)|*.exe|GP2 (gp2.exe)|gp2.exe|GP2Lap (gp2lap.exe)|gp2lap.exe|", Me.hwnd, GP2Dir, "Select Application")
    If Read = "" Then Exit Sub
    RunGP2 = Read
    oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "RunGP2", RunGP2
End Sub

Private Sub mnuSetupFile_Click()
    Var.sString1 = oMisc.ShowOpen("GP2 Setup File (*.cs*)|*.cs*|All Files (*.*)|*.*|", Me.hwnd, GP2Dir, "Open CarSetup")
    If Var.sString1 = "" Then Exit Sub
    Load frmSetup
    OpenSetup Var.sString1
    frmSetup.Show vbModal, frmMain
End Sub

Private Sub mnuSetupFile2_Click()
    mnuSetupFile_Click
End Sub

Private Sub mnuShortOpen_Click(Index As Integer)
    On Error GoTo ErrHandler

    MakeTempFile mnuShortOpen(Index).Tag

    FileInfo.Name = mnuShortOpen(Index).Caption
    FileInfo.Path = mnuShortOpen(Index).Tag
    FileInfo.Saved = True
    FileInfo.Import = False
    LoadFile
    Read = oMisc.RecentFile(OpenRecent, , , Index + 1)
    LoadRecent
    frmMain.Caption = "GP2 Track Handler v1.5 [" & Trim(FileInfo.Name) & "]"
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 53
        MsgBox LoadResString(111), vbExclamation, TH
    Case Else
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub mnuAbout_Click()
    frmMain.MousePointer = 11
    frmAbout.Show vbModal, frmMain
    frmMain.MousePointer = 0
End Sub

Private Sub mnuCCCarSetup_Click()
    On Error Resume Next
    If tabMain.Tab = 0 Then
        Read = GetExt(lstFile.SelectedItem.Key)
        If Read <> ".dat" Then Exit Sub
    Else
        If txtPath.Text = "" Then Exit Sub
    End If
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
    End
End Sub

Private Sub mnuExport_Click()
    SaveTrackData TreeNr
    SaveMisc
    frmExport.Show vbModal, frmMain
End Sub

Private Sub mnuGP2Path_Click()
    GP2Dir = SetGP2Folder
    If GP2Dir <> "" Then GetGP2Version
End Sub

Private Sub mnuGP3_Click()
    oMisc.INetLink "http://www.f1-grandprix3.com/", Me.hwnd
End Sub

Private Sub mnuGrandPrix1_Click()
    oMisc.INetLink "http://www.grandprix1.com/", Me.hwnd
End Sub

Private Sub mnuGrandPrix2_Click()
    oMisc.INetLink "http://www.grandprix2.com/", Me.hwnd
End Sub

Private Sub mnuHelp_Click()
    RetVal = WinHelp(frmMain.hwnd, App.HelpFile, HELP_CONTENTS, ByVal 0)
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
    frmMain.Caption = "GP2 Track Handler v1.5"
    Unload frmExport
    Unload frmImport
End Sub

Private Sub mnuOpen_Click()
    Var.sString1 = oMisc.ShowOpen("Track Handler Files (*.ths)|*.ths|All Files (*.*)|*.*|", Me.hwnd, ProgramDir)
    If Var.sString1 = "" Then Exit Sub
    MakeTempFile Var.sString1
    FileInfo.Saved = True
    FileInfo.Path = Var.sString1
    For X = Len(Var.sString1) To 0 Step -1
        If Mid(Var.sString1, X, 1) = "\" Then Exit For
    Next
    FileInfo.Name = Mid(Var.sString1, X + 1)
    LoadFile
    GetMisc

    Read = oMisc.RecentFile(SaveNew, FileInfo.Path, FileInfo.Name)
    frmMain.LoadRecent
    frmMain.Caption = "GP2 Track Handler v1.5 [" & FileInfo.Name & "]"
End Sub

Private Sub mnuPoint_Click()
    frmPoint.Show vbModal, frmMain
End Sub

Private Sub mnuRand_Click()
    frmMain.MousePointer = 11
    ListFiles Dat
    If lstFile.ListItems.Count > 15 Then
        TreeView1_NodeClick TreeView1.Nodes(1)
        TreeView1.Nodes(1).Selected = True
        RandomTracks lstFile.ListItems.Count
        FileInfo.Changes = True
    Else
        MsgBox "You need to have 16 track's or more in this directory to make a random season.", vbInformation, TH
    End If
    ListFiles All
    frmMain.MousePointer = 0
End Sub

Private Sub mnuReset_Click()
    SaveTrackData TreeNr
    frmReset.Show vbModal, frmMain
    GetTrackData TreeNr
End Sub

Private Sub mnuSave_Click()
    SaveTrackData TreeNr
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
        frmMain.Height = frmMain.Height - stbMain.Height
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "Statusbar", , "1"
    Else
        stbMain.Visible = True
        mnuStatusbar.Checked = True
        frmMain.Height = frmMain.Height + stbMain.Height
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "Statusbar", , "0"
    End If
End Sub

Private Sub mnuTHHome_Click()
    oMisc.INetLink "http://hem1.passagen.se/formula1/", Me.hwnd
End Sub

Private Sub mnuToolbar_Click()
    If mnuToolbar.Checked = True Then
        Toolbar1.Visible = False
        mnuToolbar.Checked = False
        TreeView1.Top = TreeView1.Top - Toolbar1.Height
        tabMain.Top = tabMain.Top - Toolbar1.Height
        frmMain.Height = frmMain.Height - Toolbar1.Height
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "Toolbar", , 1
    Else
        Toolbar1.Visible = True
        mnuToolbar.Checked = True
        TreeView1.Top = TreeView1.Top + Toolbar1.Height
        tabMain.Top = tabMain.Top + Toolbar1.Height
        frmMain.Height = frmMain.Height + Toolbar1.Height
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "Toolbar", , 0
    End If
End Sub

Private Sub mnuTrackEditPath_Click()
    Read = oMisc.ShowOpen("GP2 Track Editor (*.exe)|*.exe|", Me.hwnd, "", "Select GP2 Track Editor")
    If Read = "" Then Exit Sub
    oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "TrackEdit", Read
    mnuOpenFile.Tag = Read
    mnuOpenFile.Enabled = True
End Sub

Private Sub mnuTrackPath_Click()
    Read = oMisc.BrowseFolders("Select Track Directory", Me.hwnd)
    If Read <> "" Then
        If Right(Read, 1) <> "\" Then Read = Read & "\"
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "TrackPath", Read
        Read2 = "LoadTrackPathNow-Flag"
        Drive1.Drive = Mid(Read, 1, 2)
        Read2 = ""
        MyPath = Read
        ListFiles All
    End If
End Sub

Private Sub mnuTrackSettings_Click()
    On Error Resume Next
    frmCCSetup.Show vbModal, frmMain
End Sub

Private Sub mnuUnGP3_Click()
    oMisc.INetLink "http://www.gp3.org/", Me.hwnd
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
        Toolbar1.Buttons("Home").Enabled = True
        mnuTrackSettings.Enabled = True
        mnuBackup.Enabled = True
        mnuJamCheck2.Enabled = True
        Toolbar1.Buttons("Backup").Enabled = True
        Toolbar1.Buttons("Setup").Enabled = True
        Toolbar1.Buttons("JamCheck").Enabled = True
    ElseIf tabMain.Tab = 1 Then
        Toolbar1.Buttons("Home").Enabled = False
        mnuRand.Enabled = False
        mnuTrackSettings.Enabled = True
        mnuBackup.Enabled = True
        mnuJamCheck2.Enabled = True
        Toolbar1.Buttons("Backup").Enabled = True
        Toolbar1.Buttons("Setup").Enabled = True
        Toolbar1.Buttons("JamCheck").Enabled = True
    Else
        mnuRand.Enabled = False
        Toolbar1.Buttons("Home").Enabled = False
        mnuTrackSettings.Enabled = False
        mnuBackup.Enabled = False
        mnuJamCheck2.Enabled = False
        Toolbar1.Buttons("Backup").Enabled = False
        Toolbar1.Buttons("Setup").Enabled = False
        Toolbar1.Buttons("JamCheck").Enabled = False
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
        Read = ""
        For X = Len(RunGP2) To 0 Step -1
            If Mid(RunGP2, X, 1) = "\" Then Exit For
        Next
        Read = Mid(RunGP2, 1, X)
        RetVal = ShellExecute(frmMain.hwnd, "open", RunGP2, vbNullString, Read, 1)
    Case "GP2Edit"
        AddExe
    Case "Help"
        RetVal = WinHelp(frmMain.hwnd, App.HelpFile, HELP_CONTENTS, ByVal 0)
    Case "Backup"
        mnuBackup_Click
    Case "JamCheck"
        JamCheck
    Case "Setup"
        mnuCCCarSetup_Click
    Case "Home"
        Read = ""
        Read = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "TrackPath")
        If Read <> "" Then Read2 = oMisc.File_Exists(Read)
        If (Read <> "") And (LCase(Read2) = "true") Then
            Drive1.Drive = Mid(Read, 1, 2)
            MyPath = Read
            ListFiles All
        End If
    Case "Down"
        MoveTrackDown
    Case "Up"
        MoveTrackUp
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

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    Case "Up"
        UpOneLevel
    Case "List"
        lstFile.View = lvwList
        mnuList.Checked = True
        mnuDetails.Checked = False
        mnuLargeIcons.Checked = False
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "ToolBar", , 3
    Case "Details"
        lstFile.View = lvwReport
        mnuList.Checked = False
        mnuDetails.Checked = True
        mnuLargeIcons.Checked = False
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "ToolBar", , 4
    Case "Large"
        lstFile.View = lvwIcon
        mnuList.Checked = False
        mnuDetails.Checked = False
        mnuLargeIcons.Checked = True
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\GP2 Track Handler\Settings", "ToolBar", , 5
    End Select
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim TempIndex As Integer
Dim KeyName As String
    On Error GoTo ErrHandler
    Read = ""
    If KeyCode = 46 And TreeView1.SelectedItem.Key <> "r" Then
        X = InStr(1, TreeView1.SelectedItem.Key, "-")
        If X <> 0 Then KeyName = LCase(Mid(TreeView1.SelectedItem.Key, X + 1))
        If (KeyName = "track") Or (KeyName = "") Then
            If KeyName = "track" Then TreeView1.Nodes(TreeView1.SelectedItem.Parent.Index).Selected = True
            X = TreeView1.SelectedItem.Children
            Do Until TreeView1.SelectedItem.Children = 0
                TreeView1.Nodes.Remove (TreeView1.SelectedItem.Child.Index)
            Loop
            TreeView1.SelectedItem.Text = "Track " & Mid(TreeView1.SelectedItem.Key, 2, 2) - 10
            Tracks(TreeNr - 1) = False
            ClearText
            oMisc.WriteINI "Track " & TreeNr, "BPic", "", TempFile
            oMisc.WriteINI "Track " & TreeNr, "SPic", "", TempFile
            Set imgBPic = Nothing
            Set imgSPic = Nothing
            txtBPic.Text = ""
            txtSPic.Text = ""
            SaveTrackData TreeNr
            TreeView1.SelectedItem.Root.Selected = True
            TreeView1_NodeClick TreeView1.SelectedItem
        Else
            TempIndex = TreeView1.SelectedItem.Index
            TreeView1.Nodes(TreeView1.SelectedItem.Previous.Index).Selected = True
            TreeView1_NodeClick TreeView1.SelectedItem
            TreeView1.Nodes.Remove (TempIndex)
            If KeyName = "bpic" Then
                oMisc.WriteINI "Track " & TreeNr, "BPic", "", TempFile
                txtBPic.Text = ""
                Set imgBPic = Nothing
            ElseIf KeyName = "spic" Then
                oMisc.WriteINI "Track " & TreeNr, "SPic", "", TempFile
                txtSPic.Text = ""
                Set imgSPic = Nothing
            End If
        End If
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
        Toolbar1.Buttons("Down").Enabled = False
        Toolbar1.Buttons("Up").Enabled = False
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
        If TreeNr = 16 Then
            Toolbar1.Buttons("Down").Enabled = False
            Toolbar1.Buttons("Up").Enabled = True
        ElseIf TreeNr = 1 Then
            Toolbar1.Buttons("Up").Enabled = False
            Toolbar1.Buttons("Down").Enabled = True
        Else
            Toolbar1.Buttons("Up").Enabled = True
            Toolbar1.Buttons("Down").Enabled = True
        End If
    End If
End Sub

Private Sub TreeView1_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
    If tabMain.Tab = 4 Then
        DropTime Data, Effect, Button, Shift, X, y
    ElseIf tabMain.Tab = 0 Then
        DropTrack Data, Effect, Button, Shift, X, y
    End If
End Sub

Private Sub TreeView1_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single, State As Integer)
    Set TreeView1.DropHighlight = TreeView1.HitTest(X, y)
End Sub

Private Sub txtAdjectiv_GotFocus()
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
        mnuShortOpen(0).Visible = True
        mnuShortOpen(0).Caption = Name1
        mnuShortOpen(0).Tag = Path1
        mnuSep3.Visible = True
    End If
    If Name2 <> "" Then
        mnuShortOpen(1).Visible = True
        mnuShortOpen(1).Caption = Name2
        mnuShortOpen(1).Tag = Path2
    End If
    If Name3 <> "" Then
        mnuShortOpen(2).Visible = True
        mnuShortOpen(2).Caption = Name3
        mnuShortOpen(2).Tag = Path3
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
        Var.iInt1 = MsgBox(LoadResString(106), vbYesNoCancel, TH)
        If Var.iInt1 = vbNo Then
            CheckIfSave = ""
        ElseIf Var.iInt1 = vbCancel Then
            CheckIfSave = "Cancel"
        Else
            CheckIfSave = ""
            mnuSave_Click
        End If
        Exit Function
    End If
    If FileInfo.Changes = True Then
        Var.iInt1 = MsgBox(LoadResString(106), vbYesNoCancel, TH)
        If Var.iInt1 = vbNo Then
            CheckIfSave = ""
        ElseIf Var.iInt1 = vbCancel Then
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
Dim Misc As String
Dim Created As String
    FileNum = FreeFile
    SetAttr lstFile.SelectedItem.Key, vbNormal
    Open lstFile.SelectedItem.Key For Binary As FileNum

    Read = ""
    Read2 = ""
    Read = String(4000, " ")
    Get #FileNum, 1, Read
    'Get the Created text (Paul Hoad or Iso)
    X = InStr(1, Read, "|Created|")
    If X > 0 Then
        Var.iInt1 = X + 9
        X = InStr(X + 9, Read, "|")
        Created = Mid(Read, Var.iInt1, X - Var.iInt1)
    End If

    Read = "#GP2INFO|Name|" & lblTrackName & "|Country|" & lblCountry & "|Created|" & Created & "|Author|" & lblAuthor & _
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
    frmMain.Caption = "GP2 Track Handler v1.5 [" & Trim(FileInfo.Name) & "]"
Exit Sub

ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
End Sub

Public Sub LoadTimeData()
Dim Data(0 To 5) As String
Dim ItemX
    lstTime.ListItems.Clear
    Var.iInt1 = oDB.RecCount(dbFile)
    For X = 0 To Var.iInt1 - 1
        Read = oDB.GetRecord(dbFile, X)
        Var.lLong2 = 1
        For Var.iInt2 = 0 To 4
            Var.lLong1 = InStr(Var.lLong2, Read, ";")
            Data(Var.iInt2) = Mid(Read, Var.lLong2, Var.lLong1 - Var.lLong2)
            Var.lLong2 = Var.lLong1 + 1
        Next
        Data(5) = Mid(Read, Var.lLong2)
        Set ItemX = lstTime.ListItems.Add(, "k" & X, X + 1)
        
        With ItemX
            .SubItems(1) = Data(5) 'Track
            .SubItems(2) = Data(3) 'Driver
            .SubItems(3) = Data(0) 'Time
            .SubItems(4) = Data(2) 'Type
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

Public Sub AddExe()
    Read = oMisc.ReadINI("Misc", "ExePath", TempFile)
    If Read <> "" Then
        frmGP2Edit.Show vbModal, frmMain
    Else
        Read = ""
        Read = oMisc.ShowOpen("GP2Edit exe patch file (*.exe)|*.exe|All Files (*.*)|*.*|", Me.hwnd, ProgramDir, "GP2Edit Dos Patch File")
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
                MsgBox "This is not a valid GP2Edit Dos Patch file.", vbInformation, TH
            End If
        Else
            Exit Sub
        End If
    End If
End Sub

Public Sub DropTime(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim OldNr As Integer
Dim TimeData(0 To 4) As String
Dim Start As Long
Dim Stopp As Long
    On Error GoTo ErrHandler
    If TreeNr <> 0 Then
        SaveTrackData TreeNr
    End If
    OldNr = TreeNr
    TreeNr = Mid(TreeView1.HitTest(X, y).Key, 2, 2)
    Set TreeView1.DropHighlight = Nothing
    TreeNr = TreeNr - 10

    GetTrackData TreeNr
    If txtPath.Text = "" Then
        MsgBox "You have to add a track to this node before you can add a time.", vbInformation, TH
        Exit Sub
    End If
    Count1 = Data.GetData(1) - 1
    Read = oDB.GetRecord(dbFile, Count1)
    Start = 1
    For X = 0 To 4
        Stopp = InStr(Start, Read, ";")
        TimeData(X) = Mid(Read, Start, Stopp - Start)
        Start = Stopp + 1
    Next
    RetVal = 0
    RetVal = InStr(1, LCase(Read), "qual")
    If RetVal > 0 Then
        txtQTime.Text = TimeData(0)
        txtQTeam.Text = TimeData(4)
        txtQDriver.Text = TimeData(3)
        txtQDate.Text = TimeData(1)
    Else
        txtRTime.Text = TimeData(0)
        txtRTeam.Text = TimeData(4)
        txtRDriver.Text = TimeData(3)
        txtRDate.Text = TimeData(1)
    End If
    SaveTrackData TreeNr
    If (TreeNr <> OldNr) And (OldNr <> 0) Then
        GetTrackData OldNr
    End If
    TreeNr = OldNr
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 91
        MsgBox "You can't drop this here"
    Case Else
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
    Set TreeView1.DropHighlight = Nothing
End Sub

Public Sub DropTrack(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Path As String
Dim FileType As String
    ClearText
    On Error GoTo ErrHandler
    On Error Resume Next
    Read = ""
    Read = Data.Files(1)
    If Read <> "" Then
        FileType = LCase(Mid(Data.Files(1), Len(Data.Files(1)) - 3, 4))
        If FileType <> ".dat" Then
            MsgBox "You can only add track files this way, if you want to add a menu pic you " & vbLf & "have to use the included file manager.", vbExclamation, TH
            Set TreeView1.DropHighlight = Nothing
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
    Else
        Set TreeView1.DropHighlight = Nothing
        Exit Sub
    End If

Drop:
    If FileType = ".dat" Then
        TreeView1.Nodes(TreeView1.DropHighlight.Key).Selected = True
        TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-Track", "Track File: " & Path, 3, 3
        lstFile.Tag = ""
        SaveDropData Path
        Tracks(TreeNr - 1) = True
        Set TreeView1.DropHighlight = Nothing
    ElseIf FileType = "big" Then
        Read = oMisc.ReadINI("Track " & TreeNr, "TPath", TempFile)
        If Read <> "" Then
            TreeView1.Nodes(TreeView1.DropHighlight.Key).Selected = True
            TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-BPic", "Big Pic: " & Path, 4, 4
            lstFile.Tag = ""
            GetTrackData TreeNr
            Set imgBPic.Picture = LoadPicture(Path)
            txtBPic.Text = Path
            SaveTrackData TreeNr
        End If
        Set TreeView1.DropHighlight = Nothing
    ElseIf FileType = "small" Then
        Read = oMisc.ReadINI("Track " & TreeNr, "TPath", TempFile)
        If Read <> "" Then
            TreeView1.Nodes(TreeView1.DropHighlight.Key).Selected = True
            TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-SPic", "Small Pic: " & Path, 4, 4
            lstFile.Tag = ""
            GetTrackData TreeNr
            Set imgSPic.Picture = LoadPicture(Path)
            txtSPic.Text = Path
            SaveTrackData TreeNr
        End If
        Set TreeView1.DropHighlight = Nothing
    End If
    Set TreeView1.DropHighlight = Nothing
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
        ElseIf FileType = "small" Then
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
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub ListFiles(ByVal Show As PatFile)
Dim MyName As String
Dim vArray As Variant
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    'Remove all items from the listbox
    lstFile.ListItems.Clear

    'List all files
    MyName = Dir(MyPath, vbDirectory)
    Do While MyName <> ""
        If MyName <> "." And MyName <> ".." Then
            If LCase(MyName) <> "pagefile.sys" Then
                If (GetAttr(MyPath & MyName) And vbDirectory) <> vbDirectory Then
                    If LCase(Mid(MyName, Len(MyName) - 3, 4)) = ".dat" And ((Show = All) Or (Show = Dat)) Then
                        lstFile.ListItems.Add , MyPath & MyName, MyName, 2, 2
                    ElseIf LCase(Mid(MyName, Len(MyName) - 3, 4)) = ".bmp" And ((Show = All) Or (Show = Bmp)) Then
                        lstFile.ListItems.Add , MyPath & MyName, MyName, 3, 3
                    ElseIf LCase(Mid(MyName, Len(MyName) - 3, 4)) = ".gif" And ((Show = All) Or (Show = Gif)) Then
                        lstFile.ListItems.Add , MyPath & MyName, MyName, 3, 3
                    End If
                End If
            End If
        End If
        MyName = Dir
    Loop

    'Sort files and add them to a array
    lstFile.Sorted = True
    lstFile.Sorted = False
    ReDim vArray(lstFile.ListItems.Count, 1)
    For X = 0 To lstFile.ListItems.Count - 1
        vArray(X, 0) = lstFile.ListItems(X + 1).Key
        vArray(X, 1) = lstFile.ListItems(X + 1).Text
    Next
    'Remove all files
    lstFile.ListItems.Clear

    'Add folders
    MyName = Dir(MyPath, vbDirectory)
    lstFile.ListItems.Clear
    If Show = All Then
        Do While MyName <> ""
            If MyName <> "." And MyName <> ".." Then
                If LCase(MyName) <> "pagefile.sys" Then
                    If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
                        lstFile.ListItems.Add , "dir" & MyPath & MyName, MyName, 1, 1
                    End If
                End If
            End If
            MyName = Dir
        Loop
    End If
    'Sort folders
    lstFile.Sorted = True
    lstFile.Sorted = False
    
    'List the files
    Dim OFS As OFSTRUCT
    Dim FT_CREATE As FILETIME
    Dim FT_ACCESS As FILETIME
    Dim FT_WRITE As FILETIME
    Dim ItemX

    For X = 0 To UBound(vArray) - 1
        Var.sString1 = (vArray(X, 0))
        Var.lLong1 = OpenFile(Var.sString1, OFS, OF_READ)
        Call GetFileTime(Var.lLong1, FT_CREATE, FT_ACCESS, FT_WRITE)
        
        If LCase(Mid(vArray(X, 1), Len(vArray(X, 1)) - 3, 4)) = ".dat" Then
            Set ItemX = lstFile.ListItems.Add(, vArray(X, 0), vArray(X, 1), 2, 2)
        ElseIf LCase(Mid(vArray(X, 1), Len(vArray(X, 1)) - 3, 4)) = ".bmp" Then
            Set ItemX = lstFile.ListItems.Add(, vArray(X, 0), vArray(X, 1), 3, 3)
        ElseIf LCase(Mid(vArray(X, 1), Len(vArray(X, 1)) - 3, 4)) = ".gif" Then
            Set ItemX = lstFile.ListItems.Add(, vArray(X, 0), vArray(X, 1), 3, 3)
        End If
        ItemX.SubItems(1) = Round(FileLen(vArray(X, 0)) / 1000, 0) & " kb"
        ItemX.SubItems(2) = GetFileDateString(FT_WRITE)
        Call CloseHandle(Var.lLong1)
    Next
    If lstFile.ListItems.Count > 0 Then lstFile.ListItems(1).Selected = True
    lblMyPath = MyPath
End Sub

Public Sub ClearInfo()
    lblEvent = ""
    lblSlot = ""
    lblQual = ""
    lblMisc = ""
    lblRace = ""
    lblLen = ""
    lblWare = ""
    lblLaps = ""
    lblAuthor = ""
    lblInfoYear = ""
    lblTrackName = ""
    lblCountry = ""
    lblSetup = ""
End Sub

Public Sub SetTextProp()
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

End Sub

Public Sub GetRegValue()
Dim Temp As Button
    On Error GoTo ErrHandler
    'Set nr of times the program has been started
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "Nr")
    X = X + 1
    oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "Nr", Trim(Str(X))
    'Get deff track path (if selected)
    Read = ""
    Read = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "TrackPath")
    If Read <> "" Then Read2 = oMisc.File_Exists(Read)
    If (Read <> "") And (LCase(Read2) = "true") Then
        Drive1.Drive = Mid(Read, 1, 2)
        MyPath = Read
        ListFiles All
    Else
        Drive1.Drive = Mid(App.Path, 1, 2)
        MyPath = App.Path
        ListFiles All
    End If

    'Check how to show icons
    X = 0
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "ToolBar")
    If X <> 0 Then
        Toolbar2.Buttons(X).Value = tbrPressed
        DoEvents
        If X = 3 Then
            mnuList.Checked = True
            mnuDetails.Checked = False
            mnuLargeIcons.Checked = False
            mnuList_Click
        ElseIf X = 4 Then
            mnuList.Checked = False
            mnuDetails.Checked = True
            mnuLargeIcons.Checked = False
            mnuDetails_Click
        ElseIf X = 5 Then
            mnuList.Checked = False
            mnuDetails.Checked = False
            mnuLargeIcons.Checked = True
            mnuLargeIcons_Click
        End If
    End If

    'Check if GP2 Track Edit is installed on this system
    Read = ""
    Read = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "TrackEdit")
    If Read <> "" Then
        mnuOpenFile.Enabled = True
        mnuOpenFile.Tag = Read
    Else
        Read = ""
        Read = oReg.GetValue(HKEY_CLASSES_ROOT, ".dat", "")
        If LCase(Read) = "trackfiletype" Then
            mnuOpenFile.Enabled = True
            mnuOpenFile.Tag = ""
        End If
    End If

    'Check if Toolbar is on or off, the same with status bar, if off the hide
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "Toolbar")
    If X = 1 Then
        mnuToolbar.Checked = True
        mnuToolbar_Click
    End If
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "Statusbar")
    If X = 1 Then
        mnuStatusbar.Checked = True
        mnuStatusbar_Click
    End If

    GP2Dir = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "GP2Path")
    If GP2Dir <> "" Then Read = oMisc.File_Exists(GP2Dir & "\gp2.exe")
    If (GP2Dir = "") Or (LCase(Read) = "false") Then GP2Dir = SetGP2Folder
    If GP2Dir = "" Then
        cmdExportSettings.Enabled = False
        cmdImportSettings.Enabled = False
        Toolbar1.Buttons("Import").Enabled = False
        Toolbar1.Buttons("Export").Enabled = False
        Toolbar1.Buttons("JamCheck").Enabled = False
        mnuImport.Enabled = False
        mnuExport.Enabled = False
        mnuJamCheck.Enabled = False
        mnuJamCheck2.Enabled = False
    Else
        stbMain.Panels(3).Text = "GP2 Directory: " & GP2Dir
        GetGP2Version
    End If

    RunGP2 = ""
    RunGP2 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "RunGP2")
    If RunGP2 = "" Then
        RunGP2 = GP2Dir & "\GP2.exe"
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "RunGP2", RunGP2
    End If

    RegFileName
Exit Sub
ErrHandler:
    MsgBox Err.Number, Err.Description
End Sub

Public Sub ClearText()
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
End Sub

Public Function SetGP2Folder() As String
SelectPath:
    Read = oMisc.BrowseFolders("Select GP2 Location", Me.hwnd)
    If Read <> "" Then
        If Len(Read) = 3 Then Read = Mid(Read, 1, 2)
        Read2 = oMisc.File_Exists(Read & "\gp2.exe")
        If Read2 = True Then
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Settings", "GP2Path", Read
            stbMain.Panels(3).Text = "GP2 Directory: " & Read
            SetGP2Folder = Read

            cmdExportSettings.Enabled = True
            cmdImportSettings.Enabled = True
            Toolbar1.Buttons("Import").Enabled = True
            Toolbar1.Buttons("Export").Enabled = True
            Toolbar1.Buttons("JamCheck").Enabled = True
            mnuImport.Enabled = True
            mnuExport.Enabled = True
            mnuJamCheck.Enabled = True
            mnuJamCheck2.Enabled = True

        Else
            Var.iInt1 = MsgBox(LoadResString(105), vbRetryCancel + vbCritical, TH)
            If Var.iInt1 = vbCancel Then
                SetGP2Folder = ""
                Exit Function
            Else
                GoTo SelectPath
            End If
        End If
    End If
End Function

Public Sub MakeTempFile(ByVal sFile As String)
    On Error Resume Next
    Randomize
    X = Int((500) * Rnd)
    Kill (TempFile)
    TempFile = ProgramDir & "\File\th14" & X & ".lda"
    FileCopy sFile, TempFile
End Sub

Private Sub UpOneLevel()
'*************************************
'Function Name: UpOneLevel
'Use: Move up one level in a folder tree
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-08-17
'*************************************
Dim VolName As String, FSys As String, erg As Long
Dim VolNumber As Long, MCM As Long, FSF As Long
    ClearInfo
    If Len(MyPath) > 3 Then
        For X = Len(MyPath) - 1 To 1 Step -1
            If Mid(MyPath, X, 1) = "\" Then Exit For
        Next
        MyPath = Mid(MyPath, 1, X)
        ListFiles All
    Else
        Toolbar2.Buttons(1).Enabled = False
        lstFile.ListItems.Clear
        Read = Space(255)
        Var.iInt1 = GetLogicalDriveStrings(255, Read)
        For Var.lLong1 = 1 To Var.iInt1 Step 4
            Var.sString1 = Mid(Read, Var.lLong1, 2)
            Var.lLong2 = GetDriveType(Var.sString1)

            VolName = Space(127)
            FSys = Space(127)
            
            If Var.lLong2 <> DRIVE_REMOVABLE Then
                RetVal = GetVolumeInformation(Var.sString1 & "\", VolName, 127, VolNumber, MCM, FSF, FSys, 127)
                X = InStr(1, VolName, Chr(0))
                If X > 0 Then
                    VolName = Mid(VolName, 1, X - 1)
                Else
                    VolName = ""
                End If
            End If
            If Var.lLong2 = DRIVE_CDROM Then
                lstFile.ListItems.Add , "drive" & Var.sString1, VolName & " [" & Var.sString1 & "]", 8, 8
            ElseIf Var.lLong2 = DRIVE_FIXED Or Var.lLong2 = 1 Then
                lstFile.ListItems.Add , "drive" & Var.sString1, VolName & " [" & Var.sString1 & "]", 5, 5
            ElseIf Var.lLong2 = DRIVE_RAMDISK Then
                lstFile.ListItems.Add , "drive" & Var.sString1, VolName & " [" & Var.sString1 & "]", 7, 7
            ElseIf Var.lLong2 = DRIVE_REMOTE Then
                lstFile.ListItems.Add , "drive" & Var.sString1, VolName & " [" & Var.sString1 & "]", 6, 6
            ElseIf Var.lLong2 = DRIVE_REMOVABLE Then
                lstFile.ListItems.Add , "drive" & Var.sString1, "[" & Var.sString1 & "]", 4, 4
            End If
        Next Var.lLong1
    End If
End Sub

Private Sub JamCheck()
'*************************************
'Function Name: JamCheck
'Use: Check if all jamfiles are installed
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-08-24
'*************************************
On Error GoTo ErrHandler
    If tabMain.Tab = 1 Then
        If txtPath <> "" Then frmJamCheck.Show vbModal, frmMain
    ElseIf tabMain.Tab = 0 Then
        If LCase(Mid(lstFile.SelectedItem.Key, Len(lstFile.SelectedItem.Key) - 3, 4)) <> ".dat" Then Exit Sub
        frmJamCheck.Show vbModal, frmMain
    End If
Exit Sub
ErrHandler:
End Sub

Private Sub MoveTrackDown()
'*************************************
'Function Name: MoveTrackDown
'Use: Move a track down one level
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-11
'*************************************
Dim TempTrackData(16) As String
On Error GoTo ErrHandler

    With frmMain
        TempTrackData(0) = .txtPath
        TempTrackData(1) = .txtAdjectiv
        TempTrackData(2) = .txtBPic
        TempTrackData(3) = .txtCountry
        TempTrackData(4) = .txtLaps
        TempTrackData(5) = .txtLength
        TempTrackData(6) = .txtName
        TempTrackData(7) = .txtQDate
        TempTrackData(8) = .txtQDriver
        TempTrackData(9) = .txtQTeam
        TempTrackData(10) = .txtQTime
        TempTrackData(11) = .txtRDate
        TempTrackData(12) = .txtRDriver
        TempTrackData(13) = .txtRTeam
        TempTrackData(14) = .txtRTime
        TempTrackData(15) = .txtSPic
        TempTrackData(16) = .txtTire
    End With
    GetTrackData TreeNr + 1
    SaveTrackData TreeNr
    With frmMain
        .txtPath = TempTrackData(0)
        .txtAdjectiv = TempTrackData(1)
        .txtBPic = TempTrackData(2)
        .txtCountry = TempTrackData(3)
        .txtLaps = TempTrackData(4)
        .txtLength = TempTrackData(5)
        .txtName = TempTrackData(6)
        .txtQDate = TempTrackData(7)
        .txtQDriver = TempTrackData(8)
        .txtQTeam = TempTrackData(9)
        .txtQTime = TempTrackData(10)
        .txtRDate = TempTrackData(11)
        .txtRDriver = TempTrackData(12)
        .txtRTeam = TempTrackData(13)
        .txtRTime = TempTrackData(14)
        .txtSPic = TempTrackData(15)
        .txtTire = TempTrackData(16)
    End With
    SaveTrackData TreeNr + 1
    LoadFile
    TreeView1.Nodes("t" & TreeNr + 11).Selected = True
    TreeView1_NodeClick TreeView1.SelectedItem
Exit Sub
ErrHandler:
End Sub

Private Sub MoveTrackUp()
'*************************************
'Function Name: MoveTrackUp
'Use: move a track up one level
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-11
'*************************************
Dim TempTrackData(16) As String
On Error GoTo ErrHandler

    With frmMain
        TempTrackData(0) = .txtPath
        TempTrackData(1) = .txtAdjectiv
        TempTrackData(2) = .txtBPic
        TempTrackData(3) = .txtCountry
        TempTrackData(4) = .txtLaps
        TempTrackData(5) = .txtLength
        TempTrackData(6) = .txtName
        TempTrackData(7) = .txtQDate
        TempTrackData(8) = .txtQDriver
        TempTrackData(9) = .txtQTeam
        TempTrackData(10) = .txtQTime
        TempTrackData(11) = .txtRDate
        TempTrackData(12) = .txtRDriver
        TempTrackData(13) = .txtRTeam
        TempTrackData(14) = .txtRTime
        TempTrackData(15) = .txtSPic
        TempTrackData(16) = .txtTire
    End With
    GetTrackData TreeNr - 1
    SaveTrackData TreeNr
    With frmMain
        .txtPath = TempTrackData(0)
        .txtAdjectiv = TempTrackData(1)
        .txtBPic = TempTrackData(2)
        .txtCountry = TempTrackData(3)
        .txtLaps = TempTrackData(4)
        .txtLength = TempTrackData(5)
        .txtName = TempTrackData(6)
        .txtQDate = TempTrackData(7)
        .txtQDriver = TempTrackData(8)
        .txtQTeam = TempTrackData(9)
        .txtQTime = TempTrackData(10)
        .txtRDate = TempTrackData(11)
        .txtRDriver = TempTrackData(12)
        .txtRTeam = TempTrackData(13)
        .txtRTime = TempTrackData(14)
        .txtSPic = TempTrackData(15)
        .txtTire = TempTrackData(16)
    End With
    SaveTrackData TreeNr - 1
    LoadFile
    TreeView1.Nodes("t" & TreeNr + 9).Selected = True
    TreeView1_NodeClick TreeView1.SelectedItem
Exit Sub
ErrHandler:
End Sub
