VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gp2 Track Handler v1.6"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   9075
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      WhatsThisHelpID =   2
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgMisc"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   23
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            Object.Tag             =   ""
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Import"
            Object.ToolTipText     =   "Import data from Gp2"
            Object.Tag             =   ""
            ImageKey        =   "Import"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Export"
            Object.ToolTipText     =   "Export data to Gp2"
            Object.Tag             =   ""
            ImageKey        =   "Export"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Gp2Edit"
            Object.ToolTipText     =   "Add/Edit Gp2Edit file"
            Object.Tag             =   ""
            ImageKey        =   "GP2Edit"
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
            ImageKey        =   "Jam"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Setup"
            Object.ToolTipText     =   "CC Car Settings"
            Object.Tag             =   ""
            ImageKey        =   "Setup"
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Backup"
            Object.ToolTipText     =   "Backup a Track"
            Object.Tag             =   ""
            ImageKey        =   "Backup"
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
            ImageKey        =   "Home"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Down"
            Object.ToolTipText     =   "Move track down one level"
            Object.Tag             =   ""
            ImageKey        =   "Down"
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Up"
            Object.ToolTipText     =   "Move track up one level"
            Object.Tag             =   ""
            ImageKey        =   "Up"
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Gp2"
            Object.ToolTipText     =   "Run Gp2"
            Object.Tag             =   ""
            ImageKey        =   "GP2"
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            Object.Tag             =   ""
            ImageKey        =   "Help"
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
            ImageKey        =   "Exit"
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstTmpFile 
      Height          =   2205
      Left            =   7680
      Sorted          =   -1  'True
      TabIndex        =   132
      Top             =   840
      Width           =   1215
      Visible         =   0   'False
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   131
      Top             =   6525
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5115
            MinWidth        =   5115
            Text            =   "Gp2 Track Handler v1.6 © Viktor Gars"
            TextSave        =   "Gp2 Track Handler v1.6 © Viktor Gars"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "Gp2 Version:"
            TextSave        =   "Gp2 Version:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5556
            MinWidth        =   5556
            Text            =   "Gp2 Directory:"
            TextSave        =   "Gp2 Directory:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTrackPic 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   -10000
      TabIndex        =   124
      Top             =   800
      Width           =   5655
      Begin VB.Frame fraPicture 
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   120
         TabIndex        =   126
         Top             =   360
         Width           =   5415
         Begin VB.TextBox txtBPic 
            Height          =   285
            Left            =   3600
            TabIndex        =   128
            Top             =   4440
            Width           =   1815
            Visible         =   0   'False
         End
         Begin VB.TextBox txtSPic 
            Height          =   285
            Left            =   3600
            TabIndex        =   127
            Top             =   4800
            Width           =   1815
            Visible         =   0   'False
         End
         Begin VB.Image picMenuPic 
            BorderStyle     =   1  'Fixed Single
            Height          =   3600
            Left            =   0
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4800
         End
      End
      Begin ComctlLib.TabStrip tabPic 
         Height          =   5655
         Left            =   0
         TabIndex        =   125
         Top             =   0
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   9975
         _Version        =   327682
         BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
            NumTabs         =   2
            BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Lage Menu Picture"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Small Menu Picture"
               Key             =   ""
               Object.Tag             =   ""
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraFileManager 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   -10000
      TabIndex        =   85
      Top             =   800
      Width           =   5655
      Begin VB.Frame fraFileInfo 
         Caption         =   "Track File Info"
         Height          =   2000
         Left            =   0
         TabIndex        =   92
         Top             =   3670
         WhatsThisHelpID =   6
         Width           =   5655
         Begin VB.TextBox lblInfoYear 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   3840
            MaxLength       =   4
            TabIndex        =   104
            ToolTipText     =   "Click to edit"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox lblEvent 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   3840
            TabIndex        =   103
            ToolTipText     =   "Click to edit"
            Top             =   735
            Width           =   1695
         End
         Begin VB.TextBox lblCountry 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            TabIndex        =   102
            ToolTipText     =   "Click to edit"
            Top             =   240
            Width           =   1700
         End
         Begin VB.CommandButton cmdSaveGP2Info 
            Enabled         =   0   'False
            Height          =   345
            Left            =   5160
            Picture         =   "frmMain.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Save GP2Info"
            Top             =   0
            WhatsThisHelpID =   5
            Width           =   375
         End
         Begin VB.TextBox lblLaps 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   3
            TabIndex        =   100
            ToolTipText     =   "Click to edit"
            Top             =   720
            Width           =   1700
         End
         Begin VB.TextBox lblLen 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   4
            TabIndex        =   99
            ToolTipText     =   "Click to edit"
            Top             =   960
            Width           =   1700
         End
         Begin VB.TextBox lblWare 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   5
            TabIndex        =   98
            ToolTipText     =   "Click to edit"
            Top             =   1200
            Width           =   1700
         End
         Begin VB.TextBox lblQual 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   8
            TabIndex        =   97
            ToolTipText     =   "Click to edit"
            Top             =   1440
            Width           =   1700
         End
         Begin VB.TextBox lblRace 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            MaxLength       =   8
            TabIndex        =   96
            ToolTipText     =   "Click to edit"
            Top             =   1680
            Width           =   1700
         End
         Begin VB.TextBox lblTrackName 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   1250
            TabIndex        =   95
            ToolTipText     =   "Click to edit"
            Top             =   480
            Width           =   1700
         End
         Begin VB.TextBox lblMisc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   3000
            MultiLine       =   -1  'True
            TabIndex        =   94
            ToolTipText     =   "Click to Edit"
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox lblSlot 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   200
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   93
            ToolTipText     =   "Click to edit"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblSetup 
            Height          =   200
            Left            =   3840
            TabIndex        =   119
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblAuthor 
            Height          =   200
            Left            =   3840
            TabIndex        =   118
            Top             =   960
            Width           =   1695
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblInfoText 
            AutoSize        =   -1  'True
            Caption         =   "Misc:"
            Height          =   200
            Index           =   11
            Left            =   3000
            TabIndex        =   117
            Top             =   1230
            Width           =   375
         End
         Begin VB.Label lblInfoText 
            AutoSize        =   -1  'True
            Caption         =   "Car &Setup:"
            Height          =   195
            Index           =   12
            Left            =   3000
            TabIndex        =   116
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Year:"
            Height          =   200
            Index           =   7
            Left            =   3000
            TabIndex        =   115
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Track Name:"
            Height          =   200
            Index           =   1
            Left            =   120
            TabIndex        =   114
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblInfoText 
            Caption         =   "Best Lap &Race:"
            Height          =   200
            Index           =   6
            Left            =   120
            TabIndex        =   113
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "E&vent:"
            Height          =   200
            Index           =   8
            Left            =   3000
            TabIndex        =   112
            Top             =   735
            Width           =   735
         End
         Begin VB.Label lblInfoText 
            Caption         =   "Best Lap &Qual:"
            Height          =   200
            Index           =   5
            Left            =   120
            TabIndex        =   111
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "Tyre &Ware:"
            Height          =   200
            Index           =   4
            Left            =   120
            TabIndex        =   110
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "L&ength (m):"
            Height          =   200
            Index           =   3
            Left            =   120
            TabIndex        =   109
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Laps:"
            Height          =   200
            Index           =   2
            Left            =   120
            TabIndex        =   108
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Country:"
            Height          =   200
            Index           =   0
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblInfoText 
            Caption         =   "&Author:"
            Height          =   200
            Index           =   10
            Left            =   3000
            TabIndex        =   106
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblInfoText 
            AutoSize        =   -1  'True
            Caption         =   "&Slot:"
            Height          =   200
            Index           =   9
            Left            =   4440
            TabIndex        =   105
            Top             =   480
            Width           =   495
         End
      End
      Begin VB.Frame frameNorm 
         Caption         =   " File Manager"
         Height          =   3675
         Left            =   0
         TabIndex        =   86
         Top             =   0
         Width           =   5655
         Begin ComctlLib.Toolbar Toolbar2 
            Height          =   390
            Left            =   3960
            TabIndex        =   87
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
                  ImageKey        =   "UpOneLevel"
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
                  ImageKey        =   "Small"
                  Style           =   2
                  Value           =   1
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Details"
                  Object.ToolTipText     =   "Details"
                  Object.Tag             =   ""
                  ImageKey        =   "List"
                  Style           =   2
               EndProperty
               BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   "Large"
                  Object.ToolTipText     =   "Large Icons"
                  Object.Tag             =   ""
                  ImageKey        =   "Big"
                  Style           =   2
               EndProperty
            EndProperty
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   3735
         End
         Begin ComctlLib.ListView lstFile 
            Height          =   2750
            Left            =   120
            TabIndex        =   88
            Top             =   600
            WhatsThisHelpID =   4
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
         Begin VB.Label Label1 
            Height          =   135
            Left            =   2280
            TabIndex        =   91
            Top             =   4080
            Width           =   1335
         End
         Begin VB.Label lblMyPath 
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   3360
            Width           =   5415
         End
      End
      Begin VB.Frame fraMenuPic 
         Caption         =   "Menu Picture"
         Height          =   2000
         Left            =   0
         TabIndex        =   120
         Top             =   3670
         Width           =   5655
         Begin VB.Label lblPicInfo 
            Height          =   255
            Left            =   2280
            TabIndex        =   121
            Top             =   360
            Width           =   2655
         End
         Begin VB.Image imgPre 
            Height          =   1560
            Left            =   120
            Stretch         =   -1  'True
            Top             =   360
            Width           =   2040
         End
      End
      Begin VB.Frame fraNoSupport 
         Height          =   2000
         Left            =   0
         TabIndex        =   122
         Top             =   3670
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
            TabIndex        =   123
            Top             =   720
            Width           =   4455
         End
      End
   End
   Begin VB.Frame fraData 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   -10000
      TabIndex        =   49
      Top             =   800
      Width           =   5535
      Begin VB.Frame frameData 
         Caption         =   "Track Data"
         Height          =   2175
         Left            =   0
         TabIndex        =   79
         Top             =   2880
         Width           =   3135
         Begin Gp2_Track_Handler.UpDown updLaps 
            Height          =   285
            Left            =   120
            TabIndex        =   129
            Top             =   480
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            Max             =   126
         End
         Begin VB.TextBox txtLength 
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   81
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtTire 
            Height          =   285
            Left            =   120
            MaxLength       =   5
            TabIndex        =   80
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Laps (3-126)"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tra&ck Length (0-9999 m)"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   83
            Top             =   840
            Width           =   1755
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tyre &Ware (14000-40000)"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   82
            Top             =   1440
            Width           =   1830
         End
      End
      Begin VB.Frame frameInfo 
         Caption         =   "Track Info"
         Height          =   2775
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Width           =   3135
         Begin VB.ComboBox txtAdjectiv 
            Height          =   315
            Left            =   120
            TabIndex        =   130
            Text            =   "Combo1"
            Top             =   2280
            Width           =   2775
         End
         Begin VB.TextBox txtPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   255
            TabIndex        =   74
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   120
            TabIndex        =   73
            Top             =   1080
            Width           =   2775
         End
         Begin VB.TextBox txtCountry 
            Height          =   285
            Left            =   120
            TabIndex        =   72
            Top             =   1680
            Width           =   2775
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Adjective:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   78
            Top             =   2040
            Width           =   705
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Country:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   77
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "&Track Name:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   76
            Top             =   840
            Width           =   930
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Track Path"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame fraQual 
         Caption         =   "Lap Time Data - Qual"
         Height          =   2775
         Left            =   3360
         TabIndex        =   61
         Top             =   0
         Width           =   2175
         Begin VB.TextBox txtQDriver 
            Height          =   285
            Left            =   120
            MaxLength       =   23
            TabIndex        =   66
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtQTeam 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   65
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtQDate 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   64
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdSaveQual 
            Height          =   300
            Left            =   1800
            Picture         =   "frmMain.frx":040C
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtQTime 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   62
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Driver"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Team"
            Height          =   195
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Date (e.g 1999-06-19)"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   2040
            Width           =   1560
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Time (e.g. 1:24.145)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   1440
            Width           =   1425
         End
      End
      Begin VB.Frame fraRace 
         Caption         =   "Lap Time Data - Race"
         Height          =   2775
         Left            =   3360
         TabIndex        =   51
         Top             =   2880
         Width           =   2175
         Begin VB.TextBox txtRDriver 
            Height          =   285
            Left            =   120
            MaxLength       =   23
            TabIndex        =   56
            Top             =   480
            Width           =   1935
         End
         Begin VB.TextBox txtRTeam 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   55
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtRDate 
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   54
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdSaveRace 
            Height          =   300
            Left            =   1800
            Picture         =   "frmMain.frx":050E
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   0
            Width           =   300
         End
         Begin VB.TextBox txtRTime 
            Height          =   285
            Left            =   120
            MaxLength       =   8
            TabIndex        =   52
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Date (e.g. 1999-06-19)"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   2040
            Width           =   1605
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Team"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   405
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Driver"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Time (e.g. 1:24.145)"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1440
            Width           =   1425
         End
      End
      Begin VB.CommandButton cmdJamCheck 
         Caption         =   "&Jam Check"
         Height          =   375
         Left            =   0
         TabIndex        =   50
         Top             =   5160
         Width           =   1095
      End
   End
   Begin VB.Frame fraLapTime 
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   -10000
      TabIndex        =   45
      Top             =   800
      Width           =   5655
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   0
         TabIndex        =   48
         Top             =   5280
         Width           =   1035
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Time"
         Height          =   315
         Left            =   1200
         TabIndex        =   47
         Top             =   5280
         Width           =   1035
      End
      Begin ComctlLib.ListView lstTime 
         DragIcon        =   "frmMain.frx":0610
         Height          =   5175
         Left            =   0
         TabIndex        =   46
         Top             =   0
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
            Object.Width           =   706
         EndProperty
      End
   End
   Begin VB.Frame fraMisc 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   -10000
      TabIndex        =   3
      Top             =   800
      Width           =   5655
      Begin VB.Frame framePlayer 
         Caption         =   "Player Car"
         Height          =   5175
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   2775
         Begin VB.HScrollBar hscPRPower 
            Height          =   255
            LargeChange     =   20
            Left            =   120
            Max             =   1579
            TabIndex        =   34
            Top             =   600
            Value           =   780
            Width           =   2415
         End
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   120
            TabIndex        =   33
            Top             =   4200
            Width           =   2415
         End
         Begin VB.HScrollBar hscPQPower 
            Height          =   255
            LargeChange     =   20
            Left            =   120
            Max             =   1579
            TabIndex        =   32
            Top             =   1200
            Value           =   790
            Width           =   2415
         End
         Begin VB.HScrollBar hscWeight 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   4000
            Min             =   401
            TabIndex        =   31
            Top             =   2280
            Value           =   1313
            Width           =   2415
         End
         Begin VB.HScrollBar hscPGrip 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   1000
            TabIndex        =   30
            Top             =   2880
            Value           =   198
            Width           =   2415
         End
         Begin VB.HScrollBar hscPitSpeed 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   201
            Min             =   1
            TabIndex        =   29
            Top             =   3480
            Value           =   50
            Width           =   2415
         End
         Begin VB.CheckBox chkNoLimit 
            Caption         =   "No Speed Limit"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   3840
            Width           =   2500
         End
         Begin VB.CheckBox chkUPower 
            Caption         =   "Use Selected Team Power"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1560
            Width           =   2500
         End
         Begin VB.CommandButton cmdExportSettings 
            Caption         =   "&Export"
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   4320
            Width           =   1035
         End
         Begin VB.CommandButton cmdImportSettings 
            Caption         =   "&Import"
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   4700
            Width           =   1035
         End
         Begin VB.CommandButton cmdDefaultSettings 
            Caption         =   "GP2 Default"
            Height          =   315
            Left            =   1500
            TabIndex        =   24
            Top             =   4320
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Power in Race"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label lblPQPower 
            Alignment       =   1  'Right Justify
            Caption         =   "790"
            Height          =   195
            Left            =   2055
            TabIndex        =   43
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Power in Qual"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   990
         End
         Begin VB.Label lblPRPower 
            Alignment       =   1  'Right Justify
            Caption         =   "780"
            Height          =   195
            Left            =   2055
            TabIndex        =   41
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblWeight 
            AutoSize        =   -1  'True
            Caption         =   "Car Weight"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   2040
            Width           =   795
         End
         Begin VB.Label lblWeight2 
            Alignment       =   1  'Right Justify
            Caption         =   "1313lb (596Kg)"
            Height          =   195
            Left            =   1335
            TabIndex        =   39
            Top             =   2040
            Width           =   1200
         End
         Begin VB.Label lblGrip2 
            AutoSize        =   -1  'True
            Caption         =   "Grip"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   2640
            Width           =   285
         End
         Begin VB.Label lblGrip 
            Alignment       =   1  'Right Justify
            Caption         =   "198"
            Height          =   195
            Left            =   2055
            TabIndex        =   37
            Top             =   2640
            Width           =   480
         End
         Begin VB.Label lblPitSpeed 
            Alignment       =   1  'Right Justify
            Caption         =   "50mph (80km/h)"
            Height          =   195
            Left            =   1095
            TabIndex        =   36
            Top             =   3240
            Width           =   1440
         End
         Begin VB.Label lblPit 
            AutoSize        =   -1  'True
            Caption         =   "Pit Speed"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   3240
            Width           =   1050
         End
      End
      Begin VB.Frame frameGlobal 
         Caption         =   "Global Settings"
         Height          =   5175
         Left            =   2880
         TabIndex        =   4
         Top             =   0
         Width           =   2775
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   120
            TabIndex        =   11
            Top             =   3600
            Width           =   2415
         End
         Begin VB.HScrollBar hscCWeight 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   4000
            Min             =   401
            TabIndex        =   10
            Top             =   600
            Value           =   1313
            Width           =   2415
         End
         Begin VB.HScrollBar hscQRace 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   9
            Top             =   1200
            Value           =   5
            Width           =   2415
         End
         Begin VB.CheckBox chk0as1 
            Caption         =   "Show car 1 as 0"
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "Always save track record"
            Height          =   195
            Left            =   240
            TabIndex        =   7
            Top             =   3120
            Width           =   2175
         End
         Begin VB.HScrollBar Slider1 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   2099
            Min             =   1900
            TabIndex        =   6
            Top             =   1800
            Value           =   1994
            Width           =   2415
         End
         Begin VB.CheckBox chkCCFuel 
            Caption         =   "Show CC Car fuel load"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label lblPro 
            AutoSize        =   -1  'True
            Caption         =   "Pro"
            Height          =   180
            Left            =   1800
            TabIndex        =   22
            Top             =   4560
            Width           =   240
         End
         Begin VB.Label lblSemiPro 
            AutoSize        =   -1  'True
            Caption         =   "Semi-Pro"
            Height          =   180
            Left            =   1800
            TabIndex        =   21
            Top             =   4320
            Width           =   630
         End
         Begin VB.Label lblAmateur 
            AutoSize        =   -1  'True
            Caption         =   "Amateur"
            Height          =   180
            Left            =   1800
            TabIndex        =   20
            Top             =   4080
            Width           =   585
         End
         Begin VB.Label lblRookie 
            AutoSize        =   -1  'True
            Caption         =   "Rookie"
            Height          =   180
            Index           =   3
            Left            =   1800
            TabIndex        =   19
            Top             =   3840
            Width           =   510
         End
         Begin VB.Label lblAce 
            AutoSize        =   -1  'True
            Caption         =   "Ace"
            Height          =   180
            Left            =   1800
            TabIndex        =   18
            Top             =   4800
            Width           =   285
         End
         Begin VB.Image R 
            Height          =   180
            Index           =   0
            Left            =   240
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
            Index           =   2
            Left            =   630
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
            Index           =   4
            Left            =   1020
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
            Index           =   6
            Left            =   1410
            Top             =   3840
            Width           =   195
         End
         Begin VB.Image A 
            Height          =   180
            Index           =   0
            Left            =   240
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
            Index           =   2
            Left            =   630
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
            Index           =   4
            Left            =   1020
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
            Index           =   6
            Left            =   1410
            Top             =   4080
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   1
            Left            =   435
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
            Index           =   2
            Left            =   630
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   2
            Left            =   630
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   3
            Left            =   825
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   3
            Left            =   825
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   4
            Left            =   1020
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
            Index           =   5
            Left            =   1215
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   5
            Left            =   1215
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image AC 
            Height          =   180
            Index           =   6
            Left            =   1410
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
            Index           =   0
            Left            =   240
            Top             =   4800
            Width           =   195
         End
         Begin VB.Image P 
            Height          =   180
            Index           =   0
            Left            =   240
            Top             =   4560
            Width           =   195
         End
         Begin VB.Image S 
            Height          =   180
            Index           =   1
            Left            =   435
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
            Index           =   3
            Left            =   825
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
            Index           =   5
            Left            =   1215
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
            Index           =   0
            Left            =   240
            Top             =   4320
            Width           =   195
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Car Weight"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   795
         End
         Begin VB.Label lblCWeight 
            Alignment       =   1  'Right Justify
            Caption         =   "1313lb (596kg)"
            Height          =   195
            Left            =   1455
            TabIndex        =   16
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Quick Race Length"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   960
            Width           =   1395
         End
         Begin VB.Label lblQuick 
            Alignment       =   1  'Right Justify
            Caption         =   "5%"
            Height          =   195
            Left            =   2175
            TabIndex        =   14
            Top             =   960
            Width           =   480
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Year for this Season"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            Width           =   1425
         End
         Begin VB.Label lblYear 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1994"
            Height          =   195
            Left            =   1695
            TabIndex        =   12
            Top             =   1560
            Width           =   960
         End
      End
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   420
      WhatsThisHelpID =   3
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   10821
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
   Begin ComctlLib.TabStrip tabMain 
      Height          =   6100
      Left            =   3120
      TabIndex        =   2
      Top             =   420
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   10769
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "File Manager"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Track Data"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Menu Pics"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Misc Settings"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Lap Time Database"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
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
            Picture         =   "frmMain.frx":091A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0EC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":126A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":14FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":178E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1A20
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
            Picture         =   "frmMain.frx":1CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":20DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":21F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2302
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2414
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2526
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2638
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
         NumListImages   =   23
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":274A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":285C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":296E
            Key             =   "Track"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2A80
            Key             =   "TrackPic"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2B92
            Key             =   "New"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2CA4
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2DB6
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2EC8
            Key             =   "Import"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2FDA
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":30EC
            Key             =   "GP2Edit"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3406
            Key             =   "Jam"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3518
            Key             =   "Setup"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3832
            Key             =   "Backup"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3B4C
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3C5E
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3D70
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3E82
            Key             =   "GP2"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":419C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":44EE
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4808
            Key             =   "Small"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":491A
            Key             =   "List"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4A2C
            Key             =   "Big"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4B3E
            Key             =   "UpOneLevel"
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
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImp 
         Caption         =   "&Import"
         Begin VB.Menu mnuImport 
            Caption         =   "Data &from Gp2..."
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuImpRec 
            Caption         =   "&Lap time data from .rec file..."
         End
      End
      Begin VB.Menu mnuExp 
         Caption         =   "&Export"
         Begin VB.Menu mnuExport 
            Caption         =   "&Data to Gp2..."
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuExpRec 
            Caption         =   "&Lap time data to .rec file..."
         End
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortOpen 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit2 
      Caption         =   "&Edit"
      Index           =   0
      Begin VB.Menu mnuEdit 
         Caption         =   "&Gp2 Path..."
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Default Track Path..."
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Track Editor Path..."
         Index           =   2
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Settings for ""Run Gp2"" Button..."
         Index           =   4
      End
   End
   Begin VB.Menu mnuView2 
      Caption         =   "&View"
      Begin VB.Menu mnuView 
         Caption         =   "&List"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Details"
         Index           =   4
      End
      Begin VB.Menu mnuView 
         Caption         =   "Lar&ge Icons"
         Index           =   5
      End
      Begin VB.Menu mnuView 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Flat Toolbar"
         Index           =   7
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
      Begin VB.Menu mnuSep5 
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
         Begin VB.Menu mnuSep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditQSetup2 
            Caption         =   "Edit Qual Setup..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEditRSetup2 
            Caption         =   "Edit Race Setup..."
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuRemove2 
            Caption         =   "Remove Setup"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup Track..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuToolUninstall 
         Caption         =   "&Uninstall Track"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGp2Edit 
         Caption         =   "A&dd/Edit Gp2Edit File..."
      End
   End
   Begin VB.Menu mnuTopHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help..."
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "VG Software Online"
         Index           =   2
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "A&bout..."
         Index           =   4
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
         Begin VB.Menu mnuSep9 
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
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup2 
         Caption         =   "&Backup Track..."
      End
      Begin VB.Menu mnuPopupUninstall 
         Caption         =   "&Uninstall Track"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckSum 
         Caption         =   "&Write Checksum"
      End
      Begin VB.Menu mnuViewChecksum 
         Caption         =   "&View Checksum..."
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Edit Track..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep13 
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
Private File As String
Private RunGp2 As String

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
    lblPit.Enabled = chkNoLimit.Value - 1
    lblPitSpeed.Enabled = chkNoLimit.Value - 1
    hscPitSpeed.Enabled = chkNoLimit.Value - 1
End Sub

Private Sub chkUPower_Click()
    hscPQPower.Enabled = chkUPower.Value - 1
    hscPRPower.Enabled = chkUPower.Value - 1
    lblPRPower.Enabled = chkUPower.Value - 1
    lblPQPower.Enabled = chkUPower.Value - 1
    Label2(2).Enabled = chkUPower.Value - 1
    Label10.Enabled = chkUPower.Value - 1
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
    Exp.Gp2FileNum = FreeFile
    Open Gp2Dir & "\gp2.exe" For Binary As Exp.Gp2FileNum
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
    ExportCCFuel
    Close Exp.Gp2FileNum
    Read = oFile.FileExists(Gp2Dir & "\f1gstate.sav")
    If Read = True Then
        Exp.F1FileNum = FreeFile
        Open Gp2Dir & "\f1gstate.sav" For Binary As Exp.F1FileNum
        ExportQuickRace
        Close Exp.F1FileNum
        WriteCheckSum Gp2Dir & "\f1gstate.sav"
    End If
    GetMisc
End Sub

Private Sub cmdImportSettings_Click()
    Exp.Gp2FileNum = FreeFile
    Open Gp2Dir & "\gp2.exe" For Binary As Exp.Gp2FileNum
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
    ImportCCFuel
    Read = oFile.FileExists(Gp2Dir & "\f1gstate.sav")
    If Read = True Then
        Exp.F1FileNum = FreeFile
        Open Gp2Dir & "\f1gstate.sav" For Binary As Exp.F1FileNum
        ImportQuick
        Close Exp.F1FileNum
    End If
    Close Exp.Gp2FileNum
    GetMisc
End Sub

Private Sub cmdJamCheck_Click()
    JamCheck
End Sub

Public Sub cmdSaveGp2Info_Click()
    frmMain.MousePointer = 11
    MakeText
    WriteCheckSum frmMain.lstFile.SelectedItem.Key
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
    If txtRTime.Text <> "" Then
        Read = txtRTime.Text & ";" & txtRDate.Text & ";Race;" & txtRDriver.Text & ";" & txtRTeam.Text & ";" & txtName
        oDB.SaveNew dbFile, Read
        LoadTimeData
    Else
        MsgBox "You must have a time to save a time.", vbInformation, TH
    End If
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
        MsgBox "Device unavailable.", vbExclamation, TH
        Drive1.Drive = "C:"
    Case Else
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: Drive1_Change()", vbCritical, TH & " - Error"
            Resume Next
    End Select
End Sub

Private Sub Form_Load()
    ProgStart = True
    frmMain.MousePointer = 11
    On Error Resume Next
    tabMain.Tabs(1).Selected = True
    NewTree
    Toolbar1.Buttons("Up").Enabled = False
    Toolbar1.Buttons("Down").Enabled = False
    frmMain.Show

    ProgramDir = App.Path
    If Right(ProgramDir, 1) = "\" Then ProgramDir = Mid(ProgramDir, 1, Len(ProgramDir) - 1)
    'ProgramDir = "C:\My Documents\Mina Program\Visual Basic\Gp2 Track Handler v16\TestCenter"

    oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Settings", "Path", ProgramDir

    FullRowSelect lstTime
    FullRowSelect lstFile

    Set oData = New GP2Info

    Set oDB = New clsLapTime

    MkDir ProgramDir & "\File"
    MkDir ProgramDir & "\Bat"
    dbFile = ProgramDir & "\Time.lda"
    '*****************
    App.HelpFile = ProgramDir & "\Help.hlp"
    '*****************
    GetRegValue
    SetTextProp
    LoadAdj

    fraFileInfo.Enabled = False
    LoadTimeData
    Dim cmdLine As String
    cmdLine = Command()
    If cmdLine <> "" Then
        If oFile.GetFilePart(cmdLine, GetExt) = ".ths" Then
            OpenCommandFile
        Else
            GoTo Norm
        End If
    Else
        'Normal start
Norm:
        NewFile
        LoadRecent
        DriveHelpDefault
    End If
    frmMain.MousePointer = 0
    ProgStart = False
Exit Sub

ErrHandler:
    frmMain.MousePointer = 0
    MsgBox "Error Number: " & Err.Number & vbCrLf & _
        "Error Description: " & Err.Description & vbCrLf & _
        "Error Source: Form_Load()", vbCritical, " - Error"
    ProgStart = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
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

Private Sub lblInfoText_Click(Index As Integer)
    Select Case Index
    Case 0
        lblCountry.SetFocus
    Case 1
        lblTrackName.SetFocus
    Case 2
        lblLaps.SetFocus
    Case 3
        lblLen.SetFocus
    Case 4
        lblWare.SetFocus
    Case 5
        lblQual.SetFocus
    Case 6
        lblRace.SetFocus
    Case 7
        lblInfoYear.SetFocus
    Case 8
        lblEvent.SetFocus
    Case 9
        lblSlot.SetFocus
    Case 11
        lblMisc.SetFocus
    End Select
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
        SetAttr lstFile.SelectedItem.Key, vbNormal
        Kill lstFile.SelectedItem.Key
        TempInt = lstFile.SelectedItem.Index
        ListFiles All
        lstFile.ListItems(TempInt).Selected = True
    End If
    Me.MousePointer = 0
End Sub

Private Sub lstFile_Click()
    On Error GoTo ErrHandler
    If LCase(Mid(lstFile.SelectedItem.Key, 1, 3)) = "dir" Or LCase(Mid(lstFile.SelectedItem.Key, 1, 5)) = "drive" Then
        UnsupportedFile
        fraNoSupport.Visible = False
        fraFileInfo.Visible = True
        FileType = 3
        ClearInfo
        Exit Sub
    End If
    If oFile.GetFilePart(lstFile.SelectedItem.Key, GetExt) = ".dat" Then
        FileType = 0
        ClearInfo
        If ReadGp2Info(lstFile.SelectedItem.Key) = True Then
            ShowInfo
            RetVal = CheckSetup(lstFile.SelectedItem.Key)
            If RetVal = True Then
                lblSetup.Caption = "Yes"
            Else
                lblSetup.Caption = "No"
            End If
            mnuEditQSetup.Enabled = RetVal
            mnuEditRSetup.Enabled = RetVal
            mnuRemove.Enabled = RetVal
            mnuEditQSetup2.Enabled = RetVal
            mnuEditRSetup2.Enabled = RetVal
            mnuRemove2.Enabled = RetVal
            SupportedFile
        Else
            FileType = 4
            UnsupportedFile
        End If
    ElseIf (oFile.GetFilePart(lstFile.SelectedItem.Key, GetExt) = ".gif") Or (oFile.GetFilePart(lstFile.SelectedItem.Key, GetExt) = ".bmp") Then
        fraFileInfo.Visible = False
        fraFileInfo.Enabled = False
        fraNoSupport.Visible = False
        cmdSaveGP2Info.Enabled = False
        imgRealSize.Picture = LoadPicture(lstFile.SelectedItem.Key)
        If ((imgRealSize.Width / 15 = 640) And (imgRealSize.Height / 15 = 480)) Then
            lblPicInfo.Caption = "Large Menu Picture"
            Set imgPre.Picture = LoadPicture(lstFile.SelectedItem.Key)
            FileType = 2
            fraMenuPic.Visible = True
            SupportedFile
        ElseIf ((imgRealSize.Width / 15 = 440) And (imgRealSize.Height / 15 = 330)) Then
            lblPicInfo.Caption = "Small Menu Picture"
            Set imgPre.Picture = LoadPicture(lstFile.SelectedItem.Key)
            FileType = 1
            fraMenuPic.Visible = True
            SupportedFile
        Else
            lstFile.OLEDragMode = ccOLEDragAutomatic
            lstFile.OLEDragMode = ccOLEDragManual
            fraMenuPic.Visible = False
            fraNoSupport.Visible = True
            FileType = 4
            UnsupportedFile
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
            "Error Source: lstFile_Click()", vbCritical, TH & " - Error"
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
    Me.MousePointer = 11
    On Error GoTo ErrHandler
    If Mid(lstFile.SelectedItem.Key, 1, 3) = "dir" Then
        MyPath = Mid(lstFile.SelectedItem.Key, 4)
        ListFiles All
    ElseIf LCase(Mid(lstFile.SelectedItem.Key, 1, 5)) = "drive" Then
        Toolbar2.Buttons(1).Enabled = True
        Drive1.Drive = Mid(lstFile.SelectedItem.Key, 6)
        ListFiles All
    ElseIf oFile.GetFilePart(lstFile.SelectedItem.Key, GetExt) = ".dat" Then
        For X = 0 To 15
            If Tracks(X) = False Then
                SaveDropData lstFile.SelectedItem.Key, X + 1
                LoadFile
                ClearText
                MousePointer = 0
                Exit Sub
            End If
        Next
    End If
MousePointer = 0
Exit Sub
ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: lstFile_DblClick()", vbCritical, TH & " - Error"
    Me.MousePointer = 0
End Sub

Private Sub lstFile_ItemClick(ByVal Item As ComctlLib.ListItem)
    If File <> lstFile.SelectedItem.Text Then
        lstFile_Click
        File = lstFile.SelectedItem.Text
    End If
End Sub

Private Sub lstFile_KeyUp(KeyCode As Integer, Shift As Integer)
    If tabMain.Tabs(1).Selected = True Then
        If (KeyCode = 40) Or (KeyCode = 38) Then lstFile_Click
        If KeyCode = 13 Then lstFile_DblClick
        If KeyCode = 8 Then
            UpOneLevel
        End If
    End If
End Sub

Private Sub lstFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bButton = Button
End Sub

Private Sub lstFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 2) And (FileType = 0) Then
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuBackup_Click()
Dim sZipName As String
Dim sFilePath As String
Dim BackupDef As Boolean
    If tabMain.Tabs(1).Selected = True Then

        If FileType <> 0 Then Exit Sub
        sZipName = Mid(lstFile.SelectedItem.Text, 1, Len(lstFile.SelectedItem.Text) - 3) & "zip"
        sFilePath = lstFile.SelectedItem.Key
        GoTo ZipFile
    
    ElseIf tabMain.Tabs(2).Selected = True Then
        
        If txtPath.Text = "" Then Exit Sub
        sZipName = oFile.GetFilePart(txtPath.Text, GetFileName)
        sZipName = Mid(sZipName, 1, Len(sZipName) - 3) & "zip"
        sFilePath = txtPath.Text
        GoTo ZipFile
    
    End If
Exit Sub
ZipFile:
    sZipName = oFile.ShowSave("Zip Files (*.Zip)|*.zip|", "zip", Me.hWnd, , "Create Zip File", sZipName)
    If sZipName = "" Then Exit Sub
    RetVal = MsgBox("Do you want to backup the oridginal Gp2 Jam files?", vbYesNo, TH)
    If RetVal = vbYes Then
        BackupDef = True
    Else
        BackupDef = False
    End If
    If sZipName <> "" Then
        Me.MousePointer = 11
        Load frmProgress
        frmProgress.Show
        frmProgress.Caption = sZipName
        DoEvents
        BackupTrack sFilePath, sZipName, BackupDef
        Unload frmProgress
        Me.MousePointer = 0
    End If
End Sub

Private Sub mnuBackup2_Click()
    mnuBackup_Click
End Sub

Private Sub mnuCheckSum_Click()
    Me.MousePointer = 11
    WriteCheckSum lstFile.SelectedItem.Key
    Me.MousePointer = 0
End Sub

Private Sub mnuDelete_Click()
Dim TempInt As Integer
    tVar.iInt = MsgBox("Are you sure you want to delete this file?", vbYesNo, "Confirm File Delete")
    If tVar.iInt = vbYes Then
        Kill lstFile.SelectedItem.Key
        TempInt = lstFile.SelectedItem.Index - 1
        ListFiles All
        If TempInt > 0 Then lstFile.ListItems(TempInt).Selected = True
    End If
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case Index
    Case 0
        Gp2Dir = SetGp2Folder
        If Gp2Dir <> "" Then GetGp2Version
    Case 1
        Read = oFile.BrowseFolders("Select Track Directory", Me.hWnd)
        If Read <> "" Then
            If Right(Read, 1) <> "\" Then Read = Read & "\"
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Settings", "TrackPath", Read
            Read2 = "LoadTrackPathNow-Flag"
            Drive1.Drive = Mid(Read, 1, 2)
            Read2 = ""
            MyPath = Read
            ListFiles All
        End If
    Case 2
        Read = oFile.ShowOpen("Gp2 Track Editor (*.exe)|*.exe|", Me.hWnd, , "Select Gp2 Track Editor")
        If Read = "" Then Exit Sub
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Settings", "TrackEdit", Read
        mnuOpenFile.Tag = Read
        mnuOpenFile.Enabled = True
    Case 4
        Read = oFile.ShowOpen("All Application Files (*.exe)|*.exe|Gp2 (gp2.exe)|gp2.exe|Gp2Lap (gp2lap.exe)|gp2lap.exe|", Me.hWnd, Gp2Dir, "Select Application")
        If Read = "" Then Exit Sub
        RunGp2 = Read
        Toolbar1.Buttons(18).ToolTipText = RunGp2
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Settings", "RunGp2", RunGp2
    End Select
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
    sFileName = oFile.ShowSave("Lap Time Data File (*.rec)|*.rec|All files (*.*)|*.*|", "rec", Me.hWnd, , "Save Record File")
    If sFileName = "" Then Exit Sub
    FileNum = FreeFile
    Open Gp2Dir & "\f1gstate.sav" For Binary As FileNum
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

Private Sub mnuGp2Edit_Click()
    AddExe
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Select Case Index
    Case 0
        RetVal = WinHelp(frmMain.hWnd, App.HelpFile, HELP_CONTENTS, ByVal 5)
    Case 2
        INetLink "http://www.vgsoftware.com/", Me.hWnd
    Case 4
        frmAbout.Show vbModal, frmMain
    End Select
End Sub

Private Sub mnuImpRec_Click()
    Read = oFile.ShowOpen("Lap Time Data File (*.rec)|*.rec|All files (*.*)|*.*|", Me.hWnd, , "Select a Record File")
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

Private Sub mnuNewSetup_Click()
    frmSetup.Show vbModal, frmMain
End Sub

Private Sub mnuNewSetup2_Click()
    mnuNewSetup_Click
End Sub

Private Sub mnuOpenFile_Click()
    If mnuOpenFile.Tag = "" Then
        RetVal = ShellExecute(Me.hWnd, "open", lstFile.SelectedItem.Key, vbNullString, vbNullString, 1)
    Else
        Read = oFile.GetShortName(lstFile.SelectedItem.Key)
        Read2 = ""
        For X = Len(mnuOpenFile.Tag) To 1 Step -1
            If Mid(mnuOpenFile.Tag, X, 1) = "\" Then Exit For
        Next
        Read2 = Mid(mnuOpenFile.Tag, 1, X - 1)
        RetVal = ShellExecute(Me.hWnd, "open", mnuOpenFile.Tag, Read, Read2, 1)
    End If
End Sub

Private Sub mnuPopupUninstall_Click()
    mnuToolUninstall_Click
End Sub

Private Sub mnuRemove_Click()
    DeteteSetup lstFile.SelectedItem.Key
    WriteCheckSum frmMain.lstFile.SelectedItem.Key
End Sub

Private Sub mnuRemove2_Click()
    mnuRemove_Click
End Sub

Private Sub mnuRename_Click()
    lstFile.StartLabelEdit
End Sub

Private Sub mnuSetupFile_Click()
    Read = oFile.ShowOpen("Gp2 Setup File (*.cs*)|*.cs*|All Files (*.*)|*.*|", Me.hWnd, , "Open CarSetup")
    If Read = "" Then Exit Sub
    Load frmSetup
    OpenSetup Read
    frmSetup.Show vbModal, frmMain
End Sub

Private Sub mnuSetupFile2_Click()
    mnuSetupFile_Click
End Sub

Private Sub mnuShortOpen_Click(Index As Integer)
    If oFile.FileExists(mnuShortOpen(Index).Tag) = True Then
        ShowFiles mnuShortOpen(Index).Tag
    Else
        MsgBox "The file """ & mnuShortOpen(Index).Tag & """ was not found.", vbCritical, TH
    End If
End Sub

Private Sub mnuCCCarSetup_Click()
    On Error Resume Next
    If tabMain.Tabs(1).Selected = True Then
        Read = oFile.GetFilePart(lstFile.SelectedItem.Key, GetExt)
        If Read <> ".dat" Then Exit Sub
    Else
        If txtPath.Text = "" Then Exit Sub
    End If
    frmCCSetup.Show vbModal, frmMain
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExport_Click()
    SaveTrackData TreeNr
    SaveMisc
    frmExport.Show , frmMain
End Sub

Private Sub mnuImport_Click()
    SaveTrackData TreeNr
    SaveMisc
    frmImport.Show vbModal, frmMain
End Sub

Private Sub mnuNew_Click()
    TreeView1.Nodes.Item(1).Selected = True
    TreeView1_NodeClick TreeView1.Nodes(1)
    NewFile
    LoadFile
    frmMain.Caption = TH & " v1.6"
    Unload frmExport
    Unload frmImport
End Sub

Private Sub mnuOpen_Click()
    Read = oFile.ShowOpen("Supported Files (*.ths;*.set)|*.ths;*.set|Gp2 Track Handler Files (*.ths)|*.ths|TrackSet Files (*.set)|*.set|All Files (*.*)|*.*|", Me.hWnd)
    If Read = "" Then Exit Sub
    ShowFiles Read
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
    SaveMisc
    SaveFile
End Sub

Private Sub mnuSaveas_Click()
    SaveTrackData TreeNr
    SaveMisc
    SaveFileAs
End Sub

Private Sub mnuToolUninstall_Click()
    RetVal = MsgBox("Are you sure you want to delete all extra jam files and the track file?", vbYesNo + vbExclamation, TH)
    If RetVal = vbYes Then
        MousePointer = 11
        Load frmProgress
        frmProgress.Show
        frmProgress.Caption = "Uninstall " & oFile.GetFilePart(lstFile.SelectedItem.Key, GetFileName)
        DoEvents
        Uninstall (lstFile.SelectedItem.Key)
        ListFiles All
        Unload frmProgress
        MousePointer = 0
    End If
End Sub

Private Sub mnuTrackSettings_Click()
    On Error Resume Next
    frmCCSetup.Show vbModal, frmMain
End Sub

Private Sub mnuView_Click(Index As Integer)
    If Index > 1 And Index < 7 Then
        mnuView(3).Checked = False
        mnuView(4).Checked = False
        mnuView(5).Checked = False
        mnuView(Index).Checked = True
        oReg.SaveValue HKEY_CURRENT_USER, REG_DWORD, "Software\VG Software\Gp2 Track Handler\Settings", "ToolBar", , ByVal (Index)
    End If
    Select Case Index
    Case 3
        lstFile.View = lvwList
        Toolbar2.Buttons(3).Value = tbrPressed
    Case 4
        lstFile.View = lvwReport
        Toolbar2.Buttons(4).Value = tbrPressed
    Case 5
        lstFile.View = lvwIcon
        Toolbar2.Buttons(5).Value = tbrPressed
    Case 7
        If mnuView(7).Checked = True Then
            mnuView(7).Checked = False
            FlatToolbar Toolbar1
            FlatToolbar Toolbar2
        Else
            mnuView(7).Checked = True
            FlatToolbar Toolbar1
            FlatToolbar Toolbar2
        End If
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Settings", "FlatToolbar", mnuView(7).Checked
    End Select
End Sub

Private Sub mnuViewChecksum_Click()
Dim iCheck1 As Integer
Dim iCheck2 As Integer
Dim sCheck1 As String
Dim sCheck2 As String
    FileNum = FreeFile
    Open lstFile.SelectedItem.Key For Binary As FileNum
    Get #FileNum, FileLen(lstFile.SelectedItem.Key) - 3, iCheck1
    Get #FileNum, FileLen(lstFile.SelectedItem.Key) - 1, iCheck2
    sCheck1 = String(1, " ")
    Get #FileNum, FileLen(lstFile.SelectedItem.Key) - 3, sCheck1
    Read = String(1, " ")
    Get #FileNum, FileLen(lstFile.SelectedItem.Key) - 2, Read
    sCheck1 = Hex(Asc(sCheck1)) & " " & Hex(Asc(Read))
    
    sCheck2 = String(1, " ")
    Get #FileNum, FileLen(lstFile.SelectedItem.Key) - 1, sCheck2
    Read = String(1, " ")
    Get #FileNum, FileLen(lstFile.SelectedItem.Key), Read
    sCheck2 = Hex(Asc(sCheck2)) & " " & Hex(Asc(Read))
    
    Close FileNum
    MsgBox "Checksum1=" & Hex(iCheck1) & " (" & sCheck1 & ")" & vbLf & "Checksum2=" & Hex(iCheck2) & " (" & sCheck2 & ")", vbInformation, "CheckSum"
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

Private Sub tabPic_Click()
    If tabPic.SelectedItem.Index = 2 Then
        Set picMenuPic.Picture = LoadPicture(txtSPic.Text)
        picMenuPic.Height = 2475
        picMenuPic.Width = 3300
    Else
        Set picMenuPic.Picture = LoadPicture(txtBPic.Text)
        picMenuPic.Height = 3600
        picMenuPic.Width = 4800
    End If
End Sub

Private Sub tabPic_GotFocus()
    If tabMain.Tabs(2).Selected = True Then
        txtName.SetFocus
    End If
End Sub

Private Sub tabMain_Click()
    If tabMain.Tabs(1).Selected = True Then

        fraData.Left = -10000
        fraTrackPic.Left = -10000
        fraFileManager.Left = 3300
        fraLapTime.Left = -10000
        fraMisc.Left = -10000

        mnuRand.Enabled = True
        Toolbar1.Buttons("Home").Enabled = True
        mnuTrackSettings.Enabled = True
        mnuBackup.Enabled = True
        mnuJamCheck2.Enabled = True
        Toolbar1.Buttons("Backup").Enabled = True
        Toolbar1.Buttons("Setup").Enabled = True
        Toolbar1.Buttons("JamCheck").Enabled = True
        Toolbar1.Buttons("Up").Enabled = False
        Toolbar1.Buttons("Down").Enabled = False
        If TreeView1.Nodes.Count > 10 Then
            TreeView1.Nodes(1).Selected = True
            TreeNr = 0
        End If
        ClearText
        GoTo Save
    ElseIf tabMain.Tabs(2).Selected = True Then
        fraData.Left = 3300
        fraTrackPic.Left = -10000
        fraFileManager.Left = -10000
        fraLapTime.Left = -10000
        fraMisc.Left = -10000
        
        Toolbar1.Buttons("Home").Enabled = False
        mnuRand.Enabled = False
        mnuTrackSettings.Enabled = True
        mnuBackup.Enabled = True
        mnuJamCheck2.Enabled = True
        Toolbar1.Buttons("Backup").Enabled = True
        Toolbar1.Buttons("Setup").Enabled = True
        Toolbar1.Buttons("JamCheck").Enabled = True
        GoTo Save
    ElseIf tabMain.Tabs(3).Selected = True Then
        fraData.Left = -10000
        fraTrackPic.Left = 3300
        fraFileManager.Left = -10000
        fraLapTime.Left = -10000
        fraMisc.Left = -10000
    ElseIf tabMain.Tabs(4).Selected = True Then
        fraData.Left = -10000
        fraTrackPic.Left = -10000
        fraFileManager.Left = -10000
        fraLapTime.Left = -10000
        fraMisc.Left = 3300
    ElseIf tabMain.Tabs(5).Selected = True Then
        fraData.Left = -10000
        fraTrackPic.Left = -10000
        fraFileManager.Left = -10000
        fraLapTime.Left = 3300
        fraMisc.Left = -10000
    End If
    mnuRand.Enabled = False
    Toolbar1.Buttons("Home").Enabled = False
    mnuTrackSettings.Enabled = False
    mnuBackup.Enabled = False
    mnuJamCheck2.Enabled = False
    Toolbar1.Buttons("Backup").Enabled = False
    Toolbar1.Buttons("Setup").Enabled = False
    Toolbar1.Buttons("JamCheck").Enabled = False
Save:
    If ProgStart = False Then
        SaveMisc
        GetMisc
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
    Case "Gp2"
        Read = ""
        For X = Len(RunGp2) To 0 Step -1
            If Mid(RunGp2, X, 1) = "\" Then Exit For
        Next
        Read = Mid(RunGp2, 1, X)
        RetVal = ShellExecute(frmMain.hWnd, "open", RunGp2, vbNullString, Read, 1)
    Case "Gp2Edit"
        AddExe
    Case "Help"
        RetVal = WinHelp(frmMain.hWnd, App.HelpFile, HELP_CONTENTS, ByVal 0)
    Case "Backup"
        mnuBackup_Click
    Case "JamCheck"
        JamCheck
    Case "Setup"
        mnuCCCarSetup_Click
    Case "Home"
        Read = ""
        Read = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "TrackPath")
        If Read <> "" Then Read2 = oFile.FileExists(Read)
        If (Read <> "") And (LCase(Read2) = "true") Then
            Drive1.Drive = Mid(Read, 1, 2)
            MyPath = Read
            ListFiles All
            Toolbar2.Buttons(1).Enabled = True
        End If
    Case "Down"
        MoveTrack False
    Case "Up"
        MoveTrack True
    End Select
End Sub

Public Sub DriveHelpDefault()
    LoadGp2Aid "11111110111111011111101001110100001"
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
    Case "Up"
        UpOneLevel
    Case Else
        mnuView_Click (Button.Index)
    End Select
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
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
            WriteINI "Track " & TreeNr, "BPic", "", TempFile
            WriteINI "Track " & TreeNr, "SPic", "", TempFile
            Set picMenuPic.Picture = Nothing
            txtBPic.Text = ""
            txtSPic.Text = ""
            SaveTrackData TreeNr
            TreeView1.Nodes("t" & TreeNr + 10).Selected = True
            TreeView1_NodeClick TreeView1.SelectedItem
        Else
            TreeView1.Nodes.Remove (TreeView1.SelectedItem.Index)
            If KeyName = "bpic" Then
                WriteINI "Track " & TreeNr, "BPic", "", TempFile
                txtBPic.Text = ""
                Set picMenuPic.Picture = Nothing
            ElseIf KeyName = "spic" Then
                WriteINI "Track " & TreeNr, "SPic", "", TempFile
                txtSPic.Text = ""
                Set picMenuPic.Picture = Nothing
            End If
        End If
    End If
Exit Sub
ErrHandler:
    Exit Sub
End Sub

Public Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
    If Node.Key = "r" Then
        If TreeNr <> 0 Then SaveTrackData TreeNr
        fraData.Enabled = False
        fraTrackPic.Enabled = False
        tabMain.Tabs(1).Selected = True
        TreeNr = 0
        Toolbar1.Buttons("Down").Enabled = False
        Toolbar1.Buttons("Up").Enabled = False
        ClearText
    Else
        If TreeNr <> 0 Then SaveTrackData TreeNr
        fraData.Enabled = True
        fraTrackPic.Enabled = True

        TreeNr = Mid(TreeView1.SelectedItem.Key, 2, 2)
        TreeNr = TreeNr - 10
        GetTrackData TreeNr

        If InStr(1, TreeView1.SelectedItem.Key, "BPic") Then
            tabMain.Tabs(3).Selected = True
            tabPic.Tabs(1).Selected = True
        ElseIf InStr(1, TreeView1.SelectedItem.Key, "SPic") Then
            tabMain.Tabs(3).Selected = True
            tabPic.Tabs(2).Selected = True
        Else
            tabMain.Tabs(2).Selected = True
        End If
        If (TreeView1.SelectedItem.Children = 0) And (Len(TreeView1.SelectedItem.Key) = 3) Then
            fraData.Enabled = False
        Else
            fraData.Enabled = True
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

Private Sub TreeView1_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tabMain.Tabs(5).Selected = True Then
        DropTime Data, Effect, Button, Shift, X, Y
    ElseIf tabMain.Tabs(1).Selected = True Then
        DropTrack Data, Effect, Button, Shift, X, Y
    End If
End Sub

Private Sub TreeView1_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
End Sub

Private Sub txtCountry_GotFocus()
    TextSelected
End Sub

Private Sub txtLength_GotFocus()
    TextSelected
End Sub

Private Sub txtName_GotFocus()
    If tabMain.Tabs(1).Selected = True Then
        Drive1.SetFocus
    ElseIf tabMain.Tabs(2).Selected = True Then
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
            "Error Source: txtName_LostFocus()", vbCritical, TH & " - Error"
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

Public Sub SaveDropData(ByVal Path As String, ByVal Pos As Integer)
    ClearText
    txtPath.Text = Path
    txtCountry.Text = TrackInfo.Country
    If IsNumeric(TrackInfo.Laps) = True Then
        If lblLaps < 3 Then
            updLaps.Value = 3
        ElseIf lblLaps > 126 Then
            updLaps.Value = 126
        Else
            updLaps.Value = Int(TrackInfo.Laps)
        End If
    Else
        updLaps.Value = 3
    End If
    txtName.Text = TrackInfo.Name
    txtTire.Text = TrackInfo.Tyre
    txtLength.Text = TrackInfo.LengthMeters
    txtAdjectiv.Text = GetAdjectiv(TrackInfo.Country)
    txtRTime.Text = TrackInfo.LapRecord
    txtQTime.Text = TrackInfo.LapRecordQualify

    SaveTrackData Pos
End Sub

Public Sub LoadRecent()
Dim vArray As Variant
Dim X As Integer

    vArray = oReg.GetAllValues(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Files")
    If Not IsArray(vArray) Then Exit Sub
    mnuSep3.Visible = True
    For X = 1 To mnuShortOpen.Count - 1
        Unload mnuShortOpen(X)
    Next
    Count1 = 1
    For X = 0 To 2
        If vArray(X, 1) <> "" Then
            Load mnuShortOpen(Count1)
            mnuShortOpen(Count1).Visible = True
            mnuShortOpen(Count1).Caption = oFile.GetFilePart(vArray(Count1 - 1, 1), GetFileName)
            mnuShortOpen(Count1).Tag = vArray(X, 1)
            Count1 = Count1 + 1
        End If
    Next
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
    Read = "#GP2INFO|Name|" & lblTrackName & "|Country|" & lblCountry & "|Created|Created by Track Editor written by Paul Hoad see (License.txt about distributing this track)" & "|Author|" & lblAuthor & _
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
    Read2 = String(3800 - Len(Read), Chr(0))
    Read = Read & Read2

    Put #FileNum, 1, Read
    Close FileNum
End Sub

Private Sub OpenCommandFile()
Dim GetOpen As String
    On Error GoTo ErrHandler
   
    FileInfo.Path = Command()
    FileInfo.Name = oFile.GetFilePart(FileInfo.Path, GetFileName)
    FileInfo.Saved = True
    FileInfo.Import = False
    Randomize
    X = Int((500) * Rnd)
    TempFile = ProgramDir & "\File\th16" & Trim(Str(X)) & ".lda"
    FileCopy FileInfo.Path, TempFile
    LoadFile
    RecentFile FileInfo.Path
    LoadRecent
    Me.Caption = TH & " v1.6 [" & Trim(FileInfo.Name) & "]"
Exit Sub

ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: OpenCommandFile()", vbCritical, TH & "Error"
End Sub

Public Sub LoadTimeData()
Dim Data(0 To 5) As String
Dim ItemX
Dim iSep As Long
Dim iLoop As Integer

    lstTime.ListItems.Clear
    tVar.iInt = oDB.RecCount(dbFile)
    For X = 0 To tVar.iInt - 1
        Read = oDB.GetRecord(dbFile, X)
        tVar.lLong = 1
        For iLoop = 0 To 4
            iSep = InStr(tVar.lLong, Read, ";")
            Data(iLoop) = Mid(Read, tVar.lLong, iSep - tVar.lLong)
            tVar.lLong = iSep + 1
        Next
        Data(5) = Mid(Read, tVar.lLong)
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
    Read = ReadINI("Misc", "ExePath", TempFile)
    If Read <> "" Then
        frmGP2Edit.Show vbModal, frmMain
    Else
        Read = ""
        Read = oFile.ShowOpen("Gp2Edit exe patch file (*.exe)|*.exe|All Files (*.*)|*.*|", Me.hWnd, , "Gp2Edit Dos Patch File")
        If Read <> "" Then
            FileNum = FreeFile
            Open Read For Binary As FileNum
            Read2 = String(12, " ")
            Get #FileNum, 45445, Read2
            Close FileNum
            If Read2 = "Steven Young" Then
                WriteINI "Misc", "ExePath", Read, TempFile
                frmGP2Edit.Show vbModal, frmMain
            Else
                MsgBox "This is not a valid Gp2Edit Dos Patch file.", vbInformation, TH
            End If
        Else
            Exit Sub
        End If
    End If
End Sub

Public Sub DropTime(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim OldNr As Integer
Dim TimeData(0 To 4) As String
Dim Start As Long
Dim Stopp As Long
    On Error GoTo ErrHandler
    If TreeNr <> 0 Then
        SaveTrackData TreeNr
    End If
    OldNr = TreeNr
    TreeNr = Mid(TreeView1.HitTest(X, Y).Key, 2, 2)
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
            "Error Source: DropTime()", vbCritical, TH & " - Error"
    End Select
    Set TreeView1.DropHighlight = Nothing
End Sub

Public Sub DropTrack(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Path As String
    On Error Resume Next
    Read = ""
    Read = Data.Files(1)
    If Read <> "" Then
        FileType = oFile.GetFilePart(Data.Files(1), GetExt)
        If FileType <> ".dat" Then
            MsgBox "You can only add track files this way, if you want to add a menu pic you " & vbLf & "have to use the included file manager.", vbExclamation, TH
            Set TreeView1.DropHighlight = Nothing
            Exit Sub
        End If
        Read = ReadGp2Info(Data.Files(1))
        If Read = False Then
            Set TreeView1.DropHighlight = Nothing
            Exit Sub
        End If
        Path = Data.Files(1)
    Else
        Path = lstFile.SelectedItem.Key
        'FileType = lstFile.Tag
    End If
    On Error GoTo ErrHandler

    If TreeView1.DropHighlight.Key <> "r" Then
        TreeNr = Mid(TreeView1.DropHighlight.Key, 2, 2)
        TreeNr = TreeNr - 10
    Else
        Set TreeView1.DropHighlight = Nothing
        Exit Sub
    End If

    If FileType = 0 Then
        RemoveNodes Mid(TreeView1.DropHighlight.Key, 1, 3)
        TreeView1.Nodes(Mid(TreeView1.DropHighlight.Key, 1, 3)).Selected = True
        TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-Track", "Track File: " & Path, 3, 3
        SaveDropData Path, TreeNr
        Tracks(TreeNr - 1) = True
    ElseIf FileType = 2 Then
        Read = ReadINI("Track " & TreeNr, "TPath", TempFile)
        If Read <> "" Then
            TreeView1.Nodes(Mid(TreeView1.DropHighlight.Key, 1, 3)).Selected = True
            TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-BPic", "Big Pic: " & Path, 4, 4
            WriteINI "Track " & TreeNr, "BPic", Path, TempFile
            txtBPic.Text = Path
        End If
    ElseIf FileType = 1 Then
        Read = ReadINI("Track " & TreeNr, "TPath", TempFile)
        If Read <> "" Then
            TreeView1.Nodes(Mid(TreeView1.DropHighlight.Key, 1, 3)).Selected = True
            TreeView1.Nodes.Add TreeView1.SelectedItem.Key, tvwChild, TreeView1.SelectedItem.Key & "-SPic", "Small Pic: " & Path, 4, 4
            WriteINI "Track " & TreeNr, "SPic", Path, TempFile
            txtSPic.Text = Path
        End If
    End If
    TreeNr = 0
    Set TreeView1.DropHighlight = Nothing
    TreeView1.Nodes(1).Selected = True
    TreeView1_NodeClick TreeView1.SelectedItem
    LoadFile
    ClearText
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 35602
      If FileType = 1 Then
        TreeView1.Nodes.Remove (Mid(TreeView1.DropHighlight.Key, 1, 3) & "-SPic")
        Resume
      ElseIf FileType = 2 Then
        TreeView1.Nodes.Remove (Mid(TreeView1.DropHighlight.Key, 1, 3) & "-BPic")
        Resume
      Else
        Set TreeView1.DropHighlight = Nothing
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
          "Error Desctiption: " & Err.Description & vbLf & _
          "Error Source: DropTrack()", vbCritical, TH & " - Error"
      End If
    Case Else
      Set TreeView1.DropHighlight = Nothing
      MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: DropTrack()", vbCritical, TH & " - Error"
    End Select
End Sub

Public Sub ClearInfo()
    TrackInfo.Author = ""
    TrackInfo.Event = ""
    TrackInfo.Slot = ""
    TrackInfo.LapRecordQualify = ""
    TrackInfo.Desc = ""
    TrackInfo.LapRecord = ""
    TrackInfo.LengthMeters = ""
    TrackInfo.Tyre = ""
    TrackInfo.Laps = ""
    TrackInfo.Year = ""
    TrackInfo.Name = ""
    TrackInfo.Country = ""
    ShowInfo
End Sub

Public Sub SetTextProp()
    X = GetWindowLong(txtLength.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtLength.hWnd, GWL_STYLE, X)

    X = GetWindowLong(txtTire.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtTire.hWnd, GWL_STYLE, X)

    X = GetWindowLong(lblLen.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(lblLen.hWnd, GWL_STYLE, X)

    X = GetWindowLong(lblWare.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(lblWare.hWnd, GWL_STYLE, X)

    X = GetWindowLong(lblLaps.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(lblLaps.hWnd, GWL_STYLE, X)

    X = GetWindowLong(txtQTime.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtQTime.hWnd, GWL_STYLE, X)

    X = GetWindowLong(txtRTime.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtRTime.hWnd, GWL_STYLE, X)

    X = GetWindowLong(txtQDate.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtQDate.hWnd, GWL_STYLE, X)

    X = GetWindowLong(txtRDate.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtRDate.hWnd, GWL_STYLE, X)
End Sub

Public Sub GetRegValue()
Dim Temp As Button
Dim bValue As Boolean

    On Error GoTo ErrHandler
    'Set nr of times the program has been started
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "Nr")
    X = X + 1
    oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Settings", "Nr", Trim(Str(X))
    'Get deff track path (if selected)
    Read = ""
    Read = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "TrackPath")
    If Read <> "" Then Read2 = oFile.FileExists(Read)
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
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "ToolBar")
    If X <> 0 Then
        Toolbar2.Buttons(X).Value = tbrPressed
        DoEvents
        mnuView_Click (X)
    End If

    'Check if Gp2 Track Edit is installed on this system
    Read = ""
    Read = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "TrackEdit")
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
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "Toolbar")
    If X = 1 Then
        mnuView_Click (0)
    End If

    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "Statusbar")
    If X = 1 Then
        mnuView_Click (1)
    End If

    'Check if toolbar is flat or not (def=flat)
    bValue = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "FlatToolbar")
    If bValue = True Then
        mnuView_Click (7)
    End If

    Gp2Dir = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "Gp2Path")
    If Gp2Dir <> "" Then Read = oFile.FileExists(Gp2Dir & "\gp2.exe")
    If (Gp2Dir = "") Or (LCase(Read) = "false") Then Gp2Dir = SetGp2Folder
    If Gp2Dir = "" Then
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
        StatusBar1.Panels(3).Text = "Gp2 Directory: " & Gp2Dir
        GetGp2Version
    End If

    RunGp2 = ""
    RunGp2 = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "RunGp2")
    If RunGp2 = "" Then
        RunGp2 = Gp2Dir & "\Gp2.exe"
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Settings", "RunGp2", RunGp2
    End If
    Toolbar1.Buttons(18).ToolTipText = RunGp2
    RegFileName
Exit Sub
ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: GetRegValue()", vbCritical, TH & " - Error"
End Sub

Public Sub ClearText()
    With frmMain
        .txtAdjectiv.Text = ""
        .txtPath = ""
        .txtBPic = ""
        .txtSPic = ""
        .txtName.Text = ""
        .txtCountry.Text = ""
        .updLaps.Value = 3
        .txtLength.Text = ""
        .txtTire.Text = ""
        .txtQDate.Text = ""
        .txtQDriver.Text = ""
        .txtQTeam.Text = ""
        .txtQTime.Text = ""
        .txtRDate.Text = ""
        .txtRDriver.Text = ""
        .txtRTeam.Text = ""
        .txtRTime.Text = ""
    End With
End Sub

Private Function SetGp2Folder() As String
SelectPath:
    Read = oFile.BrowseFolders("Select Gp2 Location", Me.hWnd)
    If Read <> "" Then
        If Len(Read) = 3 Then Read = Mid(Read, 1, 2)
        Read2 = oFile.FileExists(Read & "\gp2.exe")
        If Read2 = True Then
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Settings", "Gp2Path", Read
            StatusBar1.Panels(3).Text = "Gp2 Directory: " & Read
            SetGp2Folder = Read

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
            tVar.iInt = MsgBox("Gp2.exe not found!", vbRetryCancel + vbCritical, TH)
            If tVar.iInt = vbCancel Then
                SetGp2Folder = ""
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
    TempFile = ProgramDir & "\File\th16" & X & ".lda"
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
Dim VolName As String
Dim DriveLen As Integer

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
        DriveLen = GetLogicalDriveStrings(255, Read)
        For X = 1 To DriveLen Step 4
            Read2 = Mid(Read, X, 2)
            tVar.lLong = GetDriveType(Read2)
            If tVar.lLong <> DRIVE_REMOVABLE Then
                VolName = Space(20)
                RetVal = GetVolumeInformation(Read2 & "\", VolName, 20, vbNull, vbNull, vbNull, vbNullString, 0)
                If InStr(1, VolName, Chr(0)) <> 0 Then
                    VolName = StrConv(Mid(VolName, 1, InStr(1, VolName, Chr(0)) - 1), vbProperCase)
                Else
                    VolName = ""
                End If
            End If
            If tVar.lLong = DRIVE_CDROM Then
                lstFile.ListItems.Add , "drive" & Read2, Read2 & " [" & VolName & "]", 8, 8
            ElseIf tVar.lLong = DRIVE_FIXED Or tVar.lLong = 1 Then
                lstFile.ListItems.Add , "drive" & Read2, Read2 & " [" & VolName & "]", 5, 5
            ElseIf tVar.lLong = DRIVE_RAMDISK Then
                lstFile.ListItems.Add , "drive" & Read2, Read2 & " [" & VolName & "]", 7, 7
            ElseIf tVar.lLong = DRIVE_REMOTE Then
                lstFile.ListItems.Add , "drive" & Read2, Read2 & " [" & VolName & "]", 6, 6
            ElseIf tVar.lLong = DRIVE_REMOVABLE Then
                lstFile.ListItems.Add , "drive" & Read2, Read2, 4, 4
            End If
        Next X
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
    If tabMain.Tabs(2).Selected = True Then
        If txtPath <> "" Then frmJamCheck.Show vbModal, frmMain
    ElseIf tabMain.Tabs(1).Selected = True Then
        If oFile.GetFilePart(lstFile.SelectedItem.Key, GetExt) <> ".dat" Then Exit Sub
        frmJamCheck.Show vbModal, frmMain
    End If
Exit Sub
ErrHandler:
End Sub

Private Sub MoveTrack(ByVal Up As Boolean)
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
        TempTrackData(1) = .txtAdjectiv.Text
        TempTrackData(2) = .txtBPic
        TempTrackData(3) = .txtCountry
        TempTrackData(4) = .updLaps.Value
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
    If Up = True Then
        GetTrackData TreeNr - 1
    Else
        GetTrackData TreeNr + 1
    End If
    SaveTrackData TreeNr
    With frmMain
        .txtPath = TempTrackData(0)
        .txtAdjectiv.Text = TempTrackData(1)
        .txtBPic = TempTrackData(2)
        .txtCountry = TempTrackData(3)
        .updLaps.Value = TempTrackData(4)
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
    If Up = True Then
        SaveTrackData TreeNr - 1
    Else
        SaveTrackData TreeNr + 1
    End If
    LoadFile
    If Up = True Then
        TreeView1.Nodes("t" & TreeNr + 9).Selected = True
    Else
        TreeView1.Nodes("t" & TreeNr + 11).Selected = True
    End If
    TreeView1_NodeClick TreeView1.SelectedItem
Exit Sub
ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: MoveTrack()", vbCritical, TH & " - Error"
End Sub

Private Sub RemoveNodes(sKey As String)
    On Error Resume Next
    TreeView1.Nodes.Remove (sKey & "-BPic")
    TreeView1.Nodes.Remove (sKey & "-SPic")
    TreeView1.Nodes.Remove (sKey & "-Track")
End Sub

Public Sub ShowFiles(ByVal sFileName As String)
    Read2 = String(20, " ")
    FileNum = FreeFile
    Open sFileName For Binary As FileNum
    Get #FileNum, 1, Read2
    Close FileNum
    X = InStr(1, Read2, "[Track 1]")
    If X = 0 Then
        Read2 = ProgramDir & "\File\" & Minute(Time) & Second(Time) & "tmp.ths"
        If FileLen(sFileName) = 14336 Then
            WinTrack2TH sFileName, Read2
        ElseIf FileLen(sFileName) = 14288 Then
            Conv1 sFileName, Read2
        ElseIf FileLen(sFileName) = 9826 Then
            Conv2 sFileName, Read2
        Else
            MsgBox "This file is not supported by Gp2 Track Handler.", vbInformation, TH
            Exit Sub
        End If
        TempFile = Read2
    Else
        MakeTempFile sFileName
        FileInfo.Saved = True
        FileInfo.Path = sFileName
        FileInfo.Name = oFile.GetFilePart(sFileName, GetFileName)
    End If
    Me.Caption = TH & " v1.6 [" & oFile.GetFilePart(sFileName, GetFileName) & "]"
    RecentFile sFileName
    LoadFile
    GetMisc
    LoadRecent
End Sub

Private Sub LoadAdj()
    txtAdjectiv.Clear
    FileNum = FreeFile
    Open ProgramDir & "\Adjectiv.ini" For Input As FileNum
    Do Until EOF(FileNum)
        Line Input #FileNum, Read
        If (Read <> "") And (Read <> "[Adjectiv]") And (Mid(Read, 1, 1) <> ";") Then
            X = InStr(1, Read, "=")
            If X <> 0 Then txtAdjectiv.AddItem Mid(Read, X + 1)
        End If
    Loop
    Close FileNum
End Sub

Private Sub ShowInfo()
    With frmMain
        .lblTrackName = TrackInfo.Name
        .lblCountry = TrackInfo.Country
        .lblLaps = TrackInfo.Laps
        .lblLen = TrackInfo.LengthMeters
        .lblWare = TrackInfo.Tyre
        .lblQual = TrackInfo.LapRecordQualify
        .lblRace = TrackInfo.LapRecord
        .lblInfoYear = TrackInfo.Year
        .lblSlot = TrackInfo.Slot
        .lblAuthor = TrackInfo.Author
        .lblMisc = TrackInfo.Desc
        .lblEvent = TrackInfo.Event
        .lblSetup = ""
    End With
End Sub

Private Sub UnsupportedFile()
    fraMenuPic.Visible = False
    fraFileInfo.Visible = False
    fraNoSupport.Visible = True
    fraFileInfo.Enabled = False
    cmdSaveGP2Info.Enabled = False
    
    mnuJamCheck2.Enabled = False
    mnuPCS.Enabled = False
    mnuBackup.Enabled = False
    mnuSetup.Enabled = False
    mnuToolUninstall.Enabled = False
    mnuTrackSettings.Enabled = False
    
    Toolbar1.Buttons(10).Enabled = False
    Toolbar1.Buttons(11).Enabled = False
    Toolbar1.Buttons(12).Enabled = False
End Sub

Private Sub SupportedFile()
    If FileType = 0 Then
        fraMenuPic.Visible = False
        fraNoSupport.Visible = False
        fraFileInfo.Enabled = True
        fraFileInfo.Visible = True
        cmdSaveGP2Info.Enabled = True

        mnuJamCheck2.Enabled = True
        mnuPCS.Enabled = True
        mnuBackup.Enabled = True
        mnuSetup.Enabled = True
        mnuTrackSettings.Enabled = True
        mnuToolUninstall.Enabled = True

        Toolbar1.Buttons(10).Enabled = True
        Toolbar1.Buttons(11).Enabled = True
        Toolbar1.Buttons(12).Enabled = True
    ElseIf (FileType = 1) Or (FileType = 2) Then
        fraFileInfo.Visible = False
        fraNoSupport.Visible = False
        fraFileInfo.Enabled = False
        cmdSaveGP2Info.Enabled = False
        fraMenuPic.Visible = True
        fraMenuPic.Enabled = True

        mnuTrackSettings.Enabled = False
        mnuJamCheck2.Enabled = False
        mnuPCS.Enabled = False
        mnuBackup.Enabled = False
        mnuSetup.Enabled = False
        
        Toolbar1.Buttons(10).Enabled = False
        Toolbar1.Buttons(11).Enabled = False
        Toolbar1.Buttons(12).Enabled = False
        
        
    ElseIf FileType = 4 Then
        fraNoSupport.Visible = True
        UnsupportedFile
    End If
End Sub

Private Sub ListFiles(ByVal Show As PatFile)
Dim MyName As String
Dim vArray As Variant
Dim MyExt As String

    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    'Remove all items from the listbox
    lstTmpFile.Clear
    lstFile.ListItems.Clear
    On Error GoTo ErrHandler
    Count1 = 0
    
    'list files and folders
    MyName = Dir(MyPath, vbDirectory)
    Do While MyName <> ""
        If MyName <> "." And MyName <> ".." Then
            If LCase(MyName) <> "pagefile.sys" Then
                MyExt = oFile.GetFilePart(MyName, GetExt)
                If ((GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory) And (Show = All) Then
                    lstFile.ListItems.Add , "dir" & MyPath & MyName, MyName, 1, 1
                ElseIf MyExt = ".dat" And ((Show = All) Or (Show = Dat)) Then
                    lstTmpFile.AddItem MyName
                    Count1 = Count1 + 1
                ElseIf (MyExt = ".bmp") And (Show = All) Then
                    lstTmpFile.AddItem MyName
                ElseIf (MyExt = ".gif") And (Show = All) Then
                    lstTmpFile.AddItem MyName
                End If
            End If
        End If
        MyName = Dir
    Loop
    'Sort folders
    lstFile.Sorted = True
    lstFile.Sorted = False

    If Count1 < 16 Then
        mnuRand.Enabled = False
    Else
        mnuRand.Enabled = True
    End If

    'List the files
    Dim ItemX
    If lstTmpFile.ListCount <> 0 Then
        For X = 0 To lstTmpFile.ListCount - 1
            lstTmpFile.ListIndex = X
            Read = MyPath & lstTmpFile.Text
            If oFile.GetFilePart(Read, GetExt) = ".dat" Then
                Set ItemX = lstFile.ListItems.Add(, Read, lstTmpFile.Text, 2, 2)
            ElseIf oFile.GetFilePart(Read, GetExt) = ".bmp" Then
                Set ItemX = lstFile.ListItems.Add(, Read, lstTmpFile.Text, 3, 3)
            ElseIf oFile.GetFilePart(Read, GetExt) = ".gif" Then
                Set ItemX = lstFile.ListItems.Add(, Read, lstTmpFile.Text, 3, 3)
            End If
            ItemX.SubItems(1) = Round(FileLen(Read) / 1000, 0) & " kb"
            ItemX.SubItems(2) = Mid(FileSystem.FileDateTime(Read), 1, 10)
        Next
    End If
    lblMyPath = MyPath
Exit Sub
ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: ListFiles()", vbCritical, TH & " - Error"
End Sub
