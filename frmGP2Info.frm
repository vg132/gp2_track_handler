VERSION 5.00
Begin VB.Form frmGP2Info 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GP2 Track Info Editor"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   ControlBox      =   0   'False
   Icon            =   "frmGP2Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameData 
      Caption         =   "Track Data"
      Height          =   2175
      Left            =   120
      TabIndex        =   28
      Top             =   3360
      Width           =   2535
      Begin VB.TextBox txtLaps 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtLen 
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txtWare 
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Laps (3-126)"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Track Length (0-9999 m)"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Tyre ware (14848-37887)"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   1785
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5400
      TabIndex        =   14
      Top             =   3000
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   5400
      TabIndex        =   13
      Top             =   120
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   2760
      Pattern         =   "*.dat"
      TabIndex        =   12
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Misc Items"
      Height          =   2175
      Left            =   2760
      TabIndex        =   20
      Top             =   3360
      Width           =   5415
      Begin VB.TextBox txtMisc 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   5175
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtEvent 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtAuthor 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtSlot 
         Height          =   285
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Misc Text"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Event (e.g F1, Cart)"
         Height          =   195
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Author 
         AutoSize        =   -1  'True
         Caption         =   "Author"
         Height          =   195
         Left            =   3240
         TabIndex        =   22
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lblSlot 
         Caption         =   "Slot"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frameInfo 
      Caption         =   "Track Info"
      Height          =   2175
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   2535
      Begin VB.TextBox txtRLap 
         Height          =   285
         Left            =   1275
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1680
         Width           =   1140
      End
      Begin VB.TextBox txtQLap 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1680
         Width           =   1080
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtCountry 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Best Qual Lap"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Best Race Lap"
         Height          =   195
         Left            =   1275
         TabIndex        =   26
         Top             =   1440
         Width           =   1065
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Country:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   585
      End
      Begin VB.Label lblLabel 
         AutoSize        =   -1  'True
         Caption         =   "Track Name:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmGP2Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
