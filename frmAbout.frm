VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About GP2 Track Handler v1.4"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5490
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   840
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   5535
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   4155
      TabIndex        =   1
      Top             =   2280
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   150
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":030A
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   165
      TabIndex        =   6
      Top             =   1680
      Width           =   3870
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "GP2 Track Handler"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   840
      TabIndex        =   5
      Top             =   0
      Width           =   4125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Viktor Gars"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2355
      TabIndex        =   4
      Top             =   720
      Width           =   1170
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Programme Development:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   870
      TabIndex        =   3
      Top             =   480
      Width           =   4185
   End
   Begin VB.Label lblShow 
      Alignment       =   2  'Center
      Caption         =   "Special Thanks to:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   2370
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload frmAbout
    MDIForm1.Show
End Sub

Private Sub Form_Load()
    X = 1
End Sub

Private Sub Timer1_Timer()
    X = X + 1
    If X > 21 Then X = 1
    If X = 1 Then Read = "Special Thanks to:"
    If X = 2 Then Read = "Steven Young"
    If X = 3 Then Read = "Bob Beeler"
    If X = 4 Then Read = "Ingo Serf"
    If X = 5 Then Read = "Crippen"
    If X = 6 Then Read = "Robert Kimber"
    If X = 7 Then Read = "Brett Knuchel"
    If X = 8 Then Read = "Beta Testers:"
    If X = 9 Then Read = "Per Eliasson"
    If X = 10 Then Read = "Bob Pearson"
    If X = 11 Then Read = "Fernando César"
    If X = 12 Then Read = "Jocelyn Coutu"
    If X = 13 Then Read = "John Slade"
    If X = 14 Then Read = "Richard M Comar"
    If X = 15 Then Read = "Rolph"
    If X = 16 Then Read = "Greg West"
    If X = 17 Then Read = "J Vennix"
    If X = 18 Then Read = "Graham D"
    If X = 19 Then Read = "Troy"
    If X = 20 Then Read = "Lee Armstrong"
    If X = 21 Then Read = "Marc Aarts"
    lblShow.Caption = Read
End Sub


