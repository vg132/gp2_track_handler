VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Gp2 Track Handler"
   ClientHeight    =   4230
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3195
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2880
      Top             =   0
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Viktor Gars 1998-2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   9
      Top             =   2760
      Width           =   3285
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "With help from"
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   3285
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Developed and Copyright"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   3285
   End
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "support@vgsoftware.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   705
      MouseIcon       =   "frmAbout.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3960
      Width           =   1845
   End
   Begin VB.Label lblINet 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "http://www.vgsoftware.com/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   585
      MouseIcon       =   "frmAbout.frx":045C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3720
      Width           =   2085
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "You have used this program XXX times."
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   3285
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Steven Young"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   3285
   End
   Begin VB.Label lblAbout 
      Caption         =   "Version 1.6.0"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2910
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "Gp2 Track Handler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblAbout 
      Caption         =   $"frmAbout.frx":05AE
      ForeColor       =   &H00000000&
      Height          =   1170
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   2910
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Settings", "Nr")
    lblAbout(3) = "You have used this program" & Str(X) & " times."
    X = 2
End Sub

Private Sub lblAbout_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    X = X + 1
    If X > 23 Then X = 0
    If X = 0 Then Read = "Steven Young"
    If X = 1 Then Read = "Bob Beeler"
    If X = 2 Then Read = "Ingo Serf"
    If X = 3 Then Read = "Crippen"
    If X = 4 Then Read = "Robert Kimber"
    If X = 5 Then Read = "Brett Knuchel"
    If X = 6 Then Read = "Per Eliasson"
    If X = 7 Then Read = "Bob Pearson"
    If X = 8 Then Read = "Fernando César"
    If X = 9 Then Read = "Jocelyn Coutu"
    If X = 10 Then Read = "John Slade"
    If X = 11 Then Read = "Richard M Comar"
    If X = 12 Then Read = "Rolph"
    If X = 13 Then Read = "Greg West"
    If X = 14 Then Read = "J Vennix"
    If X = 15 Then Read = "Graham D"
    If X = 16 Then Read = "Troy"
    If X = 17 Then Read = "Lee Armstrong"
    If X = 18 Then Read = "Marc Aarts"
    If X = 19 Then Read = "Willem van der Steen"
    If X = 20 Then Read = "Patricio Catalán M"
    If X = 21 Then Read = "Ricky Wakefield"
    If X = 22 Then Read = "Mal Ross"
    If X = 23 Then Read = "Gp2 Open Source Project"
    lblAbout(4).Caption = Read
End Sub

Private Sub lblEMail_Click()
    INetLink "mailto:support@vgsoftware.com", Me.hWnd
End Sub

Private Sub lblINet_Click()
    INetLink "http://www.vgsoftware.com/", Me.hWnd
End Sub
