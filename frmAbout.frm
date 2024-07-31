VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About GP2 Track Handler"
   ClientHeight    =   3660
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2526.197
   ScaleMode       =   0  'User
   ScaleWidth      =   4394.762
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   0
   End
   Begin VB.Label lblEMail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "formula1@swipnet.se"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2355
      MouseIcon       =   "frmAbout.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3360
      Width           =   1545
   End
   Begin VB.Label lblINet 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "http://hem1.passagen.se/formula1/"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1845
      MouseIcon       =   "frmAbout.frx":045C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3000
      Width           =   2565
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3450
      Left            =   120
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label lblAbout 
      Caption         =   "You have used this program XXX times."
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      Caption         =   "Special Thanks to"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   2910
   End
   Begin VB.Label lblAbout 
      Caption         =   "Version 1.4"
      Height          =   225
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   2910
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "GP2 Track Handler"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2910
   End
   Begin VB.Label lblAbout 
      Caption         =   $"frmAbout.frx":05AE
      ForeColor       =   &H00000000&
      Height          =   1170
      Index           =   1
      Left            =   1680
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
    Image1.Picture = LoadResPicture(115, 0)
    X = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Settings", "Nr")
    lblAbout(3) = "You have used this program" & Str(X) & " times"
    X = 1
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub lblAbout_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    X = X + 1
    If X > 24 Then X = 1
    If X = 1 Then Read = "Special Thanks to"
    If X = 2 Then Read = "Steven Young"
    If X = 3 Then Read = "Bob Beeler"
    If X = 4 Then Read = "Ingo Serf"
    If X = 5 Then Read = "Crippen"
    If X = 6 Then Read = "Robert Kimber"
    If X = 7 Then Read = "Brett Knuchel"
    If X = 8 Then Read = "Beta Testers"
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
    If X = 22 Then Read = "Willem van der Steen"
    If X = 23 Then Read = "Patricio Catalán M"
    If X = 24 Then Read = "Ricky Wakefield"
    lblAbout(4).Caption = Read
End Sub

Private Sub lblEMail_Click()
    Read = "mailto:formula1@swipnet.se"
    OpenPage = ShellExecute(frmAbout.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub

Private Sub lblINet_Click()
    Read = "http://hem1.passagen.se/formula1/"
    OpenPage = ShellExecute(frmAbout.hwnd, "open", Read, vbNullString, vbNullString, 1)
End Sub
