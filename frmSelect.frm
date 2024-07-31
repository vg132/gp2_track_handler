VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select GP2 Version"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3135
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Select your Version"
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   325
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   325
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   325
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Supported GP2 Versions"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2220
   End
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdHelp_Click()
    MsgBox "To see what version of GP2 you have start GP2 and select Options in the main menu, then select About Grand Prix 2 and there you will see what version of GP2 you use.", vbInformation, TH
End Sub

Private Sub cmdOk_Click()
    Read = Combo1.Text
    If Read = "US English Version 1.0b" Then
        GP2V = US
    ElseIf Read = "UK English Version 1.0b" Then
        GP2V = UK
    ElseIf Read = "Spanish Version 1.0b" Then
        GP2V = Sp
    ElseIf Read = "Dutch Version 1.0b" Then
        GP2V = NL
    ElseIf Read = "French Version 1.0b" Then
        GP2V = FR
    ElseIf Read = "Italian Version 1.0b" Then
        GP2V = IT
    ElseIf Read = "German Version 1.0b" Then
        GP2V = TY
    Else
        End
    End If
    SelectVersion = True
    Unload frmSelect
End Sub

Private Sub Form_Load()
    Combo1.AddItem "UK English Version 1.0b"
    Combo1.AddItem "US English Version 1.0b"
    Combo1.AddItem "Dutch Version 1.0b"
    Combo1.AddItem "French Version 1.0b"
    Combo1.AddItem "German Version 1.0b"
    Combo1.AddItem "Italian Version 1.0b"
    Combo1.AddItem "Spanish Version 1.0b"
    CountNr = 0
End Sub
