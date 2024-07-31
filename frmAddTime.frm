VERSION 5.00
Begin VB.Form frmAddTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Time"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmAddTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraQual 
      Caption         =   "Lap Time Data"
      Height          =   2475
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   4215
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   315
         Left            =   3060
         TabIndex        =   6
         Top             =   2040
         Width           =   1035
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1035
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtTrack 
         Height          =   285
         Left            =   2160
         MaxLength       =   30
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtTime 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtDriver 
         Height          =   285
         Left            =   2160
         MaxLength       =   23
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtTeam 
         Height          =   285
         Left            =   120
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Qual/Race Lap"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Track Name"
         Height          =   255
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Driver"
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Team"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Date (e.g 1999-06-19)"
         Height          =   195
         Left            =   2160
         TabIndex        =   9
         Top             =   1440
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmAddTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dott As Boolean

Private Sub txtTime_Change()
    If Dott = True Then
        If Len(txtTime.Text) = 1 Then
            txtTime.Text = txtTime.Text & ":"
        ElseIf Len(txtTime.Text) = 4 Then
            txtTime.Text = txtTime.Text & "."
        End If
        SendKeys ("^{END}")
    End If
End Sub

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode <> 46) And (KeyCode <> 8) Then
        Dott = True
    Else
        Dott = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Read = txtTime.Text & ";" & txtDate.Text & ";" & cboType.Text & ";" & txtDriver.Text & ";" & txtTeam.Text & ";" & txtTrack.Text
    oDB.SaveNew dbFile, Read
    Unload Me
End Sub

Private Sub Form_Load()
    cboType.AddItem "Qual"
    cboType.AddItem "Race"
    cboType.ListIndex = 0

    X = GetWindowLong(txtTime.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtTime.hWnd, GWL_STYLE, X)

    X = GetWindowLong(txtDate.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtDate.hWnd, GWL_STYLE, X)
End Sub

Private Sub txtDate_Change()
    If Dott = True Then
        If Len(txtDate.Text) = 4 Then
            txtDate.Text = txtDate.Text & "-"
        ElseIf Len(txtDate.Text) = 7 Then
            txtDate.Text = txtDate.Text & "-"
        End If
        SendKeys ("^{END}")
    End If
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode <> 46) And (KeyCode <> 8) Then
        Dott = True
    Else
        Dott = False
    End If
End Sub

Private Sub txtDate_GotFocus()
    TextSelected
End Sub

Private Sub txtDriver_GotFocus()
    TextSelected
End Sub

Private Sub txtTeam_GotFocus()
    TextSelected
End Sub

Private Sub txtTime_GotFocus()
    TextSelected
End Sub

Private Sub txtTrack_GotFocus()
    TextSelected
End Sub
