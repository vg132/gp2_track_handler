VERSION 5.00
Begin VB.Form frmPoint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Point Editor"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "frmPoint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCart 
      Caption         =   "&Cart"
      Height          =   315
      Left            =   2920
      TabIndex        =   31
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdF1 
      Caption         =   "&Formula 1"
      Height          =   315
      Left            =   2920
      TabIndex        =   30
      Top             =   480
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   2920
      TabIndex        =   29
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2920
      TabIndex        =   28
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Frame fraPoint 
      Caption         =   "Point Editor (0-99)"
      ClipControls    =   0   'False
      Height          =   3255
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   400
         MaxLength       =   2
         TabIndex        =   0
         Top             =   255
         Width           =   400
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "26th"
         Height          =   195
         Index           =   25
         Left            =   1800
         TabIndex        =   27
         Top             =   2505
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "17th"
         Height          =   195
         Index           =   16
         Left            =   900
         TabIndex        =   26
         Top             =   2505
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "24th"
         Height          =   195
         Index           =   23
         Left            =   1800
         TabIndex        =   25
         Top             =   1875
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "16th"
         Height          =   195
         Index           =   15
         Left            =   900
         TabIndex        =   24
         Top             =   2190
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "7th"
         Height          =   195
         Index           =   6
         Left            =   70
         TabIndex        =   23
         Top             =   2190
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "23th"
         Height          =   195
         Index           =   22
         Left            =   1800
         TabIndex        =   22
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "15th"
         Height          =   195
         Index           =   14
         Left            =   900
         TabIndex        =   21
         Top             =   1875
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "6th"
         Height          =   195
         Index           =   5
         Left            =   70
         TabIndex        =   20
         Top             =   1875
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "18th"
         Height          =   195
         Index           =   17
         Left            =   900
         TabIndex        =   19
         Top             =   2820
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "1st"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   210
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "10th"
         Height          =   195
         Index           =   9
         Left            =   900
         TabIndex        =   17
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "8th"
         Height          =   195
         Index           =   7
         Left            =   70
         TabIndex        =   16
         Top             =   2505
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "9th"
         Height          =   195
         Index           =   8
         Left            =   70
         TabIndex        =   15
         Top             =   2820
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "5th"
         Height          =   195
         Index           =   4
         Left            =   70
         TabIndex        =   14
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "4th"
         Height          =   195
         Index           =   3
         Left            =   70
         TabIndex        =   13
         Top             =   1245
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "3th"
         Height          =   195
         Index           =   2
         Left            =   70
         TabIndex        =   12
         Top             =   930
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "2nd"
         Height          =   195
         Index           =   1
         Left            =   70
         TabIndex        =   11
         Top             =   615
         Width           =   270
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "20th"
         Height          =   195
         Index           =   19
         Left            =   1800
         TabIndex        =   10
         Top             =   615
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "19th"
         Height          =   195
         Index           =   18
         Left            =   1800
         TabIndex        =   9
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "14th"
         Height          =   195
         Index           =   13
         Left            =   900
         TabIndex        =   8
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "13th"
         Height          =   195
         Index           =   12
         Left            =   900
         TabIndex        =   7
         Top             =   1245
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "12th"
         Height          =   195
         Index           =   11
         Left            =   900
         TabIndex        =   6
         Top             =   930
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "11th"
         Height          =   195
         Index           =   10
         Left            =   900
         TabIndex        =   5
         Top             =   615
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "22th"
         Height          =   195
         Index           =   21
         Left            =   1800
         TabIndex        =   4
         Top             =   1245
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "25th"
         Height          =   195
         Index           =   24
         Left            =   1800
         TabIndex        =   3
         Top             =   2190
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "21st"
         Height          =   195
         Index           =   20
         Left            =   1800
         TabIndex        =   2
         Top             =   930
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtLaps_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdF1_Click()
    FileInfo.Changes = True
    txtPoint(0).Text = "10"
    txtPoint(1).Text = "6"
    txtPoint(2).Text = "4"
    txtPoint(3).Text = "3"
    txtPoint(4).Text = "2"
    txtPoint(5).Text = "1"
    For X = 6 To 25
        txtPoint(X).Text = "0"
    Next
End Sub

Private Sub cmdCart_Click()
    FileInfo.Changes = True
    txtPoint(0).Text = "20"
    txtPoint(1).Text = "16"
    txtPoint(2).Text = "14"
    txtPoint(3).Text = "12"
    txtPoint(4).Text = "10"
    txtPoint(5).Text = "8"
    txtPoint(6).Text = "6"
    txtPoint(7).Text = "5"
    txtPoint(8).Text = "4"
    txtPoint(9).Text = "3"
    txtPoint(10).Text = "2"
    txtPoint(11).Text = "1"
    X = 12
    For X = 12 To 25
        txtPoint(X).Text = "0"
    Next
End Sub

Private Sub cmdSave_Click()
    FileInfo.Changes = True
    SavePoint
End Sub

Private Sub Form_Activate()
    For Var.iInt1 = 0 To 25
        Var.lLong1 = GetWindowLong(txtPoint(Var.iInt1).hwnd, GWL_STYLE)
        Var.lLong1 = Var.lLong1 Or ES_NUMBER
        Call SetWindowLong(txtPoint(Var.iInt1).hwnd, GWL_STYLE, Var.lLong1)
    Next
    GetPoint
End Sub

Private Sub Form_Load()
    CreateTextBox
End Sub

Private Sub txtPoint_GotFocus(Index As Integer)
    TextSelected
End Sub

Private Sub CreateTextBox()
'*************************************
'Function Name: CreateTextBox
'Use: The creates text boxes in runtime
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-08-29
'*************************************
    For X = 1 To 8
        Load txtPoint(X)
        txtPoint(X).Visible = True
        txtPoint(X).Top = txtPoint(X - 1).Top + 315
        txtPoint(X).TabIndex = X
    Next
    For X = 9 To 17
        Load txtPoint(X)
        txtPoint(X).Visible = True
        txtPoint(X).Left = 1320
        txtPoint(X).Top = ((X - 9) * 315) + 255
        txtPoint(X).TabIndex = X
    Next
    For X = 18 To 25
        Load txtPoint(X)
        txtPoint(X).Visible = True
        txtPoint(X).Left = 2160
        txtPoint(X).Top = ((X - 18) * 315) + 255
        txtPoint(X).TabIndex = X
    Next
End Sub
