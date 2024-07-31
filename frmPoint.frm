VERSION 5.00
Begin VB.Form frmPoint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Point Editor"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4125
   Icon            =   "frmPoint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPoint 
      Caption         =   "Point Editor (0-99)"
      ClipControls    =   0   'False
      Height          =   3255
      Left            =   60
      TabIndex        =   26
      Top             =   60
      Width           =   3975
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         Height          =   315
         Left            =   2760
         TabIndex        =   56
         Top             =   2400
         Width           =   1035
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   315
         Left            =   2760
         TabIndex        =   55
         Top             =   2760
         Width           =   1035
      End
      Begin VB.CommandButton cmdF1 
         Caption         =   "&Formula 1"
         Height          =   315
         Left            =   2760
         TabIndex        =   28
         Top             =   600
         Width           =   1035
      End
      Begin VB.CommandButton cmdCart 
         Caption         =   "&Cart"
         Height          =   315
         Left            =   2760
         TabIndex        =   27
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   25
         Top             =   2460
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   24
         Top             =   2145
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   23
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   23
         Top             =   1830
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   22
         Top             =   1515
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   21
         Top             =   1200
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   20
         Top             =   885
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   19
         Top             =   570
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   18
         Top             =   240
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2775
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   16
         Top             =   2460
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   13
         Top             =   1515
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   15
         Top             =   2145
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1830
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1200
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   11
         Top             =   885
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   10
         Top             =   570
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   400
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1515
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   400
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1200
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   400
         MaxLength       =   2
         TabIndex        =   2
         Top             =   885
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   400
         MaxLength       =   2
         TabIndex        =   1
         Top             =   570
         Width           =   400
      End
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
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   9
         Top             =   240
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   400
         MaxLength       =   2
         TabIndex        =   7
         Top             =   2460
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   400
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2145
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   400
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1830
         Width           =   400
      End
      Begin VB.TextBox txtPoint 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   400
         MaxLength       =   2
         TabIndex        =   8
         Top             =   2775
         Width           =   400
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "26th"
         Height          =   195
         Index           =   25
         Left            =   1800
         TabIndex        =   54
         Top             =   2505
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "17th"
         Height          =   195
         Index           =   16
         Left            =   900
         TabIndex        =   53
         Top             =   2505
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "24th"
         Height          =   195
         Index           =   23
         Left            =   1800
         TabIndex        =   52
         Top             =   1875
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "16th"
         Height          =   195
         Index           =   15
         Left            =   900
         TabIndex        =   51
         Top             =   2190
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "7th"
         Height          =   195
         Index           =   6
         Left            =   70
         TabIndex        =   50
         Top             =   2190
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "23th"
         Height          =   195
         Index           =   22
         Left            =   1800
         TabIndex        =   49
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "15th"
         Height          =   195
         Index           =   14
         Left            =   900
         TabIndex        =   48
         Top             =   1875
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "6th"
         Height          =   195
         Index           =   5
         Left            =   70
         TabIndex        =   47
         Top             =   1875
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "18th"
         Height          =   195
         Index           =   17
         Left            =   900
         TabIndex        =   46
         Top             =   2820
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "1st"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   300
         Width           =   210
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "10th"
         Height          =   195
         Index           =   9
         Left            =   900
         TabIndex        =   44
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "8th"
         Height          =   195
         Index           =   7
         Left            =   70
         TabIndex        =   43
         Top             =   2505
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "9th"
         Height          =   195
         Index           =   8
         Left            =   70
         TabIndex        =   42
         Top             =   2820
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "5th"
         Height          =   195
         Index           =   4
         Left            =   70
         TabIndex        =   41
         Top             =   1560
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "4th"
         Height          =   195
         Index           =   3
         Left            =   70
         TabIndex        =   40
         Top             =   1245
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "3th"
         Height          =   195
         Index           =   2
         Left            =   70
         TabIndex        =   39
         Top             =   930
         Width           =   225
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "2nd"
         Height          =   195
         Index           =   1
         Left            =   70
         TabIndex        =   38
         Top             =   615
         Width           =   270
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "20th"
         Height          =   195
         Index           =   19
         Left            =   1800
         TabIndex        =   37
         Top             =   615
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "19th"
         Height          =   195
         Index           =   18
         Left            =   1800
         TabIndex        =   36
         Top             =   300
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "14th"
         Height          =   195
         Index           =   13
         Left            =   900
         TabIndex        =   35
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "13th"
         Height          =   195
         Index           =   12
         Left            =   900
         TabIndex        =   34
         Top             =   1245
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "12th"
         Height          =   195
         Index           =   11
         Left            =   900
         TabIndex        =   33
         Top             =   930
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "11th"
         Height          =   195
         Index           =   10
         Left            =   900
         TabIndex        =   32
         Top             =   615
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "22th"
         Height          =   195
         Index           =   21
         Left            =   1800
         TabIndex        =   31
         Top             =   1245
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "25th"
         Height          =   195
         Index           =   24
         Left            =   1800
         TabIndex        =   30
         Top             =   2190
         Width           =   315
      End
      Begin VB.Label lblPoint 
         AutoSize        =   -1  'True
         Caption         =   "21st"
         Height          =   195
         Index           =   20
         Left            =   1800
         TabIndex        =   29
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
    SavePoint
    Unload Me
End Sub

Private Sub Form_Activate()
Dim Y As Integer
Dim X As Long
    For Y = 0 To 25
        X = GetWindowLong(txtPoint(Y).hwnd, GWL_STYLE)
        X = X Or ES_NUMBER
        Call SetWindowLong(txtPoint(Y).hwnd, GWL_STYLE, X)
    Next
    GetPoint
End Sub

Private Sub txtPoint_GotFocus(Index As Integer)
    TextSelected
End Sub
