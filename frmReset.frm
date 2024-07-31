VERSION 5.00
Begin VB.Form frmReset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reset Records"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   Icon            =   "frmReset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   2805
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   1035
   End
   Begin VB.TextBox txtDriver 
      Height          =   285
      Left            =   1200
      MaxLength       =   22
      TabIndex        =   0
      Text            =   "Geoff Crammond"
      Top             =   120
      Width           =   1515
   End
   Begin VB.TextBox txtTeam 
      Height          =   285
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   1
      Text            =   "Microprose"
      Top             =   480
      Width           =   1515
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "3:00.000"
      Top             =   840
      Width           =   1515
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1200
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "Driver:"
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Team:"
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Lap Time:"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Dott As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ResetTime txtDriver, txtTeam, txtTime, txtDate
    Unload Me
End Sub

Private Sub Form_Load()
    txtDate.Text = Date
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
