VERSION 5.00
Begin VB.Form frmChangeData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   3420
      TabIndex        =   2
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   3420
      TabIndex        =   3
      Top             =   600
      Width           =   1035
   End
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label lblText 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   3255
   End
End
Attribute VB_Name = "frmChangeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function ShowMsg(ByVal Title As String, ByVal Default As String, ByVal Nr As Boolean, Optional txtLen As Integer)
    If Nr = True Then
        X = GetWindowLong(txtText.hwnd, GWL_STYLE)
        X = X Or ES_NUMBER
        Call SetWindowLong(txtText.hwnd, GWL_STYLE, X)
    End If
    frmChangeData.Top = frmChangeData.Top - 1000
    frmChangeData.Caption = Title
    lblText.Caption = Title
    txtText.Text = Default
    txtText.MaxLength = txtLen
End Function

Private Sub cmdCancel_Click()
    Read = "PostGoranErrorKalle"
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Read = txtText.Text
    Unload Me
End Sub

Private Sub txtText_GotFocus()
    TextSelected
End Sub
