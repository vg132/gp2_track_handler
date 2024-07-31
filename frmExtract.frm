VERSION 5.00
Begin VB.Form frmExtract 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extract Backup File"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   Icon            =   "frmExtract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3255
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox FileList 
      Height          =   2400
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   2520
      Width           =   1035
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Extract"
      Default         =   -1  'True
      Height          =   315
      Left            =   2160
      TabIndex        =   0
      Top             =   2520
      Width           =   1035
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExtract_Click()
    Var.sString1 = oMisc.BrowseFolders("Extract to:", Me.hWnd)
    frmExtract.MousePointer = 11
    frmMain.Extract.Files.SelectNone
    frmMain.Extract.Files.SelectAll
    frmMain.Extract.Extract Var.sString1
    frmExtract.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call SendMessageLong(FileList.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, -1)
    Prepare
End Sub

Public Function Prepare() As Boolean
 Dim i As Long
    On Error GoTo BadCabinet
    'frmMain.Extract.FileName = frmExtract.Caption
    For i = 1 To frmMain.Extract.Files.Count
        FileList.AddItem frmMain.Extract.Files(i).FileName
    Next i
    Prepare = True
    On Error GoTo 0
    Exit Function
BadCabinet:
    MsgBox "There is a problem with this cabinet.  It may be damaged.", vbCritical, "Cabinet error"
    Unload Me
    Exit Function
End Function
