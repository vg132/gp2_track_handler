VERSION 5.00
Begin VB.Form frmJamCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jam Check"
   ClientHeight    =   4095
   ClientLeft      =   7455
   ClientTop       =   1890
   ClientWidth     =   4950
   Icon            =   "frmJamCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstJams 
      Height          =   3570
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3840
      TabIndex        =   0
      Top             =   3720
      Width           =   1035
   End
End
Attribute VB_Name = "frmJamCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If frmMain.tabMain.Tab = 0 Then
        frmJamCheck.Caption = "Jam Check - " & frmMain.lstFile.SelectedItem.Text
        CheckJam frmMain.lstFile.SelectedItem.Key
    Else
        frmJamCheck.Caption = "Jam Check - " & frmMain.txtPath.Text
        CheckJam frmMain.txtPath.Text
    End If
End Sub

Public Sub CheckJam(ByVal Track As String)
Dim ItemX
Dim vArray As Variant
Dim iFound As Integer

    Var.lLong1 = 0
    FileNum = FreeFile
    Open Track For Binary As FileNum
    Read = String(2500, " ")
    X = FileLen(Track) - 2504
    Get #FileNum, X, Read
    Close FileNum
    Read2 = String(2, Chr(0))
    For X = 2500 To 1 Step -1
        If Mid(Read, X, 2) = Read2 Then Exit For
    Next
    
    Read = Mid(Read, X + 2)
    ReDim vArray(0, Asc(Mid(Read, 1, 1)))
    Read = Mid(Read, 3)
    lstJams.AddItem "    "
    lstJams.AddItem "    "
    Count1 = 0
    iFound = 0
    Do Until Len(Read) < 5
        Count1 = Count1 + 1
        Stopp = InStr(1, UCase(Read), UCase(Chr(0)))
        If Stopp = 0 Then
            Read2 = Read
        Else
            Read2 = Mid(Read, 1, Stopp - 1)
        End If
        Read3 = oMisc.File_Exists(GP2Dir & "\" & Read2)
        If Read3 = False Then
            lstJams.AddItem "Not Found!    " & Read2
            Var.lLong1 = Var.lLong1 + 1
        Else
            vArray(0, iFound) = "Found.           " & Read2
            iFound = iFound + 1
        End If
        If Stopp = 0 Then
            Read = 0
        Else
            Read = Mid(Read, Stopp + 1)
        End If
    Loop

    If Var.lLong1 = 0 Then
        lstJams.List(0) = "All " & Count1 & " Jam files used by this track were found."
    Else
        lstJams.List(0) = "Failed!! " & Var.lLong1 & " of the " & Count1 & " Jam files was not found!"
    End If

    For Count1 = 0 To iFound
        lstJams.AddItem vArray(0, Count1)
    Next
    lstJams.Selected(0) = True
End Sub
