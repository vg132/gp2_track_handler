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
    If frmMain.tabMain.Tabs(1).Selected = True Then
        frmJamCheck.Caption = "Jam Check - " & frmMain.lstFile.SelectedItem.Text
        CheckJam frmMain.lstFile.SelectedItem.Key
    Else
        frmJamCheck.Caption = "Jam Check - " & frmMain.txtPath.Text
        CheckJam frmMain.txtPath.Text
    End If
End Sub

Public Sub CheckJam(ByVal Track As String)
Dim vArray As Variant
Dim vFound As Variant
Dim iFound As Integer

    lstJams.AddItem "Jam Check"
    lstJams.AddItem ""
    vArray = oFile.GetJamFiles(Track)
    iFound = 0
    ReDim vFound(0, 0 To UBound(vArray, 2))
    For X = 0 To UBound(vArray, 2)
        If vArray(0, X) = Empty Then Exit For
        Read = oFile.FileExists(GP2Dir & "\" & vArray(0, X))
        If Read = True Then
            vFound(0, iFound) = "Found          " & vArray(0, X)
            iFound = iFound + 1
        Else
            lstJams.AddItem "Not Found!  " & vArray(0, X)
        End If
    Next
    For X = 0 To UBound(vFound, 2)
        If vFound(0, X) = "" Then Exit For
        lstJams.AddItem vFound(0, X)
    Next

    If X = UBound(vArray, 2) + 1 Then
        lstJams.List(0) = "All " & UBound(vArray, 2) + 1 & " JamFiles was found in " & GP2Dir
    Else
        lstJams.List(0) = UBound(vArray, 2) + 1 - X & " of " & UBound(vArray, 2) + 1 & " JamFiles was not found in " & GP2Dir
    End If
End Sub
