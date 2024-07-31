VERSION 5.00
Begin VB.Form frmGP2Edit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GP2 Edit"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   ClipControls    =   0   'False
   Icon            =   "frmGP2Edit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   2700
      TabIndex        =   4
      Top             =   3180
      Width           =   1035
   End
   Begin VB.ListBox lstGP2Edit 
      Height          =   3435
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   60
      Width           =   2535
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   2700
      TabIndex        =   1
      Top             =   480
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2700
      TabIndex        =   0
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "* If you export menu pictures with gp2hipic then don't use this function, the pictures in gp2 will be destoyed."
      Height          =   1755
      Left            =   2700
      TabIndex        =   2
      Top             =   840
      Width           =   1035
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmGP2Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InLoad As Boolean

Private Sub cmdClose_Click()
    Var.sString1 = ""
    For Var.iInt1 = 0 To lstGP2Edit.ListCount - 1
        If lstGP2Edit.Selected(Var.iInt1) = False Then
            Var.sString1 = Var.sString1 & GetLetter(lstGP2Edit.ItemData(Var.iInt1))
        End If
    Next
    oMisc.WriteINI "Misc", "Exe", Var.sString1, TempFile
    FileInfo.Changes = True
End Sub

Private Sub cmdDelete_Click()
    oMisc.WriteINI "Misc", "Exe", "", TempFile
    oMisc.WriteINI "Misc", "ExePath", "", TempFile
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    GetEXEData
    lstGP2Edit.Selected(1) = True
End Sub

Private Sub GetEXEData()
Dim FileNum As Integer
Dim FileData As Byte
Dim X As Long
    Read = oMisc.ReadINI("Misc", "ExePath", TempFile)
    FileNum = FreeFile
    Open Read For Binary As FileNum
    Read = ""
    Read = oMisc.ReadINI("Misc", "EXE", TempFile)
    Read = Read & " "
    X = 0
    Get #FileNum, 49193, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Tram Data"
        lstGP2Edit.ItemData(X) = "0"
        Var.iInt1 = InStr(1, Read, GetLetter(0))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49194, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Cockpit Colors"
        lstGP2Edit.ItemData(X) = "1"
        Var.iInt1 = InStr(1, Read, GetLetter(1))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49195, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Pit Crew Colours"
        lstGP2Edit.ItemData(X) = "2"
        Var.iInt1 = InStr(1, Read, GetLetter(2))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49196, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Car Settings"
        lstGP2Edit.ItemData(X) = "3"
        Var.iInt1 = InStr(1, Read, GetLetter(3))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49197, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Damage Data"
        lstGP2Edit.ItemData(X) = "4"
        Var.iInt1 = InStr(1, Read, GetLetter(4))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49198, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Camera Data"
        lstGP2Edit.ItemData(X) = "5"
        Var.iInt1 = InStr(1, Read, GetLetter(5))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49199, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Game Settings"
        lstGP2Edit.ItemData(X) = "6"
        Var.iInt1 = InStr(1, Read, GetLetter(6))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49200, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Points Data"
        lstGP2Edit.ItemData(X) = "7"
        Var.iInt1 = InStr(1, Read, GetLetter(7))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49201, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Lap Data"
        lstGP2Edit.ItemData(X) = "8"
        Var.iInt1 = InStr(1, Read, GetLetter(8))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49153, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Car JAMs (" & FileData & ")"
        lstGP2Edit.ItemData(X) = "9"
        Var.iInt1 = InStr(1, Read, GetLetter(9))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49158, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Helmets JAMs"
        lstGP2Edit.ItemData(X) = "10"
        Var.iInt1 = InStr(1, Read, GetLetter(10))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49183, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "*Menu Helmets (" & FileData & ")"
        lstGP2Edit.ItemData(X) = "11"
        Var.iInt1 = InStr(1, Read, GetLetter(11))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
        Label1.Visible = True
    Else
        Label1.Visible = False
    End If

    Get #FileNum, 49173, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Cockpits (" & FileData & ")"
        lstGP2Edit.ItemData(X) = "12"
        Var.iInt1 = InStr(1, Read, GetLetter(12))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49163, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Wheel JAMs (" & FileData & ")"
        lstGP2Edit.ItemData(X) = "13"
        Var.iInt1 = InStr(1, Read, GetLetter(13))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49178, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "Sound Effects (" & FileData & ")"
        lstGP2Edit.ItemData(X) = "14"
        Var.iInt1 = InStr(1, Read, GetLetter(14))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49168, FileData
    If FileData <> 0 Then
        lstGP2Edit.AddItem "JAM Files (" & FileData & ")"
        lstGP2Edit.ItemData(X) = "15"
        Var.iInt1 = InStr(1, Read, GetLetter(15))
        If Var.iInt1 <> 0 Then
            lstGP2Edit.Selected(X) = False
        Else
            lstGP2Edit.Selected(X) = True
        End If
        X = X + 1
    End If
    Close FileNum
End Sub

Private Function GetLetter(ByVal Nr) As String
'*************************************
'Function Name: GetLetter
'Use:
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-08-24
'*************************************
    If Nr = 0 Then
        GetLetter = "TD "
    ElseIf Nr = 1 Then GetLetter = "CC "
    ElseIf Nr = 2 Then GetLetter = "PC "
    ElseIf Nr = 3 Then GetLetter = "CS "
    ElseIf Nr = 4 Then GetLetter = "DD "
    ElseIf Nr = 5 Then GetLetter = "CD "
    ElseIf Nr = 6 Then GetLetter = "GS "
    ElseIf Nr = 7 Then GetLetter = "PD "
    ElseIf Nr = 8 Then GetLetter = "LD "
    ElseIf Nr = 9 Then GetLetter = "CJ "
    ElseIf Nr = 10 Then GetLetter = "HJ "
    ElseIf Nr = 11 Then GetLetter = "MH "
    ElseIf Nr = 12 Then GetLetter = "CP "
    ElseIf Nr = 13 Then GetLetter = "WJ "
    ElseIf Nr = 14 Then GetLetter = "SE "
    ElseIf Nr = 15 Then GetLetter = "JF "
    End If
End Function
