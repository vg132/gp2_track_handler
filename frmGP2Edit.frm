VERSION 5.00
Begin VB.Form frmGP2Edit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gp2Edit"
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
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   2700
      TabIndex        =   4
      Top             =   3180
      Width           =   1035
   End
   Begin VB.ListBox lstGp2Edit 
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
   Begin VB.CommandButton cmdSave 
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
  Unload Me
End Sub

Private Sub cmdDelete_Click()
    WriteINI "Misc", "Exe", "", TempFile
    WriteINI "Misc", "ExePath", "", TempFile
    Unload Me
End Sub

Private Sub cmdSave_Click()
  Read = ""
  For tVar.iInt = 0 To lstGp2Edit.ListCount - 1
      If lstGp2Edit.Selected(tVar.iInt) = False Then
          Read = Read & GetLetter(lstGp2Edit.ItemData(tVar.iInt))
      End If
  Next
  WriteINI "Misc", "Exe", Read, TempFile
End Sub

Private Sub Form_Load()
  GetEXEData
  lstGp2Edit.Selected(1) = True
End Sub

Private Sub GetEXEData()
Dim FileNum As Integer
Dim FileData As Byte
Dim X As Long
    Read = ReadINI("Misc", "ExePath", TempFile)
    FileNum = FreeFile
    Open Read For Binary As FileNum
    Read = ""
    Read = ReadINI("Misc", "EXE", TempFile)
    Read = Read & " "
    X = 0
    Get #FileNum, 49193, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Tram Data"
        lstGp2Edit.ItemData(X) = "0"
        If InStr(1, Read, GetLetter(0)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49194, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Cockpit Colors"
        lstGp2Edit.ItemData(X) = "1"
        If InStr(1, Read, GetLetter(1)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49195, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Pit Crew Colours"
        lstGp2Edit.ItemData(X) = "2"
        If InStr(1, Read, GetLetter(2)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49196, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Car Settings"
        lstGp2Edit.ItemData(X) = "3"
        If InStr(1, Read, GetLetter(3)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49197, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Damage Data"
        lstGp2Edit.ItemData(X) = "4"
        If InStr(1, Read, GetLetter(4)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49198, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Camera Data"
        lstGp2Edit.ItemData(X) = "5"
        If InStr(1, Read, GetLetter(5)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49199, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Game Settings"
        lstGp2Edit.ItemData(X) = "6"
        If InStr(1, Read, GetLetter(6)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49200, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Points Data"
        lstGp2Edit.ItemData(X) = "7"
        If InStr(1, Read, GetLetter(7)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49201, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Lap Data"
        lstGp2Edit.ItemData(X) = "8"
        If InStr(1, Read, GetLetter(8)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49153, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Car JAMs (" & FileData & ")"
        lstGp2Edit.ItemData(X) = "9"
        If InStr(1, Read, GetLetter(9)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49158, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Helmets JAMs"
        lstGp2Edit.ItemData(X) = "10"
        If InStr(1, Read, GetLetter(10)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49183, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "*Menu Helmets (" & FileData & ")"
        lstGp2Edit.ItemData(X) = "11"
        If InStr(1, Read, GetLetter(11)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
        Label1.Visible = True
    Else
        Label1.Visible = False
    End If

    Get #FileNum, 49173, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Cockpits (" & FileData & ")"
        lstGp2Edit.ItemData(X) = "12"
        If InStr(1, Read, GetLetter(12)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49163, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Wheel JAMs (" & FileData & ")"
        lstGp2Edit.ItemData(X) = "13"
        If InStr(1, Read, GetLetter(13)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49178, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "Sound Effects (" & FileData & ")"
        lstGp2Edit.ItemData(X) = "14"
        If InStr(1, Read, GetLetter(14)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
        End If
        X = X + 1
    End If

    Get #FileNum, 49168, FileData
    If FileData <> 0 Then
        lstGp2Edit.AddItem "JAM Files (" & FileData & ")"
        lstGp2Edit.ItemData(X) = "15"
        If InStr(1, Read, GetLetter(15)) <> 0 Then
            lstGp2Edit.Selected(X) = False
        Else
            lstGp2Edit.Selected(X) = True
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
