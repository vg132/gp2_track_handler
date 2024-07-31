VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmGP2Edit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GP2 Edit"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   ClipControls    =   0   'False
   Icon            =   "frmGP2Edit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView lstGP2Edit 
      Height          =   4335
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   7646
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Data"
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   4440
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "* If you export menu pictures with gp2hipic then don't use this function, the pictures in gp2 will be destoyed."
      Height          =   795
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   2355
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
Dim R As Long
    Read = ""
    For X = 0 To lstGP2Edit.ListItems.Count - 1
        R = SendMessageLong(lstGP2Edit.hwnd, LVM_GETITEMSTATE, X, LVIS_STATEIMAGEMASK)
        If R <> 8192 Then
            Read = Read & lstGP2Edit.ListItems(X + 1).Key
        End If
    Next
    oMisc.WriteINI "Misc", "Exe", Read, TempFile
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    oMisc.WriteINI "Misc", "Exe", "", TempFile
    oMisc.WriteINI "Misc", "ExePath", "", TempFile
    Unload Me
End Sub

Private Sub Form_Load()
    Call SendMessageLong(lstGP2Edit.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_CHECKBOXES, -1)
    Call SendMessageLong(lstGP2Edit.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, -1)
    Call SendMessageLong(lstGP2Edit.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, -1)
    GetEXEData
End Sub

Private Sub GetEXEData()
Dim FileNum As Integer
Dim FileData As Byte
Dim X As Long
    Read = oMisc.ReadINI("Misc", "ExePath", TempFile)
    FileNum = FreeFile
    Open Read For Binary As FileNum
    Read = ""
    Read = oMisc.ReadINI("Misc", "Exe", TempFile)

    Get #FileNum, 49193, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "TD ", "Team Data"
        X = InStr(1, Read, "TD")
    End If

    Get #FileNum, 49194, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "CC ", "Cockpit Colours"
        X = InStr(1, Read, "CC")
    End If

    Get #FileNum, 49195, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "PC ", "Pit Crew Colours"
        X = InStr(1, Read, "PC")
    End If

    Get #FileNum, 49196, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "CS ", "Car Settings"
        X = InStr(1, Read, "CS")
    End If

    Get #FileNum, 49197, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "DD ", "Damage Data"
        X = InStr(1, Read, "DD")
    End If

    Get #FileNum, 49198, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "CD ", "Camera Data"
        X = InStr(1, Read, "CD")
    End If

    Get #FileNum, 49199, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "GS ", "Game Settings"
        X = InStr(1, Read, "GS")
    End If

    Get #FileNum, 49200, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "PD ", "Points Data"
        X = InStr(1, Read, "PD")
    End If

    Get #FileNum, 49201, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "LD ", "Laps Data"
        X = InStr(1, Read, "LD")
    End If

    Get #FileNum, 49153, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "CJ ", "Car JAMs (" & FileData & ")"
        X = InStr(1, Read, "CJ")
    End If

    Get #FileNum, 49158, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "HJ ", "Helmets JAMs"
        X = InStr(1, Read, "HJ")
    End If

    Get #FileNum, 49183, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "MH ", "*Menu Helmets (" & FileData & ")"
        X = InStr(1, Read, "MH")
        frmGP2Edit.Height = 5985
    Else
        frmGP2Edit.Height = 5160
    End If

    Get #FileNum, 49173, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "CP ", "Cockpits (" & FileData & ")"
        X = InStr(1, Read, "CP")
    End If

    Get #FileNum, 49163, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "WJ ", "Wheel JAMs (" & FileData & ")"
        X = InStr(1, Read, "WJ")
    End If

    Get #FileNum, 49178, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "SE ", "Sound Effects (" & FileData & ")"
        X = InStr(1, Read, "SE")
    End If

    Get #FileNum, 49168, FileData
    If FileData <> 0 Then
        lstGP2Edit.ListItems.Add , "JF ", "JAM Files (" & FileData & ")"
        X = InStr(1, Read, "JF")
    End If
    Close FileNum
End Sub
