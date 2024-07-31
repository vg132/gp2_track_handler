VERSION 5.00
Begin VB.Form frmCCSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Track Settings"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmCCSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFuel 
      Caption         =   "Fuel Consumption"
      Height          =   735
      Left            =   60
      TabIndex        =   30
      Top             =   2880
      Width           =   2175
      Begin VB.TextBox txtFuel 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   33
         Text            =   "6874"
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Fuel Consumption:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   3560
      TabIndex        =   24
      Top             =   3720
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   60
      TabIndex        =   25
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Frame fraWing 
      Caption         =   "Wings"
      Height          =   1000
      Left            =   60
      TabIndex        =   29
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtRWing 
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "10"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtFWing 
         Height          =   285
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "11"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Rear Wing:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Front Wing:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tyre Compound"
      Height          =   735
      Left            =   2300
      TabIndex        =   28
      Top             =   2880
      Width           =   2295
      Begin VB.ComboBox cboTire 
         Height          =   315
         ItemData        =   "frmCCSetup.frx":030A
         Left            =   1440
         List            =   "frmCCSetup.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tyre Compound:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gears"
      Height          =   2715
      Left            =   2300
      TabIndex        =   27
      Top             =   120
      Width           =   2295
      Begin VB.TextBox txt2 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "42"
         Top             =   705
         Width           =   615
      End
      Begin VB.TextBox txt5 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "58"
         Top             =   1785
         Width           =   615
      End
      Begin VB.TextBox txt6 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "64"
         Top             =   2145
         Width           =   615
      End
      Begin VB.TextBox txt3 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "48"
         Top             =   1065
         Width           =   615
      End
      Begin VB.TextBox txt4 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   17
         Text            =   "53"
         Top             =   1425
         Width           =   615
      End
      Begin VB.TextBox txt1 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "38"
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "2'nd Gear:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "5'th Gear:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "6th Gear:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "4'th Gear:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "3'rd Gear:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "1'st Gear:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Psysics"
      Height          =   1700
      Left            =   60
      TabIndex        =   26
      Top             =   1150
      Width           =   2175
      Begin VB.TextBox txtAir 
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "64"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtBrack 
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "64"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtGrip 
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "64"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtAcce 
         Height          =   285
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "64"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Track Grip:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Brackbalance:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Air Resistance:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Acceleration:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCCSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path As String
Dim Path2 As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Read2 = Mid(Read, 1, 1)
    Count1 = txtFWing.Text + 151
    Read2 = Read2 & Chr(Count1)
    Count1 = txtRWing.Text + 151
    Read2 = Read2 & Chr(Count1)
    Count1 = txt1.Text + 151
    Read2 = Read2 & Chr(Count1)
    Count1 = txt2.Text + 151
    Read2 = Read2 & Chr(Count1)
    Count1 = txt3.Text + 151
    Read2 = Read2 & Chr(Count1)
    Count1 = txt4.Text + 151
    Read2 = Read2 & Chr(Count1)
    Count1 = txt5.Text + 151
    Read2 = Read2 & Chr(Count1)
    Count1 = txt6.Text + 151
    Read2 = Read2 & Chr(Count1)
    Count1 = cboTire.ListIndex + 52
    Read2 = Read2 & Chr(Count1)
    Read2 = Read2 & Mid(Read, 11, 2)
    Read2 = Read2 & Chr(txtGrip.Text) & Mid(Read, 14, 1)
    Read2 = Read2 & Chr(txtBrack.Text) & Mid(Read, 16, 5)
    Read2 = Read2 & Chr(txtAcce.Text) & Mid(Read, 22, 1)
    Read2 = Read2 & Chr(txtAir.Text) & Mid(Read, 24, 2)
    SaveTrackSetup Path, Read2, txtFuel.Text

    Read = ""
    Read = oMisc.File_Exists(ProgramDir & "\Bat\Setup.bat")
    If Read = True Then Kill (ProgramDir & "\Bat\Setup.bat")
    Path2 = frmMain.lstFile.SelectedItem.Key

    FileNum = FreeFile
    Open ProgramDir & "\Bat\Setup.bat" For Append As FileNum
    Read = oMisc.GetShortName(frmMain.Dir1.Path)
    Read2 = oMisc.GetShortName(Path2)
    For X = Len(Read2) To 0 Step -1
        If Mid(Read2, X, 1) = "\" Then Exit For
    Next
    Read2 = Mid(Read2, X + 1)
    Print #FileNum, "@echo off"
    Print #FileNum, "cd " & Read
    Print #FileNum, Mid(Read, 1, 2)
    Print #FileNum, "thcheck " & Read2
    Read = ""
    Read = oMisc.File_Exists(frmMain.Dir1.Path & "\thcheck.exe")
    If Read = False Then
        If Len(frmMain.Dir1.Path) = 3 Then
            FileCopy ProgramDir & "\gp2utils\thcheck.exe", frmMain.Dir1.Path & "thcheck.exe"
            Print #FileNum, "del thcheck.exe"
        Else
            FileCopy ProgramDir & "\gp2utils\thcheck.exe", frmMain.Dir1.Path & "\thcheck.exe"
            Print #FileNum, "del thcheck.exe"
        End If
    End If
    Print #FileNum, "cls"
    Print #FileNum, "echo You can close this window now."
    Close FileNum
    RetVal = ShellExecute(frmMain.hwnd, "open", ProgramDir & "\Bat\Setup.bat", vbNullString, vbNullString, 1)
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    SetNr
    If frmMain.tabMain.Tab = 0 Then
        Path = frmMain.lstFile.SelectedItem.Key
    ElseIf frmMain.tabMain.Tab = 1 Then
        Path = frmMain.txtPath.Text
    Else
        Exit Sub
    End If
    Read = GetTrackSetup(Path)
    If Read <> "" Then
        frmCCSetup.Caption = "Advanced Track Settings [" & frmMain.lstFile.SelectedItem.Text & "]"
        cboTire.ListIndex = Asc(Mid(Read, 10, 1)) - 52
    
        txtFWing = Asc(Mid(Read, 2, 1)) - 151
        txtRWing = Asc(Mid(Read, 3, 1)) - 151
    
        txt1 = Asc(Mid(Read, 4, 1)) - 151
        txt2 = Asc(Mid(Read, 5, 1)) - 151
        txt3 = Asc(Mid(Read, 6, 1)) - 151
        txt4 = Asc(Mid(Read, 7, 1)) - 151
        txt5 = Asc(Mid(Read, 8, 1)) - 151
        txt6 = Asc(Mid(Read, 9, 1)) - 151
        
        txtAcce = Asc(Mid(Read, 21, 1))
        txtBrack = Asc(Mid(Read, 15, 1))
        txtGrip = Asc(Mid(Read, 13, 1))
        txtAir = Asc(Mid(Read, 23, 1))
        txtFuel = Mid(Read, 26)
    Else
        MsgBox LoadResString(127), vbInformation, TH
        Unload Me
    End If
Exit Sub
ErrTrap:
    Select Case Err.Number
    Case 380
        MsgBox LoadResString(128), vbInformation, TH
        Unload Me
    Case Else
        Print #Log, Date & " " & Time & " frmCCSetup_Load, Error Number: " & Err.Number & ", Error Description: " & Err.Description
        MsgBox "Error Number: " & Err.Number & vbCrLf & _
            "Error Description: " & Err.Description & vbCrLf & _
            "Error Source: " & Err.Source, vbCritical, "Error"
    End Select
End Sub

Private Sub txt1_GotFocus()
    TextSelected
End Sub

Private Sub txt2_GotFocus()
    TextSelected
End Sub

Private Sub txt3_GotFocus()
    TextSelected
End Sub

Private Sub txt4_GotFocus()
    TextSelected
End Sub

Private Sub txt5_GotFocus()
    TextSelected
End Sub

Private Sub txt6_GotFocus()
    TextSelected
End Sub

Private Sub txtAcce_GotFocus()
    TextSelected
End Sub

Private Sub txtAcce_LostFocus()
    If txtAcce.Text = "" Then txtAcce.Text = 0
    If txtAcce.Text > 100 Then txtAcce.Text = 100
    If txtAcce.Text < 0 Then txtAcce.Text = 0
End Sub

Private Sub txtAir_GotFocus()
    TextSelected
End Sub

Private Sub txtAir_LostFocus()
    If txtAir = "" Then txtAir = 0
    If txtAir > 100 Then txtAir = 100
    If txtAir < 0 Then txtAir = 0
End Sub

Private Sub txtBrack_GotFocus()
    TextSelected
End Sub

Private Sub txtBrack_LostFocus()
    If txtBrack.Text = "" Then txtBrack.Text = 0
    If txtBrack.Text > 100 Then txtBrack.Text = 100
    If txtBrack.Text < 0 Then txtBrack.Text = 0
End Sub

Private Sub txtFuel_LostFocus()
    If txtFuel = "" Then txtFuel = "0"
    If txtFuel > "32767" Then txtFuel = "32767"
    If txtFuel < "0" Then txtFuel = "0"
End Sub

Private Sub txtFWing_GotFocus()
    TextSelected
End Sub

Private Sub txtFuel_GotFocus()
    TextSelected
End Sub

Private Sub txtFWing_LostFocus()
    If txtFWing.Text = "" Then txtFWing = 0
    If txtFWing > 20 Then txtFWing = 20
    If txtFWing < 0 Then txtFWing = 0
End Sub

Private Sub txtGrip_GotFocus()
    TextSelected
End Sub

Private Sub txtGrip_LostFocus()
    If txtGrip.Text = "" Then txtGrip.Text = 100
    If txtGrip.Text > 100 Then txtGrip.Text = 100
    If txtGrip.Text < 0 Then txtGrip.Text = 0
End Sub

Private Sub txtRWing_GotFocus()
    TextSelected
End Sub

Private Sub txtRWing_LostFocus()
    If txtRWing.Text = "" Then txtRWing = 0
    If txtRWing > 20 Then txtRWing = 20
    If txtRWing < 0 Then txtRWing = 0
End Sub

Public Sub SetNr()
Dim oCtl As Control
    For Each oCtl In frmCCSetup.Controls
        If TypeOf oCtl Is TextBox Then
                X = GetWindowLong(oCtl.hwnd, GWL_STYLE)
                X = X Or ES_NUMBER
                Call SetWindowLong(oCtl.hwnd, GWL_STYLE, X)
        End If
    Next
    Set oCtl = Nothing
End Sub
