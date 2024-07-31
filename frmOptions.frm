VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "GP2 Location"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3690
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPath 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   325
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   325
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   325
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GP2 Path (e.g ""c:\games\gp2"")"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   2265
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Sub cmdCancel_Click()
    Read = GetSetting(GP2TH, "Settings", "GP2 Path")
    If Read = "" Then
        Unload MDIForm1
        Unload frmOptions
        Unload frmPoint
        Unload MDIForm1
        Unload MDIForm1
        End
    End If
    Unload frmOptions
    MDIForm1.Show
    Exit Sub
End Sub

Private Sub cmdOk_Click()
    On Error GoTo ErrorTrap
    If txtPath.Text = "" Then Exit Sub
    Read2 = txtPath.Text
    If Len(Read2) = 3 Then Read2 = Mid(Read2, 1, 2)
    Read2 = Read2 + "\gp2.exe"
    Read = oMisc.File_Exists(Read2)
    If Read = False Then
        Responce = MsgBox(TH + " was not able to find your GP2.EXE file, please check your GP2 directory.", vbRetryCancel + vbCritical, TH)
        If Responce = vbCancel Then
            Read = GetSetting(GP2TH, "Settings", "GP2 Path")
            If Read = "" Then
                Unload MDIForm1
                Unload frmOptions
                Unload frmPoint
                Unload MDIForm1
                Unload MDIForm1
                End
            Else
                Unload frmOptions
            End If
        End If
        Exit Sub
    End If
    SelectVersion = False
    Gp2Dir = txtPath
    CountNr = Len(Gp2Dir)
    If Mid(Gp2Dir, CountNr, 1) = "\" Then Gp2Dir = Mid(Gp2Dir, 1, CountNr - 1)
    SaveSetting GP2TH, "Settings", "GP2 Path", Gp2Dir
    GetGP2Version
    Unload frmOptions
    MDIForm1.Show
    MDIForm1.StatusBar1.Panels(2) = "GP2 Version: " + GP2Country
    MDIForm1.StatusBar1.Panels(3) = "GP2 Directory: " + Gp2Dir
    
    Exit Sub
ErrorTrap:
    Select Case Err.Number
    Case "76"
        MsgBox "Path not found, use the Browse Function (click on the Browse button) for best result.", vbInformation, TH
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Private Sub cmdBrowse_Click()
    szTitle = "Select GP2 Location"
    With tBrowseInfo
        .hWndOwner = Me.hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        txtPath = sBuffer
    End If
End Sub

Private Sub Form_Load()
    txtPath.Text = GetSetting(GP2TH, "Settings", "GP2 Path")
End Sub
