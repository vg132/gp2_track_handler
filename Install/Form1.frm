VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "GP2 Track Handler v1.5"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3120
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "&Install"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdUninstall 
      Caption         =   "&UnInstall"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SysPath As String

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub DrawBackGround()
    Const intBLUESTART% = 255
    Const intBLUEEND% = 0
    Const intBANDHEIGHT% = 50
    Const intSHADOWSTART% = 8
    Const intSHADOWCOLOR% = 0
    Const intTEXTSTART% = 4
    Const intTEXTCOLOR% = 15
    Const intRed% = 1
    Const intGreen% = 2
    Const intBlue% = 4
    Const intBackRed% = 8
    Const intBackGreen% = 16
    Const intBackBlue% = 32
    Dim sngBlueCur As Single
    Dim sngBlueStep As Single
    Dim intFormHeight As Integer
    Dim intFormWidth As Integer
    Dim intY As Integer
    Dim iColor As Integer
    Dim iRed As Single, iBlue As Single, iGreen As Single
    
    '
    'Get system values for height and width
    '
    intFormHeight = ScaleHeight
    intFormWidth = ScaleWidth

    iColor = intBlue
    'Calculate step size and blue start value
    '
    sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / intFormHeight
    sngBlueCur = intBLUESTART

    '
    'Paint blue screen
    '
    For intY = 0 To intFormHeight Step intBANDHEIGHT
        If iColor And intBlue Then iBlue = sngBlueCur
        If iColor And intRed Then iRed = sngBlueCur
        If iColor And intGreen Then iGreen = sngBlueCur
        If iColor And intBackBlue Then iBlue = 255 - sngBlueCur
        If iColor And intBackRed Then iRed = 255 - sngBlueCur
        If iColor And intBackGreen Then iGreen = 255 - sngBlueCur
        Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(iRed, iGreen, iBlue), BF
        sngBlueCur = sngBlueCur + sngBlueStep
    Next intY

    CurrentX = intSHADOWSTART
    CurrentY = intSHADOWSTART
    ForeColor = QBColor(intSHADOWCOLOR)
    Print "GP2 Track Handler v1.5"
    CurrentX = intTEXTSTART
    CurrentY = intTEXTSTART
    ForeColor = QBColor(intTEXTCOLOR)
    Print "GP2 Track Handler v1.5"
End Sub

Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdInstall_Click()
Dim iRet As Integer
Dim lSize As Long
Dim sBuffer As String
Dim RetVal As Long
Dim FileFound As Boolean
Dim AppDir As String
    AppDir = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\")
    
    iRet = 0
    If Not RegJad2Jam = ERROR_SUCCESS Then
        MsgBox "The installation program was unable to install Jad2Jam.ocx.", vbInformation
        iRet = 1
    End If
    If Not RegUpDown = ERROR_SUCCESS Then
        MsgBox "The installation program was unable to install UpDown.ocx.", vbInformation
        iRet = 1
    End If

    lSize = 256
    sBuffer = Space(256)
    RetVal = GetSystemDirectory(sBuffer, lSize)
    SysPath = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    FileFound = File_Exists(SysPath & "\TabCtl32.ocx")
    If FileFound = False Then
        FileCopy AppDir & "TabCtl32.ocx", SysPath & "\TabCtl32.ocx"
        FileCopy AppDir & "TabCtl32.ocx", AppDir & "TabCtl32.tmp"
        DoEvents
        Kill (AppDir & "TabCtl32.ocx")
        DoEvents
        If Not RegTabCtl32 = ERROR_SUCCESS Then
            MsgBox "The installation program was unable to install TabCtl32.ocx.", vbInformation
            iRet = 1
        End If
        FileCopy AppDir & "TabCtl32.tmp", AppDir & "TabCtl32.ocx"
        Kill (AppDir & "TabCtl32.tmp")
        DoEvents
    End If

    lSize = 256
    sBuffer = Space(256)
    RetVal = GetSystemDirectory(sBuffer, lSize)
    SysPath = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    FileFound = File_Exists(SysPath & "\ComCtl32.ocx")
    If FileFound = False Then
        FileCopy AppDir & "ComCtl32.ocx", SysPath & "\ComCtl32.ocx"
        FileCopy AppDir & "ComCtl32.ocx", AppDir & "ComCtl32.tmp"
        DoEvents
        Kill (AppDir & "ComCtl32.ocx")
        DoEvents
        If Not RegComCtl32 = ERROR_SUCCESS Then
            MsgBox "The installation program was unable to install ComCtl32.ocx.", vbInformation
            iRet = 1
        End If
        FileCopy AppDir & "ComCtl32.tmp", AppDir & "ComCtl32.ocx"
        Kill (AppDir & "ComCtl32.tmp")
    End If


    If iRet = 0 Then
        MsgBox "Setup was completed successfully.", vbInformation
    Else
        MsgBox "Setup was not completed successfully.", vbInformation
    End If
    End
End Sub

Private Sub cmdUninstall_Click()
Dim iRet As Byte
    If Not UnRegJad2Jam = ERROR_SUCCESS Then
        MsgBox "The installation program was unable to uninstall Jad2Jam.ocx.", vbInformation
        iRet = 1
    End If
    If Not UnRegUpDown = ERROR_SUCCESS Then
        MsgBox "The installation program was unable to uninstall UpDown.ocx.", vbInformation
        iRet = 1
    End If
    If iRet = 0 Then
        MsgBox "Setup was completed successfully.", vbInformation
    Else
        MsgBox "Setup was not completed successfully.", vbInformation
    End If
    End
End Sub

Private Sub Form_Resize()
    Me.Show
    DrawBackGround
    cmdInstall.Left = (Form1.Width / 2) - (cmdInstall.Width / 2)
    cmdInstall.Top = (Form1.Height / 2) - (cmdInstall.Height / 2) - 100 - (cmdInstall.Height / 2)
    cmdUninstall.Left = (Form1.Width / 2) - (cmdInstall.Width / 2)
    cmdUninstall.Top = (Form1.Height / 2) - (cmdInstall.Height / 2) + 100 + (cmdInstall.Height / 2)
    cmdClose.Top = Form1.Height - 1000
    cmdClose.Left = Form1.Width - 2000
End Sub
