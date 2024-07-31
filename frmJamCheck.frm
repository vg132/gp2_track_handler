VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmJamCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jam Check"
   ClientHeight    =   4230
   ClientLeft      =   7455
   ClientTop       =   1890
   ClientWidth     =   5340
   Icon            =   "frmJamCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   315
      Left            =   2130
      TabIndex        =   1
      Top             =   3840
      Width           =   1035
   End
   Begin ComctlLib.ListView lstJam 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      Icons           =   "imlJam"
      SmallIcons      =   "imlJam"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Jam File"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Found/Not Found"
         Object.Width           =   2593
      EndProperty
   End
   Begin ComctlLib.ImageList imlJam 
      Left            =   360
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   16
      MaskColor       =   12566463
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmJamCheck.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmJamCheck.frx":065C
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmJamCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call SendMessageLong(lstJam.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, -1)
    Call SendMessageLong(lstJam.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, -1)
    If frmMain.tabMain.Tab = 0 Then
        frmJamCheck.Caption = "Jam Check - " & frmMain.lstFile.SelectedItem.Text
    Else
        frmJamCheck.Caption = "Jam Check - " & frmMain.txtPath.Text
    End If
End Sub
