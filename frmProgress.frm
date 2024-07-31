VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar prbBackup 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Min             =   1e-4
   End
   Begin VB.Label lblText 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblFileName 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

