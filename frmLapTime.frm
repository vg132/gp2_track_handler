VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLapTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Add New"
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   6000
      Width           =   1035
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   315
      Left            =   3870
      TabIndex        =   1
      Top             =   6000
      Width           =   1035
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Pos."
         Object.Width           =   441
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Qual/Race"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Track"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Driver"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Team"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Time"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmLapTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    If ListView1.SelectedItem.Selected = True Then
        Count2 = ListView1.SelectedItem.Index
        X = Mid(ListView1.SelectedItem.Key, 2, Len(ListView1.SelectedItem.Key) - 1)
        oMisc.WriteINI "Time", "T" & Trim(Str(X)), "|D|", prigramdir & "\LapTime.lda"
        LoadTime
        If ListView1.ListItems.Count > 0 Then
            ListView1.ListItems(Count2).Selected = True
        End If
    Else
        MsgBox "You have to select a record to delete it.", vbInformation, TH
    End If
End Sub

Private Sub cmdNew_Click()
    frmAddTime.Show vbModal, frmLapTime
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LoadTime
End Sub

Public Sub LoadTime()
Dim ItemX
Dim SF As String 'FileName
Dim Track As String
Dim Driver As String
Dim LTime As String
Dim Team As String
Dim LDate As String
    ListView1.ListItems.Clear
    Read = ""
    Read2 = ""
    Read3 = ""
    SF = ProgramDir & "\LapTime.lda"
    Y = oMisc.ReadINI("Time", "Nr", SF)
    Count1 = 0
    For X = 1 To Y
        Read = oMisc.ReadINI("Time", "T" & Trim(Str(X)), SF)
        If Read <> "|D|" Then
            Count1 = Count1 + 1
            Start = InStr(1, Read, "|Track|")
            Stopp = InStr(1, Read, "|Driver|")
            Track = Mid(Read, Start + 7, Stopp - Start - 7)
            Start = Stopp
            Stopp = InStr(1, Read, "|Team|")
            Driver = Mid(Read, Start + 8, Stopp - Start - 8)
            Start = Stopp
            Stopp = InStr(1, Read, "|Time|")
            Team = Mid(Read, Start + 6, Stopp - Start - 6)
            Start = Stopp
            Stopp = InStr(1, Read, "|Date|")
            LTime = Mid(Read, Start + 6, Stopp - Start - 6)
            Start = Stopp
            Stopp = InStr(1, Read, "|End|")
            LDate = Mid(Read, Start + 6, Stopp - Start - 6)
            Set ItemX = ListView1.ListItems.Add(, "k" & X, Count1)
            With ItemX
                .SubItems(1) = Mid(Read, 2, 4)
                .SubItems(2) = Track
                .SubItems(3) = Driver
                .SubItems(4) = Team
                .SubItems(5) = LTime
                .SubItems(6) = LDate
            End With
        End If
    Next
        '|Qual|Track|Monza|Driver|Viktor Gars|Team|Ferrari|Time|1:14.123|Date|1999-01-15|End|
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    frmMain.Caption = ColumnHeader.Width
End Sub
