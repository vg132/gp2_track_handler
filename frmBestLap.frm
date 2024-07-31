VERSION 5.00
Begin VB.Form frmBestLap 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BestLap"
   ClientHeight    =   4485
   ClientLeft      =   1095
   ClientTop       =   285
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   325
      Left            =   1920
      TabIndex        =   11
      Top             =   4080
      Width           =   1000
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   325
      Left            =   600
      TabIndex        =   9
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Insert"
      Height          =   325
      Left            =   1920
      TabIndex        =   10
      Top             =   3600
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   325
      Left            =   3240
      TabIndex        =   12
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   325
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Prev Record"
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   325
      Left            =   4320
      TabIndex        =   13
      ToolTipText     =   "Next Record"
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   8
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   7
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   7
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   6
      Left            =   240
      MaxLength       =   10
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   5
      Left            =   2640
      MaxLength       =   12
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   4
      Left            =   240
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   2640
      MaxLength       =   22
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   2
      Left            =   240
      MaxLength       =   22
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   240
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Race"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   25
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Qual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   24
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Track:"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Date: (e.g 1998-11-29)"
      Height          =   195
      Index           =   7
      Left            =   2640
      TabIndex        =   22
      Top             =   2280
      Width           =   1605
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Date: (e.g 1998-11-29)"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   21
      Top             =   2280
      Width           =   1605
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Team:"
      Height          =   195
      Index           =   5
      Left            =   2640
      TabIndex        =   20
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Team:"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Driver:"
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   18
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Driver:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   17
      Top             =   1080
      Width           =   465
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Best Lap: (e.g 1,24,254)"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   16
      Top             =   2880
      Width           =   1710
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Best Lap: (e.g 1,24,254)"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   1710
   End
End
Attribute VB_Name = "frmBestLap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload frmBestLap
    MDIForm1.Show
    Close DataBaseFileNum
    If CurrentRecord3 > 1 Then SaveCurrentRecord2
End Sub

Private Sub cmdClose_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        X = X + 1
        Do Until X > LastRecord2 + 1
            Get #DataBaseFileNum, X, TimeBase
            Read2 = Trim(TimeBase.TName)
            If UCase(Trim(Read2)) = UCase(Read) Then
                SaveCurrentRecord2
                CurrentRecord3 = X
                ShowCurrentRecord2
                Exit Sub
            End If
            X = X + 1
        Loop
        MsgBox "The Track was not found in this database.", vbInformation, "Find"
    End If
End Sub

Private Sub cmdDelete_Click()
    Count1 = CurrentRecord3
    X = 1
    Count2 = 1
    FileNum = FreeFile
    RecordLen = Len(TimeBase)
    Open ProgramDir + "\$tmp$.$tm" For Random As FileNum Len = RecordLen
    Do Until X > LastRecord2
        If X <> Count1 Then
            Get #DataBaseFileNum, X, TimeBase
            Put #FileNum, Count2, TimeBase
            Count2 = Count2 + 1
        End If
        X = X + 1
    Loop
    Close FileNum
    Close DataBaseFileNum
    DeleteFile ProgramDir + "\database.tdb"
    SourceFile = ProgramDir + "\$tmp$.$tm"
    TargetFile = ProgramDir + "\database.tdb"
    FileCopy SourceFile, TargetFile
    DeleteFile SourceFile

    DataBaseFileNum = FreeFile
    RecordLen = Len(TimeBase)
    Open ProgramDir + "\database.tdb" For Random As DataBaseFileNum Len = RecordLen
    LastRecord2 = FileLen(ProgramDir + "\database.tdb") / RecordLen
    If CurrentRecord3 > LastRecord2 Then CurrentRecord3 = LastRecord2
    If CurrentRecord > 0 Then ShowCurrentRecord2
End Sub

Private Sub cmdDelete_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        X = X + 1
        Do Until X > LastRecord2 + 1
            Get #DataBaseFileNum, X, TimeBase
            Read2 = Trim(TimeBase.TName)
            If UCase(Trim(Read2)) = UCase(Read) Then
                SaveCurrentRecord2
                CurrentRecord3 = X
                ShowCurrentRecord2
                Exit Sub
            End If
            X = X + 1
        Loop
        MsgBox "The Track was not found in this database.", vbInformation, "Find"
    End If
End Sub

Private Sub cmdExport_Click()
    MDIForm1.txtQDate = txtFields(6)
    MDIForm1.txtRDate = txtFields(7)
    MDIForm1.txtQDriver = txtFields(2)
    MDIForm1.txtRDriver = txtFields(3)
    MDIForm1.txtQTime = txtFields(0)
    MDIForm1.txtRTime = txtFields(1)
    MDIForm1.txtQTeam = txtFields(4)
    MDIForm1.txtRTeam = txtFields(5)
End Sub

Public Sub SaveCurrentRecord2()
    If CurrentRecord3 <> 1 Then
        TimeBase.QDate = txtFields(6).Text
        TimeBase.RDate = txtFields(7).Text
        TimeBase.QDriver = txtFields(2).Text
        TimeBase.RDriver = txtFields(3).Text
        TimeBase.QTeam = txtFields(4).Text
        TimeBase.RTeam = txtFields(5).Text
        TimeBase.QTime = txtFields(0).Text
        TimeBase.RTime = txtFields(1).Text
        TimeBase.TName = txtFields(8).Text
        Put #DataBaseFileNum, CurrentRecord3, TimeBase
    End If
End Sub

Public Sub ShowCurrentRecord2()
    If LastRecord2 > 0 Then
        Get #DataBaseFileNum, CurrentRecord3, TimeBase
        txtFields(6).Text = Trim(TimeBase.QDate)
        txtFields(7).Text = Trim(TimeBase.RDate)
        txtFields(2).Text = Trim(TimeBase.QDriver)
        txtFields(3).Text = Trim(TimeBase.RDriver)
        txtFields(4).Text = Trim(TimeBase.QTeam)
        txtFields(5).Text = Trim(TimeBase.RTeam)
        txtFields(0).Text = Trim(TimeBase.QTime)
        txtFields(1).Text = Trim(TimeBase.RTime)
        txtFields(8).Text = Trim(TimeBase.TName)
        frmBestLap.Caption = "Record" + Str(CurrentRecord3) + "/" + Str(LastRecord2)
    End If
End Sub

Private Sub cmdExport_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        X = X + 1
        Do Until X > LastRecord2 + 1
            Get #DataBaseFileNum, X, TimeBase
            Read2 = Trim(TimeBase.TName)
            If UCase(Trim(Read2)) = UCase(Read) Then
                SaveCurrentRecord2
                CurrentRecord3 = X
                ShowCurrentRecord2
                Exit Sub
            End If
            X = X + 1
        Loop
        MsgBox "The Track was not found in this database.", vbInformation, "Find"
    End If
End Sub

Private Sub cmdFind_Click()
    Read = InputBox("Track Name (F3=Find Next)", "Find", Read)
    X = 1
    Do Until X > LastRecord2 + 1
        Get #DataBaseFileNum, X, TimeBase
        Read2 = Trim(TimeBase.TName)
        If UCase(Trim(Read2)) = UCase(Read) Then
            SaveCurrentRecord2
            CurrentRecord3 = X
            ShowCurrentRecord2
            Exit Sub
        End If
        X = X + 1
    Loop
    MsgBox "The Track was not found in this database.", vbInformation, "Find"
End Sub

Private Sub cmdFind_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        X = X + 1
        Do Until X > LastRecord2 + 1
            Get #DataBaseFileNum, X, TimeBase
            Read2 = Trim(TimeBase.TName)
            If UCase(Trim(Read2)) = UCase(Read) Then
                SaveCurrentRecord2
                CurrentRecord3 = X
                ShowCurrentRecord2
                Exit Sub
            End If
            X = X + 1
        Loop
        MsgBox "The Track was not found in this database.", vbInformation, "Find"
    End If
End Sub

Private Sub cmdNext_Click()
    If CurrentRecord3 = LastRecord2 Then
        Exit Sub
    Else
        SaveCurrentRecord2
        CurrentRecord3 = CurrentRecord3 + 1
        ShowCurrentRecord2
    End If
End Sub

Private Sub cmdNext_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        X = X + 1
        Do Until X > LastRecord2 + 1
            Get #DataBaseFileNum, X, TimeBase
            Read2 = Trim(TimeBase.TName)
            If UCase(Trim(Read2)) = UCase(Read) Then
                SaveCurrentRecord2
                CurrentRecord3 = X
                ShowCurrentRecord2
                Exit Sub
            End If
            X = X + 1
        Loop
        MsgBox "The Track was not found in this database.", vbInformation, "Find"
    End If
End Sub

Private Sub cmdPrev_Click()
    If CurrentRecord3 < 2 Then
        Exit Sub
    Else
        SaveCurrentRecord2
        CurrentRecord3 = CurrentRecord3 - 1
        ShowCurrentRecord2
    End If
End Sub

Private Sub cmdPrev_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        X = X + 1
        Do Until X > LastRecord2 + 1
            Get #DataBaseFileNum, X, TimeBase
            Read2 = Trim(TimeBase.TName)
            If UCase(Trim(Read2)) = UCase(Read) Then
                SaveCurrentRecord2
                CurrentRecord3 = X
                ShowCurrentRecord2
                Exit Sub
            End If
            X = X + 1
        Loop
        MsgBox "The Track was not found in this database.", vbInformation, "Find"
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        X = X + 1
        Do Until X > LastRecord2 + 1
            Get #DataBaseFileNum, X, TimeBase
            Read2 = Trim(TimeBase.TName)
            If UCase(Trim(Read2)) = UCase(Read) Then
                SaveCurrentRecord2
                CurrentRecord3 = X
                ShowCurrentRecord2
                Exit Sub
            End If
            X = X + 1
        Loop
        MsgBox "The Track was not found in this database.", vbInformation, "Find"
    End If
End Sub

Private Sub Form_Load()
    DataBaseFileNum = FreeFile
    RecordLen = Len(TimeBase)
    Open ProgramDir + "\database.tdb" For Random As DataBaseFileNum Len = RecordLen
    LastRecord2 = FileLen(ProgramDir + "\database.tdb") / RecordLen
    CurrentRecord3 = 1
    If CurrentRecord3 > LastRecord2 Then CurrentRecord3 = LastRecord2
    If CurrentRecord > 0 Then ShowCurrentRecord2
    Read = ""
End Sub

Private Sub txtFields_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        X = X + 1
        Do Until X > LastRecord2 + 1
            Get #DataBaseFileNum, X, TimeBase
            Read2 = Trim(TimeBase.TName)
            If UCase(Trim(Read2)) = UCase(Read) Then
                SaveCurrentRecord2
                CurrentRecord3 = X
                ShowCurrentRecord2
                Exit Sub
            End If
            X = X + 1
        Loop
        MsgBox "The Track was not found in this database.", vbInformation, "Find"
    End If
End Sub
