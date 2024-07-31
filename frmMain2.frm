VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12405
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblFull 
      AutoSize        =   -1  'True
      Caption         =   "FullView Picture"
      Height          =   195
      Left            =   4560
      TabIndex        =   0
      Top             =   240
      Width           =   1125
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A_Click(Index As Integer)
    If A(Index).Tag = "On" Then
        A(Index).Picture = On1(Index).Picture
        A(Index).Tag = "Off"
    Else
        A(Index).Picture = Off(Index).Picture
        A(Index).Tag = "On"
    End If
    FileSaved = False
End Sub
Private Sub AC_Click(Index As Integer)
    If AC(Index).Tag = "On" Then
        AC(Index).Picture = On1(Index).Picture
        AC(Index).Tag = "Off"
    Else
        AC(Index).Picture = Off(Index).Picture
        AC(Index).Tag = "On"
    End If
    FileSaved = False
End Sub

Private Sub chk0as1_Click()
    FileSaved = False
End Sub

Private Sub chkSave_Click()
    FileSaved = False
End Sub

Private Sub cmdDefaultSettings_Click()
    hscPitSpeed.Value = 50
    hscPQPower.Value = 790
    hscPRPower.Value = 780
    hscWeight.Value = 1313
    HScroll1.Value = 1313
    HScroll2.Value = 5
    hscPGrip.Value = 198
    GP2AidsSet
    chkNoLimit.Value = 0
    chkSave.Value = 0
    chk0as1.Value = 1
    Slider1.Value = 1994
    chkSelectedTeam.Value = 0
    FileSaved = False
End Sub

Private Sub cmdExportSettings_Click()
    SavePlayerData
    If HScroll2.Enabled = True Then
        F1SaveFileNum = FreeFile
        Open Gp2Dir + "\F1gstate.sav" For Binary As F1SaveFileNum
            ExportQuickRace
        Close F1SaveFileNum
    End If
    GP2FileNum = FreeFile
    Open Gp2Dir + "\GP2.exe" For Binary As GP2FileNum
        ExportNullAsOne
        ExportLevel
        ExportSaveLap
        ExportCarHelp
        ExportPQPower
        ExportPQPower
        ExportPRPower
        ExportPGrip
        ExportPWeight
        ExportCWeight
        ExportSpeed
        ExportUseTeam
    Close GP2FileNum
    If HScroll2.Enabled = True Then
        DeleteFile Gp2Dir + "\$$Check$.bat"
        SourceFile = ProgramDir + "\gp2utils\check.exe"
        TargetFile = Gp2Dir + "\$$check.exe"
        FileCopy SourceFile, TargetFile
        FileNum = FreeFile
        Open Gp2Dir + "\$$Check$.bat" For Append As FileNum
        Print #FileNum, Mid(Gp2Dir, 1, 2)
        Print #FileNum, "cd " + Gp2Dir
        Print #FileNum, Gp2Dir + "\$$Check f1gstate.sav"
        Print #FileNum, "del $$Check.exe"
        Close FileNum
        Read = File_Exists("c:\command.com")
        Dim RetVal
        If Read = True Then
            ChDir Gp2Dir
            RetVal = Shell("c:\command.com /c " + Gp2Dir + "\$$Check$.bat", vbNormalFocus)
        Else
            ChDir Gp2Dir
            RetVal = Shell(Gp2Dir + "\$$Check$.bat", vbNormalFocus)
        End If
    End If
End Sub

Private Sub cmdImportSettings_Click()
    SavePlayerData
    GP2FileNum = FreeFile
    Open Gp2Dir + "\GP2.exe" For Binary As GP2FileNum
        ImportNullAsOne
        ImportLevel
        ImportSaveLap
        ImportGameSettings
        ImportPQPower
        ImportPRPower
        ImportPGrip
        ImportSpeed
        ImportCWeight
        ImportPWeight
        ImportUseTeam
    Close GP2FileNum
    If HScroll2.Enabled = True Then
        F1SaveFileNum = FreeFile
        Open Gp2Dir + "\F1gstate.sav" For Binary As F1SaveFileNum
            ImportQuick
        Close F1SaveFileNum
    End If
    FileSaved = False
    GetPlayerData
End Sub

Private Sub Command1_Click()
    MsgBox "The time must be writen in this formate #,##,###  and the date must be writen like this, 1999-01-24. You my not enter a date 'lower' then 1978-01-01 and not a time higher the 9,59,999.", vbInformation, TH
End Sub

Private Sub lblYear2_Click()
    On Error GoTo ErrorTrap
    Read = InputBox("Year (1900-2099):", "Select Year")
    If Read = "" Then Exit Sub
    If (Read > 1899) And (Read < 3000) Then Slider1.Value = Read
ErrorTrap:
    Exit Sub
End Sub

Private Sub P_Click(Index As Integer)
    If P(Index).Tag = "On" Then
        P(Index).Picture = On1(Index).Picture
        P(Index).Tag = "Off"
    Else
        P(Index).Picture = Off(Index).Picture
        P(Index).Tag = "On"
    End If
    FileSaved = False
End Sub

Private Sub R_Click(Index As Integer)
    If R(Index).Tag = "On" Then
        R(Index).Picture = On1(Index).Picture
        R(Index).Tag = "Off"
    Else
        R(Index).Picture = Off(Index).Picture
        R(Index).Tag = "On"
    End If
    FileSaved = False
End Sub

Private Sub S_Click(Index As Integer)
    If S(Index).Tag = "On" Then
        S(Index).Picture = On1(Index).Picture
        S(Index).Tag = "Off"
    Else
        S(Index).Picture = Off(Index).Picture
        S(Index).Tag = "On"
    End If
    FileSaved = False
End Sub

Private Sub chkNoLimit_Click()
    If chkNoLimit.Value = 1 Then
        hscPitSpeed.Enabled = False
    Else
        hscPitSpeed.Enabled = True
    End If
    FileSaved = False
End Sub

Private Sub chkSelectedTeam_Click()
    If chkSelectedTeam.Value = 1 Then
        hscPQPower.Enabled = False
        hscPRPower.Enabled = False
    Else
        hscPQPower.Enabled = True
        hscPRPower.Enabled = True
    End If
    FileSaved = False
End Sub

Private Sub cmdAdd_Click()
    DataBaseFileNum = FreeFile
    RecordLen = Len(TimeBase)
    Open ProgramDir + "\database.tdb" For Random As DataBaseFileNum Len = RecordLen
    LastRecord2 = FileLen(ProgramDir + "\database.tdb") / RecordLen
    LastRecord2 = LastRecord2 + 1
    TimeBase.TName = frmMain.txtName
    TimeBase.QTime = frmMain.txtQTime
    TimeBase.RTime = frmMain.txtRTime
    TimeBase.QDriver = frmMain.txtQDriver
    TimeBase.RDriver = frmMain.txtRDriver
    TimeBase.QTeam = frmMain.txtQTeam
    TimeBase.RTeam = frmMain.txtRTeam
    TimeBase.QDate = frmMain.txtQDate
    TimeBase.RDate = frmMain.txtRDate
    TimeBase.TName = frmMain.txtName
    Put #DataBaseFileNum, LastRecord2, TimeBase
    Close DataBaseFileNum
End Sub

Private Sub cmdBrowse_Click()
    X17 = 32000
    On Error GoTo ErrorTrap
    If TrackPath <> "" Then
        CommonDialog1.InitDir = TrackPath
    Else
        CommonDialog1.InitDir = DefaultTrackPath
    End If
    CommonDialog1.Filter = "GP2 Track Files (*.dat)|*.dat|All Files (*.*)|*.*|"
    CommonDialog1.ShowOpen
    Read3 = CommonDialog1.FileName
    TrackPath = Read3
    ReadTrackFile (CommonDialog1.FileName)
    If NoSupport = True Then Exit Sub

    Read = frmMain.lblCountry
    GetAdjectiv
    frmMain.txtAdjectiv = Read
    frmMain.txtCountry = frmMain.lblCountry
    frmMain.txtLaps = frmMain.lblLaps
    frmMain.txtLength = frmMain.lblLength
    frmMain.txtName = frmMain.lblName
    frmMain.txtPath = CommonDialog1.FileName
    frmMain.txtTire = frmMain.lblWare
    frmMain.txtQTime = frmMain.lblQLap
    frmMain.txtRTime = frmMain.lblRLap

    X17 = 1

    X = Mid(Form1.TreeView1.SelectedItem.Key, 2, 2) - 10
    Form1.TreeView1.SelectedItem.Text = Trim(Str(X)) + ". " + frmMain.txtName

    
    Dim nodX As Node    ' Create variable.
    Read = Form1.TreeView1.SelectedItem.Key
    If Form1.TreeView1.SelectedItem.Children > 0 Then
        If Form1.TreeView1.SelectedItem.Child.Key = "t" + Trim(Str(X + 100)) Then
            Form1.TreeView1.Nodes.Remove (Form1.TreeView1.SelectedItem.Child.Index)
        End If
        If Form1.TreeView1.SelectedItem.Children = 2 Then
            If Form1.TreeView1.SelectedItem.Child.Next.Key = "t" + Trim(Str(X + 100)) Then
                Form1.TreeView1.Nodes.Remove (Form1.TreeView1.SelectedItem.Child.Next.Index)
            End If
        End If
        If Form1.TreeView1.SelectedItem.Children = 3 Then
            If Form1.TreeView1.SelectedItem.Child.Next.Next.Key = "t" + Trim(Str(X + 100)) Then
                Form1.TreeView1.Nodes.Remove (Form1.TreeView1.SelectedItem.Child.Next.Next.Index)
            End If
        End If
    End If
    Set nodX = Form1.TreeView1.Nodes.Add(Read, tvwChild, "t" + Trim(Str(X + 100)), "Track File: " + CommonDialog1.FileName, 4, 4)
    FileSaved = False
    
    Exit Sub
ErrorTrap:
    Select Case Err.Number
    Case "32755"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Private Sub cmdGet_Click()
    frmBestLap.Show , MDIForm1
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo ErrorTrap
    Dir1.Path = Drive1.Drive
Exit Sub
ErrorTrap:
    Select Case Err.Number
    Case 68
        MsgBox Err.Description, vbCritical
        Drive1.Drive = "c:"
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Private Sub File1_Click()
    lblAuthor.Caption = "Author:"
    lblEvent.Caption = "Event:"
    lblYear.Caption = "Year:"
    lblMisc.Caption = "Misc Info:"
    lblRLap.Caption = ""
    lblQLap.Caption = ""
    lblWare.Caption = ""
    lblLength.Caption = ""
    lblLaps.Caption = ""
    lblCountry.Caption = ""
    lblName.Caption = ""
    CountNr = Len(frmMain.File1.FileName)
    Read = UCase(Mid(frmMain.File1.FileName, CountNr - 2, 3))
    If (UCase(Read) = UCase("bmp")) Or (UCase(Read) = UCase("gif")) Then
        Read = Dir1.Path + "\" + File1.FileName
        Set imgSize.Picture = LoadPicture(Read)
        PicY = imgSize.Height / 15
        PicX = imgSize.Width / 15
        If ((PicX = 640) And (PicY = 480)) Then
            imgPre.Height = (PicY * 15) / 2
            imgPre.Width = (PicX * 15) / 2
            Set imgPre.Picture = LoadPicture(Read)
            NoSupport = False
            Exit Sub
        End If
        If ((PicX = 440) And (PicY = 330)) Then
            imgPre.Height = (PicY * 15) / 2
            imgPre.Width = (PicX * 15) / 2
            Set imgPre.Picture = LoadPicture(Read)
            NoSupport = False
            Exit Sub
        End If
        MsgBox "This picture is not supported by " + TH + ".", vbInformation, TH
        Exit Sub
    End If
    Set imgPre = Nothing
    If (Mid(Dir1.Path, 3, 1) = "\") And (Len(Dir1.Path) = 3) Then
        ReadTrackFile (Dir1.Path + File1.FileName)
    Else
        ReadTrackFile (Dir1.Path + "\" + File1.FileName)
    End If
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If NoSupport = False Then
        Dim DY  ' Declare variable.
        DY = TextHeight("A")    ' Get height of one line.
        Label1.Move File1.Left + 30, File1.Top + Y + DY / 3, File1.Width - 30, DY
        Label1.Drag ' Drag label outline.
        InDrag = True
    End If
End Sub

Private Sub File1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Source.Tag = Dir1.Path + "\" + File1.FileName
End Sub

Private Sub Form_Load()
    frmMain.frameFile.Visible = True
    frmMain.frameInfo.Visible = False
    frmMain.frameData.Visible = False
    frmMain.frmaeTime.Visible = False
    frmMain.frameFile.Top = frmMain.frameInfo.Top
    frmMain.frameFile.Left = frmMain.frameInfo.Left
    Read4 = ""
End Sub

Private Sub hscPGrip_Change()
    lblGrip.Caption = hscPGrip.Value
    FileSaved = False
End Sub

Private Sub hscPGrip_Scroll()
    lblGrip.Caption = hscPGrip.Value
End Sub

Private Sub hscPitSpeed_Change()
    X = hscPitSpeed.Value
    X = X * 1.5966
    lblPitSpeed.Caption = Str(hscPitSpeed.Value) + "mph (" + Trim(Str(X)) + "km/h)"
    FileSaved = False
End Sub

Private Sub hscPitSpeed_Scroll()
    X = hscPitSpeed.Value
    X = X * 1.5966
    lblPitSpeed.Caption = Str(hscPitSpeed.Value) + "mph (" + Trim(Str(X)) + "km/h)"
End Sub

Private Sub hscPQPower_Change()
    lblPQPower.Caption = hscPQPower.Value
    FileSaved = False
End Sub

Private Sub hscPQPower_Scroll()
    lblPQPower.Caption = hscPQPower.Value
End Sub

Private Sub hscPRPower_Change()
    lblPRPower.Caption = hscPRPower.Value
    FileSaved = False
End Sub

Private Sub hscPRPower_Scroll()
    lblPRPower.Caption = hscPRPower.Value
End Sub

Private Sub HScroll1_Change()
    X = HScroll1.Value
    X = X / 2.203020134
    lblCWeight.Caption = Str(HScroll1.Value) + "lb (" + Trim(Str(X)) + "kg)"
    FileSaved = False
End Sub

Private Sub HScroll1_Scroll()
    X = HScroll1.Value
    X = X / 2.203020134
    lblCWeight.Caption = Str(HScroll1.Value) + "lb (" + Trim(Str(X)) + "kg)"
End Sub

Private Sub HScroll2_Change()
    lblQuick.Caption = Str(HScroll2.Value) + "%"
    FileSaved = False
End Sub

Private Sub HScroll2_Scroll()
    lblQuick.Caption = Str(HScroll2.Value) + "%"
End Sub

Private Sub hscRPower_Change()
    lblRacePower.Caption = hscRPower.Value
    FileSaved = False
End Sub

Private Sub hscRPower_Scroll()
    lblRacePower.Caption = hscRPower.Value
End Sub

Private Sub hscWeight_Change()
    X = hscWeight.Value
    X = X / 2.203020134
    lblWeight2.Caption = Str(hscWeight.Value) + "lb (" + Trim(Str(X)) + "kg)"
    FileSaved = False
End Sub

Private Sub hscWeight_Scroll()
    X = hscWeight.Value
    X = X / 2.203020134
    lblWeight2.Caption = Str(hscWeight.Value) + "lb (" + Trim(Str(X)) + "kg)"
End Sub

Private Sub Slider1_Change()
    lblYear2.Caption = Slider1.Value
    FileSaved = False
End Sub

Private Sub Slider1_Scroll()
    lblYear2.Caption = Slider1.Value
End Sub

Private Sub txtAdjectiv_GotFocus()
    txtAdjectiv.SelStart = 0
    txtAdjectiv.SelLength = Len(txtAdjectiv)
End Sub

Private Sub txtAdjectiv_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtCountry_GotFocus()
    txtCountry.SelStart = 0
    txtCountry.SelLength = Len(txtCountry)
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtLaps_Change()
    If (txtLaps <> "") And (txtLaps <> "0") Then
        VScroll1.Value = txtLaps.Text
    End If
End Sub

Private Sub txtLaps_GotFocus()
    txtLaps.SelStart = 0
    txtLaps.SelLength = Len(txtLaps)
End Sub

Private Sub txtLaps_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtLength_GotFocus()
    txtLength.SelStart = 0
    txtLength.SelLength = Len(txtLength)
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_LostFocus()
    FileSaved = False
    If (Len(frmMain.txtName) > 0) Then
        X = Mid(Form1.TreeView1.SelectedItem.Key, 2, Len(Form1.TreeView1.SelectedItem.Key) - 1)
        If (X > 10) And (X < 100) Then X = X - 10
        If (X > 100) And (X < 200) Then X = X - 100
        If (X > 200) And (X < 300) Then X = X - 200
        If X > 300 Then X = X - 300
        
        Form1.TreeView1.SelectedItem.Text = Trim(Str(X)) + ". " + txtName.Text
    Else
        X = Mid(Form1.TreeView1.SelectedItem.Key, 2, 2) - 10
        Form1.TreeView1.SelectedItem.Text = "Track " + Trim(Str(X))
    End If
End Sub

Private Sub txtQDate_GotFocus()
    txtQDate.SelStart = 0
    txtQDate.SelLength = Len(txtQDate)
End Sub

Private Sub txtQDate_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtQDriver_GotFocus()
    txtQDriver.SelStart = 0
    txtQDriver.SelLength = Len(txtQDriver)
End Sub

Private Sub txtQDriver_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtQTeam_GotFocus()
    txtQTeam.SelStart = 0
    txtQTeam.SelLength = Len(txtQTeam)
End Sub

Private Sub txtQTeam_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtQTime_GotFocus()
    txtQTime.SelStart = 0
    txtQTime.SelLength = Len(txtQTime)
End Sub

Private Sub txtQTime_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtRDate_GotFocus()
    txtRDate.SelStart = 0
    txtRDate.SelLength = Len(txtRDate)
End Sub

Private Sub txtRDate_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtRDriver_GotFocus()
    txtRDriver.SelStart = 0
    txtRDriver.SelLength = Len(txtRDriver)
End Sub

Private Sub txtRDriver_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtRTeam_GotFocus()
    txtRTeam.SelStart = 0
    txtRTeam.SelLength = Len(txtRTeam)
End Sub

Private Sub txtRTeam_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtRTime_GotFocus()
    txtRTime.SelStart = 0
    txtRTime.SelLength = Len(txtRTime)
End Sub

Private Sub txtRTime_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub txtTire_GotFocus()
    txtTire.SelStart = 0
    txtTire.SelLength = Len(txtTire)
End Sub

Private Sub txtTire_KeyPress(KeyAscii As Integer)
    FileSaved = False
End Sub

Private Sub VScroll1_Change()
    txtLaps = VScroll1.Value
End Sub
