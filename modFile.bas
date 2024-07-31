Attribute VB_Name = "modFile"
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


Public Sub NewFile()
    On Error Resume Next
    Kill (TempFile)
    Randomize
    X = Int((500) * Rnd)
    TempFile = ProgramDir & "\File\th14" & Trim(Str(X)) & ".lda"
    FileCopy ProgramDir & "\Mall.lda", TempFile
    FileInfo.Name = ""
    FileInfo.Path = ""
    FileInfo.Saved = True
    FileInfo.Changes = False
    FileInfo.Import = False
    For X = 0 To 15
        Tracks(X) = False
    Next
End Sub

Public Sub SaveTrackData(ByVal TNr As Integer)
    With frmMain
        oMisc.WriteINI "Track " & TNr, "TPath", .txtPath, TempFile
        oMisc.WriteINI "Track " & TNr, "Laps", .txtLaps, TempFile
        oMisc.WriteINI "Track " & TNr, "Ware", .txtTire, TempFile
        oMisc.WriteINI "Track " & TNr, "Length", .txtLength, TempFile
        oMisc.WriteINI "Track " & TNr, "Name", .txtName, TempFile
        oMisc.WriteINI "Track " & TNr, "Country", .txtCountry, TempFile

        oMisc.WriteINI "Track " & TNr, "Adjective", .txtAdjectiv, TempFile
        oMisc.WriteINI "Track " & TNr, "RTime", .txtRTime, TempFile
        oMisc.WriteINI "Track " & TNr, "QTime", .txtQTime, TempFile
        oMisc.WriteINI "Track " & TNr, "RDate", .txtRDate, TempFile
        oMisc.WriteINI "Track " & TNr, "QDate", .txtQDate, TempFile
        oMisc.WriteINI "Track " & TNr, "RDriver", .txtRDriver, TempFile
        oMisc.WriteINI "Track " & TNr, "QDriver", .txtQDriver, TempFile
        oMisc.WriteINI "Track " & TNr, "RTeam", .txtRTeam, TempFile
        oMisc.WriteINI "Track " & TNr, "QTeam", .txtQTeam, TempFile
        oMisc.WriteINI "Track " & TNr, "BPic", .txtBPic.Text, TempFile
        oMisc.WriteINI "Track " & TNr, "SPic", .txtSPic.Text, TempFile
    End With
End Sub

Public Sub GetTrackData(ByVal TNr As Integer)
    With frmMain
        .txtPath = oMisc.ReadINI("Track " & TNr, "TPath", TempFile)
        .txtLaps = oMisc.ReadINI("Track " & TNr, "Laps", TempFile)
        .txtTire = oMisc.ReadINI("Track " & TNr, "Ware", TempFile)
        .txtLength = oMisc.ReadINI("Track " & TNr, "Length", TempFile)
        .txtName = oMisc.ReadINI("Track " & TNr, "Name", TempFile)
        .txtCountry = oMisc.ReadINI("Track " & TNr, "Country", TempFile)

        .txtAdjectiv = oMisc.ReadINI("Track " & TNr, "Adjective", TempFile)
        .txtRTime = oMisc.ReadINI("Track " & TNr, "RTime", TempFile)
        .txtQTime = oMisc.ReadINI("Track " & TNr, "QTime", TempFile)
        .txtRDate = oMisc.ReadINI("Track " & TNr, "RDate", TempFile)
        .txtQDate = oMisc.ReadINI("Track " & TNr, "QDate", TempFile)
        .txtRDriver = oMisc.ReadINI("Track " & TNr, "RDriver", TempFile)
        .txtQDriver = oMisc.ReadINI("Track " & TNr, "QDriver", TempFile)
        .txtRTeam = oMisc.ReadINI("Track " & TNr, "RTeam", TempFile)
        .txtQTeam = oMisc.ReadINI("Track " & TNr, "QTeam", TempFile)
        Read = oMisc.ReadINI("Track " & TNr, "BPic", TempFile)
        If Read <> "" Then
            Set frmMain.imgBPic.Picture = LoadPicture(Read)
            .txtBPic.Text = Read
        Else
            Set frmMain.imgBPic.Picture = Nothing
            .txtBPic.Text = ""
        End If
        Read = oMisc.ReadINI("Track " & TNr, "SPic", TempFile)
        If Read <> "" Then
            Set frmMain.imgSPic.Picture = LoadPicture(Read)
            .txtSPic.Text = Read
        Else
            Set frmMain.imgSPic.Picture = Nothing
            .txtSPic.Text = ""
        End If
    End With
End Sub

Public Sub LoadFile()
Dim Name As String
Dim BPic As String
Dim SPic As String
    'Bygg om trädet
    frmMain.TreeView1.Nodes.Clear
    Dim nodX As Node    ' Create variable.

    Set nodX = frmMain.TreeView1.Nodes.Add(, , "r", "GP2 Track's", 1, 2)

    For X = 1 To 16
        Read = oMisc.ReadINI("Track " & X, "TPath", TempFile)
        Name = oMisc.ReadINI("Track " & X, "Name", TempFile)
        BPic = oMisc.ReadINI("Track " & X, "BPic", TempFile)
        SPic = oMisc.ReadINI("Track " & X, "SPic", TempFile)
        If Read <> "" Then
            Tracks(X - 1) = True
            frmMain.TreeView1.Nodes.Add "r", tvwChild, "t" & X + 10, Trim(Str(X)) & ". " & Name, 1, 2
            frmMain.TreeView1.Nodes.Add "t" & Trim(Str(X + 10)), tvwChild, "t" & Trim(Str(X + 10)) & "-Track", Read, 3, 3
            If BPic <> "" Then
                frmMain.TreeView1.Nodes.Add "t" & Trim(Str(X + 10)), tvwChild, "t" & Trim(Str(X + 10)) & "-BPic", "Big Pic: " & BPic, 4, 4
            End If
            If SPic <> "" Then
                frmMain.TreeView1.Nodes.Add "t" & Trim(Str(X + 10)), tvwChild, "t" & Trim(Str(X + 10)) & "-SPic", "Small Pic: " & SPic, 4, 4
            End If
        Else
            frmMain.TreeView1.Nodes.Add "r", tvwChild, "t" & X + 10, "Track " & X, 1, 2
        End If
    Next
    frmMain.TreeView1.Nodes(2).EnsureVisible
    If TreeNr > 0 Then
        GetTrackData TreeNr
        frmMain.TreeView1.Nodes("t" & TreeNr + 10).Selected = True
        frmMain.TreeView1_NodeClick frmMain.TreeView1.Nodes("t" & TreeNr + 10)
    End If
    GetMisc
End Sub

Public Sub SaveFile()
    SaveMisc
    If FileInfo.Import = True Then
        SaveImport
        Exit Sub
    End If
    If FileInfo.Name <> "" Then
        CopyFile TempFile, FileInfo.Path, 0
        DoEvents
    Else
        SaveFileAs
    End If
End Sub

Public Sub SaveFileAs()
Dim GetSave As String
    On Error GoTo ErrHandler
    If FileInfo.Import = True Then
        SaveImport
        Exit Sub
    End If

    Var.sString1 = oMisc.ShowSave("Track Handler Files (*.ths)|*.ths|All Files (*.*)|*.*|", "ths", frmMain.hwnd, ProgramDir)
    If Var.sString1 = "" Then Exit Sub
    FileInfo.Path = Var.sString1
    For X = Len(Var.sString1) To 0 Step -1
        If Mid(Var.sString1, X, 1) = "\" Then Exit For
    Next
    FileInfo.Name = Mid(Var.sString1, X + 1)

    FileCopy TempFile, FileInfo.Path
    FileInfo.Saved = True
    Read = oMisc.RecentFile(SaveNew, FileInfo.Path, FileInfo.Name)
    frmMain.LoadRecent
    frmMain.Caption = "GP2 Track Handler v1.5 [" & FileInfo.Name & "]"

Exit Sub

ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: " & Err.Source, vbCritical, "Error"
End Sub

Public Sub SaveMisc()
    With frmMain
        oMisc.WriteINI "Misc", "Year", .Slider1.Value, TempFile
        oMisc.WriteINI "Misc", "Quick", .hscQRace.Value, TempFile
        oMisc.WriteINI "Misc", "CWeight", .hscCWeight.Value, TempFile
        oMisc.WriteINI "Misc", "0as1", .chk0as1.Value, TempFile
        oMisc.WriteINI "Misc", "SaveLap", .chkSave.Value, TempFile

        oMisc.WriteINI "Player", "RPower", .hscPRPower.Value, TempFile
        oMisc.WriteINI "Player", "QPower", .hscPQPower.Value, TempFile
        oMisc.WriteINI "Player", "Grip", .hscPGrip.Value, TempFile
        oMisc.WriteINI "Player", "Speed", .hscPitSpeed.Value, TempFile
        oMisc.WriteINI "Player", "NoSpeed", .chkNoLimit.Value, TempFile
        oMisc.WriteINI "Player", "Weight", .hscWeight.Value, TempFile
        oMisc.WriteINI "Player", "UseTeam", .chkUPower.Value, TempFile

        Read = ""
        For X = 0 To 6
            If frmMain.R(X).Tag = "On" Then
                Read = Read & "1"
            Else
                Read = Read & "0"
            End If
        Next
        For X = 0 To 6
            If frmMain.A(X).Tag = "On" Then
                Read = Read & "1"
            Else
                Read = Read & "0"
            End If
        Next
        For X = 0 To 6
            If frmMain.S(X).Tag = "On" Then
                Read = Read & "1"
            Else
                Read = Read & "0"
            End If
        Next
        For X = 0 To 6
            If frmMain.P(X).Tag = "On" Then
                Read = Read & "1"
            Else
                Read = Read & "0"
            End If
        Next
        For X = 0 To 6
            If frmMain.AC(X).Tag = "On" Then
                Read = Read & "1"
            Else
                Read = Read & "0"
            End If
        Next
        oMisc.WriteINI "Misc", "Aids", Read, TempFile
    End With
End Sub

Public Sub GetMisc()
    With frmMain
        .Slider1.Value = oMisc.ReadINI("Misc", "Year", TempFile)
        .hscQRace.Value = oMisc.ReadINI("Misc", "Quick", TempFile)
        .hscCWeight.Value = oMisc.ReadINI("Misc", "CWeight", TempFile)
        .chk0as1.Value = oMisc.ReadINI("Misc", "0as1", TempFile)
        .chkSave.Value = oMisc.ReadINI("Misc", "SaveLap", TempFile)

        .hscPRPower.Value = oMisc.ReadINI("Player", "RPower", TempFile)
        .hscPQPower.Value = oMisc.ReadINI("Player", "QPower", TempFile)
        .hscPGrip.Value = oMisc.ReadINI("Player", "Grip", TempFile)
        .hscPitSpeed.Value = oMisc.ReadINI("Player", "Speed", TempFile)
        .chkNoLimit.Value = oMisc.ReadINI("Player", "NoSpeed", TempFile)
        .hscWeight.Value = oMisc.ReadINI("Player", "Weight", TempFile)
        .chkUPower.Value = oMisc.ReadINI("Player", "UseTeam", TempFile)
    End With
    LoadGP2Aid
End Sub

Public Sub ResetTime(ByVal sDName As String, ByVal sTName As String, ByVal sTime As String, ByVal sDate As String)
    For X = 1 To 16
        oMisc.WriteINI "Track " & X, "QTime", sTime, TempFile
        oMisc.WriteINI "Track " & X, "RTime", sTime, TempFile
        oMisc.WriteINI "Track " & X, "QDriver", sDName, TempFile
        oMisc.WriteINI "Track " & X, "RDriver", sDName, TempFile
        oMisc.WriteINI "Track " & X, "QTeam", sTName, TempFile
        oMisc.WriteINI "Track " & X, "RTeam", sTName, TempFile
        oMisc.WriteINI "Track " & X, "QDate", sDate, TempFile
        oMisc.WriteINI "Track " & X, "RDate", sDate, TempFile
    Next
End Sub

Public Sub SaveImport()
    Var.iInt1 = MsgBox(LoadResString(125), vbOKCancel, TH)
    If Var.iInt1 = vbCancel Then Exit Sub
    If Var.iInt1 = vbOK Then
        Read = oMisc.BrowseFolders("Select Track Directory", frmMain.hwnd)
        If Read = "" Then Exit Sub
        If Len(Read) = 3 Then
            Read = Mid(Read, 1, 2)
        End If
        For X = 1 To 16
            If X < 10 Then
                Read3 = GP2Dir & "\Circuits\f1ct0" & X & ".dat"
            Else
                Read3 = GP2Dir & "\Circuits\f1ct" & X & ".dat"
            End If
            Read2 = oMisc.ReadINI("Track " & X, "Name", TempFile)
            Read4 = oMisc.File_Exists(Read & "\" & Read2 & ".dat")
            If Read4 = True Then
                Read2 = "TH" & Read2
            End If
            FileCopy Read3, Read & "\" & Read2 & ".dat"
            oMisc.WriteINI "Track " & X, "TPath", Read & "\" & Read2 & ".dat", TempFile
        Next
        FileInfo.Import = False
        If Trim(FileInfo.Name) <> "" Then
            FileCopy TempFile, FileInfo.Path
        Else
            SaveFileAs
        End If
    End If
End Sub

Public Sub SavePoint()
    With frmPoint
        Load frmPoint
        Read = ""
        For X = 0 To 25
            If Len(.txtPoint(X)) = 2 Then
                Read = Read & .txtPoint(X)
            ElseIf Len(.txtPoint(X)) = 1 Then
                Read = Read & "0" & .txtPoint(X)
            ElseIf Len(.txtPoint(X)) = 0 Then
                Read = Read & "00"
            End If
        Next
    End With
    oMisc.WriteINI "Misc", "Point", Read, TempFile
End Sub

Public Sub GetPoint()
    With frmPoint
        Load frmPoint
        Read = ""
        Read = oMisc.ReadINI("Misc", "Point", TempFile)
        For X = 0 To 25
            Read2 = Mid(Read, X * 2 + 1, 2)
            If Mid(Read2, 1, 1) = "0" Then
                .txtPoint(X).Text = Mid(Read2, 2, 1)
            Else
                .txtPoint(X).Text = Read2
            End If
        Next
    End With
End Sub
