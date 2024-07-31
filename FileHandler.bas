Attribute VB_Name = "FileHandler"
Dim oMisc As New TrackHandler.Misc

Public Sub NewFile()
    MDIForm1.Caption = GP2TH + " v1.4 - Noname.ths"
    DeleteFile ProgramDir + "\WorkCopy.lda"
    TargetFile = ProgramDir + "\WorkCopy.lda"
    SourceFile = ProgramDir + "\Mall.lda"
    FileCopy SourceFile, TargetFile
    NewTree
    Rensa
    MDIForm1.txtFramedPath = ""
    MDIForm1.txtFullPath = ""
    Set MDIForm1.imgFramed = Nothing
    Set MDIForm1.imgFull = Nothing
    FileInfo.FileName = ""
    FileInfo.FilePath = ""
    FileInfo.FileType = FileNew
End Sub

Public Sub GetFileData(NodeNr As Integer)
    Read = ProgramDir + "\WorkCopy.lda"
    Read2 = "Track " + Trim(Str(NodeNr))
    MDIForm1.txtAdjectiv.Text = oMisc.ReadINI(Read2, "Adjective", Read)
    MDIForm1.txtCountry.Text = oMisc.ReadINI(Read2, "Country", Read)
    MDIForm1.txtFramedPath.Text = oMisc.ReadINI(Read2, "SPic", Read)
    MDIForm1.txtFullPath.Text = oMisc.ReadINI(Read2, "BPic", Read)
    MDIForm1.txtLaps.Text = oMisc.ReadINI(Read2, "Laps", Read)
    If MDIForm1.txtLaps.Text <> "" Then
        MDIForm1.VScroll1.Value = MDIForm1.txtLaps.Text
    Else
        MDIForm1.VScroll1.Value = 3
        MDIForm1.txtLaps = ""
    End If
    MDIForm1.txtLength.Text = oMisc.ReadINI(Read2, "Length", Read)
    MDIForm1.txtName.Text = oMisc.ReadINI(Read2, "Name", Read)
    MDIForm1.txtPath.Text = oMisc.ReadINI(Read2, "TPath", Read)
    MDIForm1.txtQDate.Text = oMisc.ReadINI(Read2, "QDate", Read)
    MDIForm1.txtRDate.Text = oMisc.ReadINI(Read2, "RDate", Read)
    MDIForm1.txtQDriver.Text = oMisc.ReadINI(Read2, "QDriver", Read)
    MDIForm1.txtRDriver.Text = oMisc.ReadINI(Read2, "RDriver", Read)
    MDIForm1.txtTire.Text = oMisc.ReadINI(Read2, "Ware", Read)
    MDIForm1.txtQTime.Text = oMisc.ReadINI(Read2, "QTime", Read)
    MDIForm1.txtRTime.Text = oMisc.ReadINI(Read2, "RTime", Read)
    MDIForm1.txtQTeam.Text = oMisc.ReadINI(Read2, "QTeam", Read)
    MDIForm1.txtRTeam.Text = oMisc.ReadINI(Read2, "RTeam", Read)
    If MDIForm1.txtFramedPath <> "" Then
        MDIForm1.lblFramed.Visible = True
        Set MDIForm1.imgFramed.Picture = LoadPicture(MDIForm1.txtFramedPath)
    Else
        MDIForm1.lblFramed.Visible = False
        MDIForm1.imgFramed.Picture = Nothing
    End If
    If MDIForm1.txtFullPath <> "" Then
        MDIForm1.lblFull.Visible = True
        Set MDIForm1.imgFull.Picture = LoadPicture(MDIForm1.txtFullPath)
    Else
        MDIForm1.imgFull.Picture = Nothing
        MDIForm1.lblFull.Visible = False
    End If
End Sub

Public Sub SaveFileData()
    Read = ProgramDir + "\WorkCopy.lda"
    Read2 = "Track " + Trim(Str(CurrentRecord))
    Read3 = oMisc.WriteINI(Read2, "Adjective", MDIForm1.txtAdjectiv.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "Country", MDIForm1.txtCountry.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "SPic", MDIForm1.txtFramedPath.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "BPic", MDIForm1.txtFullPath.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "Laps", MDIForm1.txtLaps.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "Length", MDIForm1.txtLength.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "Name", MDIForm1.txtName.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "TPath", MDIForm1.txtPath.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "QDate", MDIForm1.txtQDate.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "RDate", MDIForm1.txtRDate.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "QDriver", MDIForm1.txtQDriver.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "RDriver", MDIForm1.txtRDriver.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "Ware", MDIForm1.txtTire.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "QTime", MDIForm1.txtQTime.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "RTime", MDIForm1.txtRTime.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "QTeam", MDIForm1.txtQTeam.Text, Read)
    Read3 = oMisc.WriteINI(Read2, "RTeam", MDIForm1.txtRTeam.Text, Read)
End Sub

Public Sub GetPlayerData()
    Read = ProgramDir + "\WorkCopy.lda"
    Read2 = "Player"
    MDIForm1.hscPGrip = oMisc.ReadINI(Read2, "Grip", Read)
    MDIForm1.hscPitSpeed = oMisc.ReadINI(Read2, "Speed", Read)
    MDIForm1.hscPQPower = oMisc.ReadINI(Read2, "QPower", Read)
    MDIForm1.hscPRPower = oMisc.ReadINI(Read2, "RPower", Read)
    MDIForm1.hscWeight = oMisc.ReadINI(Read2, "Weight", Read)
    MDIForm1.chkNoLimit = oMisc.ReadINI(Read2, "NoSpeed", Read)
    MDIForm1.chkSelectedTeam = oMisc.ReadINI(Read2, "UseTeam", Read)
    
    Read2 = "Misc"
    MDIForm1.HScroll1.Value = oMisc.ReadINI(Read2, "CWeight", Read)
    MDIForm1.HScroll2.Value = oMisc.ReadINI(Read2, "Quick", Read)
    MDIForm1.chk0as1 = oMisc.ReadINI(Read2, "0as1", Read)
    MDIForm1.chkSave = oMisc.ReadINI(Read2, "SaveLap", Read)
    MDIForm1.Slider1 = oMisc.ReadINI(Read2, "Year", Read)
    
    Read2 = oMisc.ReadINI("Misc", "Aids", ProgramDir + "\WorkCopy.lda")
    If Read2 = "" Then
        GP2AidsSet
    Else
        Read = Mid(Read2, 1, 7)
        X = 1
        Do Until X > 7
            Read3 = Mid(Read, X, 1)
            If Read3 = "1" Then
                MDIForm1.R(X - 1).Picture = MDIForm1.On1(X - 1).Picture
                MDIForm1.R(X - 1).Tag = "Off"
            Else
                MDIForm1.R(X - 1).Picture = MDIForm1.Off(X - 1).Picture
                MDIForm1.R(X - 1).Tag = "On"
            End If
            X = X + 1
        Loop
        X = 1
        Read = Mid(Read2, 8, 7)
        Do Until X > 7
            Read3 = Mid(Read, X, 1)
            If Read3 = "1" Then
                MDIForm1.A(X - 1).Picture = MDIForm1.On1(X - 1).Picture
                MDIForm1.A(X - 1).Tag = "Off"
            Else
                MDIForm1.A(X - 1).Picture = MDIForm1.Off(X - 1).Picture
                MDIForm1.A(X - 1).Tag = "On"
            End If
            X = X + 1
        Loop
        Read = Mid(Read2, 15, 7)
        X = 1
        Do Until X > 7
            Read3 = Mid(Read, X, 1)
            If Read3 = "1" Then
                MDIForm1.S(X - 1).Picture = MDIForm1.On1(X - 1).Picture
                MDIForm1.S(X - 1).Tag = "Off"
            Else
                MDIForm1.S(X - 1).Picture = MDIForm1.Off(X - 1).Picture
                MDIForm1.S(X - 1).Tag = "On"
            End If
            X = X + 1
        Loop
        Read = Mid(Read2, 22, 7)
        X = 1
        Do Until X > 7
            Read3 = Mid(Read, X, 1)
            If Read3 = "1" Then
                MDIForm1.P(X - 1).Picture = MDIForm1.On1(X - 1).Picture
                MDIForm1.P(X - 1).Tag = "Off"
            Else
                MDIForm1.P(X - 1).Picture = MDIForm1.Off(X - 1).Picture
                MDIForm1.P(X - 1).Tag = "On"
            End If
            X = X + 1
        Loop
        Read = Mid(Read2, 29, 7)
        X = 1
        Do Until X > 7
            Read3 = Mid(Read, X, 1)
            If Read3 = "1" Then
                MDIForm1.AC(X - 1).Picture = MDIForm1.On1(X - 1).Picture
                MDIForm1.AC(X - 1).Tag = "Off"
            Else
                MDIForm1.AC(X - 1).Picture = MDIForm1.Off(X - 1).Picture
                MDIForm1.AC(X - 1).Tag = "On"
            End If
            X = X + 1
        Loop
    End If
End Sub

Public Sub SavePlayerData()
    Read = ""
    X = 0
    Do Until X = 7
        If MDIForm1.R(X).Tag = "On" Then
           Read = Read + "0"
        Else
            Read = Read + "1"
        End If
        X = X + 1
    Loop
    X = 0
    Do Until X = 7
        If MDIForm1.A(X).Tag = "On" Then
           Read = Read + "0"
        Else
            Read = Read + "1"
        End If
        X = X + 1
    Loop
    X = 0
    Do Until X = 7
        If MDIForm1.S(X).Tag = "On" Then
           Read = Read + "0"
        Else
            Read = Read + "1"
        End If
        X = X + 1
    Loop
    X = 0
    Do Until X = 7
        If MDIForm1.P(X).Tag = "On" Then
           Read = Read + "0"
        Else
            Read = Read + "1"
        End If
        X = X + 1
    Loop
    X = 0
    Do Until X = 7
        If MDIForm1.AC(X).Tag = "On" Then
           Read = Read + "0"
        Else
            Read = Read + "1"
        End If
        X = X + 1
    Loop
    Read2 = ProgramDir + "\WorkCopy.lda"
    Read = oMisc.WriteINI("Misc", "Aids", Read, ProgramDir + "\WorkCopy.lda")

    Read = ProgramDir + "\WorkCopy.lda"
    Read2 = "Player"
    Read3 = oMisc.WriteINI(Read2, "Grip", MDIForm1.hscPGrip, Read)
    Read3 = oMisc.WriteINI(Read2, "Speed", MDIForm1.hscPitSpeed, Read)
    Read3 = oMisc.WriteINI(Read2, "QPower", MDIForm1.hscPQPower, Read)
    Read3 = oMisc.WriteINI(Read2, "RPower", MDIForm1.hscPRPower, Read)
    Read3 = oMisc.WriteINI(Read2, "Weight", MDIForm1.hscWeight, Read)
    Read3 = oMisc.WriteINI(Read2, "NoSpeed", MDIForm1.chkNoLimit, Read)
    Read3 = oMisc.WriteINI(Read2, "UseTeam", MDIForm1.chkSelectedTeam, Read)
    
    Read2 = "Misc"
    Read3 = oMisc.WriteINI(Read2, "CWeight", MDIForm1.HScroll1.Value, Read)
    Read3 = oMisc.WriteINI(Read2, "Quick", MDIForm1.HScroll2.Value, Read)
    Read3 = oMisc.WriteINI(Read2, "0as1", MDIForm1.chk0as1, Read)
    Read3 = oMisc.WriteINI(Read2, "SaveLap", MDIForm1.chkSave, Read)
    Read3 = oMisc.WriteINI(Read2, "Year", MDIForm1.Slider1, Read)
End Sub

Public Sub SaveLastClick()
    If LastClick = "Player" Then
        SavePlayerData
    End If
    If LastClick = "Track" Then
        SaveFileData
    End If
End Sub

Public Sub OpenThFile(ByVal FilePath As String)
    DeleteFile ProgramDir & "\WorkCopy.lda"
    TargetFile = ProgramDir & "\WorkCopy.lda"
    SourceFile = FilePath
    FileCopy SourceFile, TargetFile
    MakeNewTree
End Sub

Public Sub SaveThFile()
    RetVal = False
    If FileInfo.FileType = FileOpen Then RetVal = oMisc.File_Exists(FileInfo.FilePath)
    If RetVal = True Then
        DeleteFile Trim(FileInfo.FilePath)
    End If
    TargetFile = Trim(FileInfo.FilePath)
    SourceFile = ProgramDir & "\WorkCopy.lda"
    FileCopy SourceFile, TargetFile
    oMisc.RecentFile 1, Trim(FileInfo.FilePath), Trim(FileInfo.FileName), GP2TH, SaveNew
End Sub
