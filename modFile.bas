Attribute VB_Name = "modFile"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINI(Section, KeyName, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))
    DoEvents
End Function

Public Function WriteINI(ByVal sSection As String, ByVal sKeyName As String, ByVal sNewString As String, ByVal sFileName)
    Dim R
    R = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
    DoEvents
End Function

Public Sub NewFile()
    On Error Resume Next
    Kill (TempFile)
    Randomize
    X = Int((500) * Rnd)
    TempFile = ProgramDir & "\File\th16" & Trim(Str(X)) & ".lda"
    FileCopy ProgramDir & "\Mall.lda", TempFile
    FileInfo.Name = ""
    FileInfo.Path = ""
    FileInfo.Saved = True
    FileInfo.Import = False
    For X = 0 To 15
        Tracks(X) = False
    Next
End Sub

Public Sub SaveTrackData(ByVal TNr As Integer)
    With frmMain
        WriteINI "Track " & TNr, "TPath", .txtPath, TempFile
        WriteINI "Track " & TNr, "Laps", .updLaps.Text, TempFile
        WriteINI "Track " & TNr, "Ware", .txtTire, TempFile
        WriteINI "Track " & TNr, "Length", .txtLength, TempFile
        WriteINI "Track " & TNr, "Name", .txtName, TempFile
        WriteINI "Track " & TNr, "Country", .txtCountry, TempFile

        WriteINI "Track " & TNr, "Adjective", .txtAdjectiv.Text, TempFile
        WriteINI "Track " & TNr, "RTime", .txtRTime, TempFile
        WriteINI "Track " & TNr, "QTime", .txtQTime, TempFile
        WriteINI "Track " & TNr, "RDate", .txtRDate, TempFile
        WriteINI "Track " & TNr, "QDate", .txtQDate, TempFile
        WriteINI "Track " & TNr, "RDriver", .txtRDriver, TempFile
        WriteINI "Track " & TNr, "QDriver", .txtQDriver, TempFile
        WriteINI "Track " & TNr, "RTeam", .txtRTeam, TempFile
        WriteINI "Track " & TNr, "QTeam", .txtQTeam, TempFile
        WriteINI "Track " & TNr, "BPic", .txtBPic.Text, TempFile
        WriteINI "Track " & TNr, "SPic", .txtSPic.Text, TempFile
    End With
End Sub

Public Sub GetTrackData(ByVal TNr As Integer)
Dim sLaps As String
Dim iLaps As Integer
    With frmMain
        .txtPath = ReadINI("Track " & TNr, "TPath", TempFile)
        sLaps = ReadINI("Track " & TNr, "Laps", TempFile)
        If sLaps <> "" Then
            iLaps = Int(sLaps)
            If (iLaps > 2) And (iLaps < 127) Then
                .updLaps.Value = iLaps
            Else
                .updLaps.Value = 3
            End If
        Else
            .updLaps.Value = 3
        End If
        .txtTire = ReadINI("Track " & TNr, "Ware", TempFile)
        .txtLength = ReadINI("Track " & TNr, "Length", TempFile)
        .txtName = ReadINI("Track " & TNr, "Name", TempFile)
        .txtCountry = ReadINI("Track " & TNr, "Country", TempFile)

        .txtAdjectiv.Text = ReadINI("Track " & TNr, "Adjective", TempFile)
        .txtRTime = ReadINI("Track " & TNr, "RTime", TempFile)
        .txtQTime = ReadINI("Track " & TNr, "QTime", TempFile)
        .txtRDate = ReadINI("Track " & TNr, "RDate", TempFile)
        .txtQDate = ReadINI("Track " & TNr, "QDate", TempFile)
        .txtRDriver = ReadINI("Track " & TNr, "RDriver", TempFile)
        .txtQDriver = ReadINI("Track " & TNr, "QDriver", TempFile)
        .txtRTeam = ReadINI("Track " & TNr, "RTeam", TempFile)
        .txtQTeam = ReadINI("Track " & TNr, "QTeam", TempFile)
        Read = ReadINI("Track " & TNr, "BPic", TempFile)
        If Read <> "" Then
            Set frmMain.picMenuPic = LoadPicture(Read)
            .txtBPic.Text = Read
        Else
            Set frmMain.picMenuPic.Picture = Nothing
            .txtBPic.Text = ""
        End If
        Read = ReadINI("Track " & TNr, "SPic", TempFile)
        If Read <> "" Then
            Set frmMain.picMenuPic = LoadPicture(Read)
            .txtSPic.Text = Read
        Else
            Set frmMain.picMenuPic.Picture = Nothing
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

    Set nodX = frmMain.TreeView1.Nodes.Add(, , "r", "Gp2 Track's", 1, 2)

    For X = 1 To 16
        Read = ReadINI("Track " & X, "TPath", TempFile)
        Name = ReadINI("Track " & X, "Name", TempFile)
        BPic = ReadINI("Track " & X, "BPic", TempFile)
        SPic = ReadINI("Track " & X, "SPic", TempFile)
        If Read <> "" Then
            Tracks(X - 1) = True
            If Name <> "" Then
                frmMain.TreeView1.Nodes.Add "r", tvwChild, "t" & X + 10, Trim(Str(X)) & ". " & Name, 1, 2
            Else
                frmMain.TreeView1.Nodes.Add "r", tvwChild, "t" & X + 10, Trim(Str(X)) & ". -[No Name]-", 1, 2
            End If
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
        FileCopy TempFile, FileInfo.Path
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

    Read = oFile.ShowSave("Track Handler Files (*.ths)|*.ths|All Files (*.*)|*.*|", "ths", frmMain.hWnd)
    If Read = "" Then Exit Sub
    FileInfo.Path = Read
    FileInfo.Name = oFile.GetFilePart(FileInfo.Path, GetFileName)

    FileCopy TempFile, FileInfo.Path
    FileInfo.Saved = True
    RecentFile FileInfo.Path
    frmMain.LoadRecent
    frmMain.Caption = TH & " v1.6 [" & FileInfo.Name & "]"

Exit Sub

ErrHandler:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: SaveFileAs()", vbCritical, TH & " - Error"
End Sub

Public Sub SaveMisc()
    With frmMain
        WriteINI "Misc", "Year", .Slider1.Value, TempFile
        WriteINI "Misc", "Quick", .hscQRace.Value, TempFile
        WriteINI "Misc", "CWeight", .hscCWeight.Value, TempFile
        WriteINI "Misc", "0as1", .chk0as1.Value, TempFile
        WriteINI "Misc", "SaveLap", .chkSave.Value, TempFile

        WriteINI "Player", "RPower", .hscPRPower.Value, TempFile
        WriteINI "Player", "QPower", .hscPQPower.Value, TempFile
        WriteINI "Player", "Grip", .hscPGrip.Value, TempFile
        WriteINI "Player", "Speed", .hscPitSpeed.Value, TempFile
        WriteINI "Player", "NoSpeed", .chkNoLimit.Value, TempFile
        WriteINI "Player", "Weight", .hscWeight.Value, TempFile
        WriteINI "Player", "UseTeam", .chkUPower.Value, TempFile
        WriteINI "Misc", "CCFuel", .chkCCFuel.Value, TempFile

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
        WriteINI "Misc", "Aids", Read, TempFile
    End With
End Sub

Public Sub GetMisc()
    On Error Resume Next
    With frmMain
        .Slider1.Value = ReadINI("Misc", "Year", TempFile)
        .hscQRace.Value = ReadINI("Misc", "Quick", TempFile)
        .hscCWeight.Value = ReadINI("Misc", "CWeight", TempFile)
        .chk0as1.Value = ReadINI("Misc", "0as1", TempFile)
        .chkSave.Value = ReadINI("Misc", "SaveLap", TempFile)

        .hscPRPower.Value = ReadINI("Player", "RPower", TempFile)
        .hscPQPower.Value = ReadINI("Player", "QPower", TempFile)
        .hscPGrip.Value = ReadINI("Player", "Grip", TempFile)
        .hscPitSpeed.Value = ReadINI("Player", "Speed", TempFile)
        .chkNoLimit.Value = ReadINI("Player", "NoSpeed", TempFile)
        .hscWeight.Value = ReadINI("Player", "Weight", TempFile)
        .chkUPower.Value = ReadINI("Player", "UseTeam", TempFile)
        .chkCCFuel.Value = ReadINI("Misc", "CCFuel", TempFile)
    End With
    Read = ReadINI("Misc", "Aids", TempFile)
    LoadGp2Aid Read
End Sub

Public Sub ResetTime(ByVal sDName As String, ByVal sTName As String, ByVal sTime As String, ByVal sDate As String)
    For X = 1 To 16
        WriteINI "Track " & X, "QTime", sTime, TempFile
        WriteINI "Track " & X, "RTime", sTime, TempFile
        WriteINI "Track " & X, "QDriver", sDName, TempFile
        WriteINI "Track " & X, "RDriver", sDName, TempFile
        WriteINI "Track " & X, "QTeam", sTName, TempFile
        WriteINI "Track " & X, "RTeam", sTName, TempFile
        WriteINI "Track " & X, "QDate", sDate, TempFile
        WriteINI "Track " & X, "RDate", sDate, TempFile
    Next
End Sub

Public Sub SaveImport()
    tVar.iInt = MsgBox("Please select a destination directory of the imported track's." & vbLf & _
        "The track files will get the same name as the track.", vbOKCancel, TH)
    If tVar.iInt = vbCancel Then Exit Sub
    If tVar.iInt = vbOK Then
        Read = oFile.BrowseFolders("Select Track Directory", frmMain.hWnd)
        If Read = "" Then Exit Sub
        If Len(Read) = 3 Then
            Read = Mid(Read, 1, 2)
        End If
        For X = 1 To 16
            If X < 10 Then
                Read3 = Gp2Dir & "\Circuits\f1ct0" & X & ".dat"
            Else
                Read3 = Gp2Dir & "\Circuits\f1ct" & X & ".dat"
            End If
            Read2 = ReadINI("Track " & X, "Name", TempFile)
            Read4 = oFile.FileExists(Read & "\" & Read2 & ".dat")
            If Read4 = True Then
                Read2 = "TH" & Read2
            End If
            FileCopy Read3, Read & "\" & Read2 & ".dat"
            WriteINI "Track " & X, "TPath", Read & "\" & Read2 & ".dat", TempFile
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
    WriteINI "Misc", "Point", Read, TempFile
End Sub

Public Sub GetPoint()
    With frmPoint
        Load frmPoint
        Read = ""
        Read = ReadINI("Misc", "Point", TempFile)
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

Public Sub RecentFile(ByVal FilePath As String)
Dim vArray As Variant
Dim Found As Boolean
    Found = False
    vArray = oReg.GetAllValues(HKEY_CURRENT_USER, "Software\VG Software\Gp2 Track Handler\Files")
    If Not IsArray(vArray) Then
        ReDim vArray(2, 1)
    End If
    For X = 0 To UBound(vArray, 1)
        If vArray(X, 1) = FilePath Then
            Found = True
            If X = 0 Then
                Exit Sub
            ElseIf X = 1 Then
                vArray(1, 1) = vArray(0, 1)
                vArray(0, 1) = FilePath
            ElseIf X = 2 Then
                vArray(2, 1) = vArray(1, 1)
                vArray(1, 1) = vArray(0, 1)
                vArray(0, 1) = FilePath
            End If
            Exit For
        End If
    Next
    If Found = False Then
        vArray(2, 1) = vArray(1, 1)
        vArray(1, 1) = vArray(0, 1)
        vArray(0, 1) = FilePath
    End If
    For X = 0 To UBound(vArray, 1)
        oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\VG Software\Gp2 Track Handler\Files", X, vArray(X, 1)
    Next
End Sub
