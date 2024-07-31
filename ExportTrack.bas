Attribute VB_Name = "ExportTrack"
Public Sub SetAttribut()
    On Error Resume Next
    X = 1
    Do Until X > 16
        Read = Str(X)
        Read = Trim(Read)
        If Len(Read) < 2 Then Read = "0" + Read
        Read = Gp2Dir + "\Circuits\f1ct" + Read + ".dat"
        SetAttr Read, vbNormal
        X = X + 1
    Loop
    Read = Gp2Dir + "\gp2.exe"
    SetAttr Read, vbNormal
End Sub

Public Sub ExportLaps()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Laps", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Count2 = Read
        If Count2 > 126 Then Count2 = 126
        If Count2 < 3 Then Count2 = 3
        Read = Chr(Count2)
        Put #GP2FileNum, oData.Laps(GP2V) + CountExport, Read
    End If
End Sub

Public Sub ExportName()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Name", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + TrackName(CountExport) + Chr(0)
    End If
End Sub

Public Sub ExportCountry()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Country", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + Country(CountExport) + Chr(0)
    End If
End Sub

Public Sub ExportAdjectiv()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Adjective", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(32) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + Adj(CountExport) + Chr(0)
    End If
End Sub

Public Sub ExportTracks()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "TPath", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        SourceFile = Read
        If CountExport + 1 < 10 Then
            Read2 = "0" + Trim(Str(CountExport + 1))
        Else
            Read2 = Trim(Str(CountExport + 1))
        End If
        Read2 = Gp2Dir + "\Circuits\F1ct" + Read2 + ".dat"
        TargetFile = Read2
        If UCase(TargetFile) <> UCase(SourceFile) Then FileCopy SourceFile, TargetFile
    End If
End Sub

Public Sub ExportLength()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Length", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        WareCount = Read
        CountNr = WareCount
        WareCount = WareCount / 78
        If WareCount >= 100 Then Read = Mid(Str(WareCount), 2, 3)
        If (WareCount >= 10) And (WareCount < 100) Then Read = Mid(Str(WareCount), 2, 2)
        If (WareCount >= 1) And (WareCount < 10) Then Read = Mid(Str(WareCount), 2, 1)
        If WareCount < 1 Then Read = 0
        Count2 = Read
        Count1 = (CountNr - (Count2 * 78)) * 3.33333
        If Count1 < 1 Then Count1 = 1
        If Count1 > 255 Then
            Count1 = Count1 - 255
            Count2 = Count2 + 1
        End If
        Read = Chr(Count1)
        Read2 = Chr(Count2)
        Read = Read + Read2
        Put #GP2FileNum, oData.Length(GP2V) + (CountExport * 7), Read
    End If
End Sub

Public Sub ExportWare()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Ware", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        WareCount = Read
        If WareCount < 14848 Then WareCount = 14848
        If WareCount > 37887 Then WareCount = 37887
        CountNr = WareCount
        WareCount = WareCount - 14848
        WareCount = WareCount / 256
        If (WareCount >= 10) Then Read = Mid(Str(WareCount), 2, 2)
        If (WareCount >= 1) And (WareCount < 10) Then Read = Mid(Str(WareCount), 2, 1)
        If WareCount < 1 Then Read = 0
        WareCount = Read

        CountNr = (CountNr - ((256 * WareCount) + 14848))

        Read = ""
        Read = String(1, " ")
        Read2 = String(1, " ")

        If CountNr = 0 Then CountNr = 1
        Read = Chr(CountNr)
        Read2 = Chr(WareCount + 58)
        Read = Read + Read2
        Count2 = oData.Ware + (CountExport * 2)
        Put #GP2FileNum, Count2, Read
    End If
End Sub

Public Sub ExportPoints()
    Read = oMisc.ReadINI("Misc", "Point", ProgramDir + "\WorkCopy.lda")
    X = 1
    Read3 = ""
    Do Until X > 52
        Read2 = Mid(Read, X, 2)
        Read2 = Trim(Read2)
        Count1 = Read2
        Read2 = Chr(Count1)
        Read3 = Read3 + Read2
        X = X + 2
    Loop
    Put #GP2FileNum, oData.Point(GP2V), Read3
End Sub

Public Sub ExportNullAsOne()
    Read = oMisc.ReadINI("Misc", "0as1", ProgramDir + "\WorkCopy.lda")
    If Read = "1" Then Read = "255"
    If Read = "0" Then Read = "254"
    Read = Chr(Read)
    Put #GP2FileNum, oData.OneAsNull(GP2V), Read
End Sub

Public Sub ExportQuickRace()
    Count1 = oMisc.ReadINI("Misc", "Quick", ProgramDir + "\WorkCopy.lda")
    Read = Chr(Count1)
    Put F1SaveFileNum, 648, Read
End Sub

Public Sub ExportSaveLap()
    Read = oMisc.ReadINI("Misc", "SaveLap", ProgramDir + "\WorkCopy.lda")
    If Read = 1 Then
        Read = Chr(100) + Chr(144)
        Put #GP2FileNum, oData.SaveLapTime, Read
        Read = Chr(144) + Chr(144)
        Put #GP2FileNum, oData.SaveLapTime2, Read
    Else
        Read2 = ""
        Read = Chr(92)
        Read2 = Read
        Read = Chr(114)
        Read2 = Read2 + Read
        
        Put #GP2FileNum, oData.SaveLapTime, Read2
        Read2 = ""
        Read = Chr(114)
        Read2 = Read
        Read = Chr(92)
        Read2 = Read2 + Read
        Put #GP2FileNum, oData.SaveLapTime2, Read2
    End If
End Sub

Public Sub ExportLevel()
    Read = oMisc.ReadINI("Misc", "Year", ProgramDir + "\WorkCopy.lda")
    Put #GP2FileNum, oData.Level(GP2V), Read
End Sub

Public Sub ExportCarHelp()
    X = 1
    Count1 = 0
    Count2 = 0
    Read = oMisc.ReadINI("Misc", "Aids", ProgramDir + "\WorkCopy.lda")
    CountNr = 0
    Count3 = 1
    Do Until CountNr > 4
        Read2 = Mid(Read, Count3, 7)
        Do Until X > 7
            Read3 = Mid(Read2, X, 1)
            If Read3 = 1 Then
                If X = 1 Then
                    Count1 = Count1 + 2
                End If
                If X = 2 Then
                    Count1 = Count1 + 3
                End If
                If X = 3 Then
                    Count1 = Count1 + 5
                End If
                If X = 4 Then
                    Count1 = Count1 + 9
                End If
                If X = 5 Then
                    Count1 = Count1 + 17
                End If
                If X = 6 Then
                    Count1 = Count1 + 33
                End If
                If X = 7 Then
                    Count1 = Count1 + 65
                End If
                Count2 = Count2 + 1
            End If
            X = X + 1
            Count3 = Count3 + 1
        Loop
        Count1 = Count1 - Count2
        Read4 = Chr(Count1)
        Put #GP2FileNum, oData.Help + CountNr, Read4
        Count2 = 0
        CountNr = CountNr + 1
        Count1 = 0
        X = 1
    Loop
End Sub

Public Sub ExportPQPower()
    Read = oMisc.ReadINI("Player", "QPower", ProgramDir + "\WorkCopy.lda")
    TempDouble = Read
    TempDouble = TempDouble / 256
    Read2 = Mid(TempDouble, 1, 1)
    X = Read
    X = X - (Read2 * 256)
    Read = Chr(X) + Chr(Read2)
    Put #GP2FileNum, oData.PQPower, Read
End Sub

Public Sub ExportPRPower()
    Read = oMisc.ReadINI("Player", "RPower", ProgramDir + "\WorkCopy.lda")
    TempDouble = Read
    TempDouble = TempDouble / 256
    Read2 = Mid(TempDouble, 1, 1)
    X = Read
    X = X - (Read2 * 256)
    Read = Chr(X) + Chr(Read2)
    Put #GP2FileNum, oData.PRPower, Read
End Sub

Public Sub ExportPGrip()
    Read = oMisc.ReadINI("Player", "Grip", ProgramDir + "\WorkCopy.lda")
    TempDouble = Read
    TempDouble = TempDouble / 256
    Read2 = Mid(TempDouble, 1, 1)
    X = Read - (Read2 * 256)
    Read = Chr(X) + Chr(Read2)
    Put #GP2FileNum, oData.PGrip, Read
End Sub

Public Sub ExportPWeight()
    Read = oMisc.ReadINI("Player", "Weight", ProgramDir + "\WorkCopy.lda")
    TempDouble = Read
    TempDouble = TempDouble / 256
    Read2 = Mid(TempDouble, 2, 1)
    If Read2 = "." Then
        Read2 = Mid(TempDouble, 1, 1)
    Else
        Read2 = Mid(TempDouble, 1, 1) + Read2
    End If
    X = Read
    X = X - (256 * Read2)
    Read = Chr(X) + Chr(Read2)
    Put #GP2FileNum, oData.PWeight, Read
End Sub

Public Sub ExportCWeight()
    Read = oMisc.ReadINI("Misc", "CWeight", ProgramDir + "\WorkCopy.lda")
    TempDouble = Read
    TempDouble = TempDouble / 256
    Read2 = Mid(TempDouble, 2, 1)
    If Read2 = "." Then
        Read2 = Mid(TempDouble, 1, 1)
    Else
        Read2 = Mid(TempDouble, 1, 1) + Read2
    End If
    X = Read
    X = X - (256 * Read2)
    Read = Chr(X) + Chr(Read2)
    Put #GP2FileNum, oData.CWeight, Read
End Sub

Public Sub ExportSpeed()
    Read = oMisc.ReadINI("Player", "Speed", ProgramDir + "\WorkCopy.lda")
    If Read = 1 Then
        Read = Chr(235)
        Put #GP2FileNum, oData.NoPitSpeed, Read
    Else
        Read = Chr(116)
        Put #GP2FileNum, oData.NoPitSpeed, Read
    End If
    Read = oMisc.ReadINI("Player", "Speed", ProgramDir + "\WorkCopy.lda")
    CountNr = 204
    Count2 = 2
    Count1 = 0
    Do Until Count1 > Read - 2
        CountNr = CountNr + 68
        If CountNr > 255 Then
            CountNr = CountNr - 256
            Count2 = Count2 + 1
        End If
        Count2 = Count2 + 1
        Count1 = Count1 + 1
    Loop
    Read = Chr(CountNr) + Chr(Count2)
    Put #GP2FileNum, oData.PitSpeed, Read
End Sub

Public Sub ExportUseTeam()
    Read = oMisc.ReadINI("Player", "UseTeam", ProgramDir + "\WorkCopy.lda")
    If Read = 0 Then Read = 255
    If Read = 1 Then Read = 0
    Read = Chr(Read)
    Put #GP2FileNum, oData.UseTeam, Read
End Sub

Public Sub ExportPictures()
    FileNum = FreeFile
    Open Gp2Dir + "\_menupic.bat" For Append As FileNum
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "BPic", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Read = oMisc.GetShortName(Read)
        Read2 = oMisc.GetShortName(Gp2Dir)
        Read = Read2 + "\gp2hipic.exe -q #" + Trim(Str(CountExport + 1)) + " " + Read2 + "\bitmaps\f1pcsvga.bin " + Read
        Print #FileNum, LCase(Read)
    End If
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "SPic", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Read = oMisc.GetShortName(Read)
        Read2 = oMisc.GetShortName(Gp2Dir)
        Read = Read2 + "\gp2hipic.exe -q #" + Trim(Str(CountExport + 17)) + " " + Read2 + "\bitmaps\f1pcsvga.bin " + Read
        Print #FileNum, LCase(Read)
    End If
    If CountExport + 1 = 16 Then
        Read2 = oMisc.GetShortName(Gp2Dir)
        Read = Read2 + "\gp2hipic.exe -d " + Read2 + "\bitmaps\f1pcsvga.bin"
        Print #FileNum, LCase(Read)
    End If
    Close FileNum
End Sub

Public Sub ExportRTeam()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "RTeam", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Dim TempInteger As Integer
        TempInteger = Len(Read)
        Read2 = Chr(0)
        Read2 = String(12 - TempInteger, Read2)
        Read = Read & Read2
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        Temp = 718 + (CountExport * 88)
        Put #FileNum, Temp, Read
        Close FileNum
    End If
End Sub

Public Sub ExportQTeam()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "QTeam", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Dim TempInteger As Integer
        TempInteger = Len(Read)
        Read2 = Chr(0)
        Read2 = String(12 - TempInteger, Read2)
        Read = Read & Read2
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        Temp = 674 + (CountExport * 88)
        Put #FileNum, Temp, Read
        Close FileNum
    End If
End Sub

Public Sub ExportRName()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "RDriver", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Dim TempInteger As Integer
        TempInteger = Len(Read)
        Read2 = Chr(0)
        Read2 = String(22 - TempInteger, Read2)
        Read = Read & Read2
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        Temp = 694 + (CountExport * 88)
        Put #FileNum, Temp, Read
        Close FileNum
    End If
End Sub

Public Sub ExportQName()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "QDriver", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Dim TempInteger As Integer
        TempInteger = Len(Read)
        Read2 = Chr(0)
        Read2 = String(22 - TempInteger, Read2)
        Read = Read & Read2
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        Temp = 650 + (CountExport * 88)
        Put #FileNum, Temp, Read
        Close FileNum
    End If
End Sub

Public Sub ExportRaceTime()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "RTime", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Read2 = Mid(Read, 1, 1)
        Read3 = Mid(Read, 3, 2)
        Read4 = Mid(Read, 6, 3)
        X = Read2
        X = X * 60
        TempDouble = X + Read3 + (Read4 / 1000)
        Count1 = 1
        Do Until TempDouble < 65.536
            TempDouble = TempDouble - 65.536
            Count1 = Count1 + 1
        Loop
        Count2 = 1
        Do Until TempDouble < 0.256
            TempDouble = TempDouble - 0.256
            Count2 = Count2 + 1
        Loop
        X = TempDouble * 1000 + 1
        Read2 = ""
        Read = Chr(X - 1)
        Read2 = Read2 & Read
        Read = Chr(Count2 - 1)
        Read2 = Read2 & Read
        Read = Chr(Count1 - 1)
        Read2 = Read2 & Read
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        Temp = 732 + (CountExport * 88)
        Put #FileNum, Temp, Read2
        Close FileNum
    End If
End Sub

Public Sub ExportQualTime()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "QTime", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        Read2 = Mid(Read, 1, 1)
        Read3 = Mid(Read, 3, 2)
        Read4 = Mid(Read, 6, 3)
        X = Read2
        X = X * 60
        TempDouble = X + Read3 + (Read4 / 1000)
        Count1 = 1
        Do Until TempDouble < 65.536
            TempDouble = TempDouble - 65.536
            Count1 = Count1 + 1
        Loop
        Count2 = 1
        Do Until TempDouble < 0.256
            TempDouble = TempDouble - 0.256
            Count2 = Count2 + 1
        Loop
        X = TempDouble * 1000 + 1
        Read2 = ""
        Read = Chr(X - 1)
        Read2 = Read2 + Read
        Read = Chr(Count2 - 1)
        Read2 = Read2 + Read
        Read = Chr(Count1 - 1)
        Read2 = Read2 + Read
        Close FileNum
        
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        Temp = 688 + (CountExport * 88)
        Put #FileNum, Temp, Read2
        Close FileNum
    End If
End Sub

Public Sub ExportQDate()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "QDate", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        TheDate = Read
        Read = "1978-01-01"
        Read = DateDiff("d", Read, TheDate)
        TempDouble = Read
        If TempDouble < 0 Then Exit Sub
        TempDouble = TempDouble / 256
        If (TempDouble >= 10) And (TempDouble <= 100) Then
            Read2 = TempDouble
            Read2 = Mid(Read2, 1, 2)
            Count1 = Read2
        ElseIf TempDouble >= 100 Then
            Read2 = TempDouble
            Read2 = Mid(Read2, 1, 3)
            Count1 = Read2
        ElseIf (TempDouble < 10) And (TempDouble >= 1) Then
            Read2 = TempDouble
            Read2 = Mid(Read2, 1, 1)
            Count1 = Read2
        ElseIf TempDouble < 1 Then
            Count1 = 1
        End If
        TempDouble = Read
        TempDouble = TempDouble - ((Count1) * 256)
        Read = Chr(Count1)
        Read2 = Chr(TempDouble)
        Read = Read2 & Read
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        Temp = 692 + (CountExport * 88)
        Put #FileNum, Temp, Read
        Close FileNum
    End If
End Sub

Public Sub ExportRDate()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "RDate", ProgramDir + "\WorkCopy.lda")
    If Read <> "" Then
        TheDate = Read
        Read = "1978-01-01"
        Read = DateDiff("d", Read, TheDate)
        TempDouble = Read
        If TempDouble < 0 Then Exit Sub
        TempDouble = TempDouble / 256
        If (TempDouble >= 10) And (TempDouble <= 100) Then
            Read2 = TempDouble
            Read2 = Mid(Read2, 1, 2)
            Count1 = Read2
        ElseIf TempDouble >= 100 Then
            Read2 = TempDouble
            Read2 = Mid(Read2, 1, 3)
            Count1 = Read2
        ElseIf (TempDouble < 10) And (TempDouble >= 1) Then
            Read2 = TempDouble
            Read2 = Mid(Read2, 1, 1)
            Count1 = Read2
        ElseIf TempDouble < 1 Then
            Count1 = 1
        End If
        TempDouble = Read
        TempDouble = TempDouble - ((Count1) * 256)
        Read = Chr(Count1)
        Read2 = Chr(TempDouble)
        Read = Read2 & Read
        FileNum = FreeFile
        Open Gp2Dir + "\f1gstate.sav" For Binary As FileNum
        Temp = 736 + (CountExport * 88)
        Put #FileNum, Temp, Read
        Close FileNum
    End If
End Sub

Public Sub ExportDos()
    FileNum = FreeFile
    Open Gp2Dir + "\_menupic.bat" For Append As FileNum
    Read = oMisc.ReadINI("Misc", "EXEPath", ProgramDir + "\WorkCopy.lda")
    Read2 = oMisc.GetShortName(Read)
    Read = oMisc.GetShortName(Gp2Dir)
    Read3 = oMisc.ReadINI("Misc", "EXE", ProgramDir + "\WorkCopy.lda")
    If Read3 = "" Then
        Read = Read2 & " " & Read
        Print #FileNum, Read
    Else
        Read = Read2 & " " & Read & " " & Read3
        Print #FileNum, Read
    End If
    Close FileNum
End Sub
