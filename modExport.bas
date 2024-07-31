Attribute VB_Name = "modExport"
Public Sub SetAttribut()
    On Error Resume Next
    For X = 1 To 16
        Read = Str(X)
        Read = Trim(Read)
        If Len(Read) < 2 Then Read = "0" + Read
        Read = GP2Dir + "\Circuits\f1ct" + Read + ".dat"
        SetAttr Read, vbNormal
    Next
    Read = GP2Dir + "\gp2.exe"
    SetAttr Read, vbNormal
End Sub

Public Sub ExportLaps()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Laps", TempFile)
    If Read <> "" Then
        tExp.bByte = Read
        If tExp.bByte > 126 Then tExp.bByte = 126
        If tExp.bByte < 3 Then tExp.bByte = 3
        Put #GP2FileNum, oData.Laps(GP2V) + CountExport, tExp.bByte
    End If
End Sub

Public Sub ExportName()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Name", TempFile)
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + TrackName(CountExport) + Chr(0)
    End If
End Sub

Public Sub ExportCountry()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Country", TempFile)
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + Country(CountExport) + Chr(0)
    End If
End Sub

Public Sub ExportAdjectiv()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Adjective", TempFile)
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(32) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + Adj(CountExport) + Chr(0)
    End If
End Sub

Public Sub ExportTracks()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "TPath", TempFile)
    If Read <> "" Then
        SourceFile = Read
        If CountExport + 1 < 10 Then
            Read2 = "0" + Trim(Str(CountExport + 1))
        Else
            Read2 = Trim(Str(CountExport + 1))
        End If
        Read2 = GP2Dir + "\Circuits\F1ct" + Read2 + ".dat"
        TargetFile = Read2
        If UCase(TargetFile) <> UCase(SourceFile) Then FileCopy SourceFile, TargetFile
    End If
End Sub

Public Sub ExportLength()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Length", TempFile)
    If Read <> "" Then
        tExp.lLong = Read
        TempDouble = tExp.lLong * 3.28212677519917
        tExp.lLong = Round(TempDouble, 0)
        If tExp.lLong > 32767 Then tExp.lLong = tExp.lLong - 65535
        tExp.iInt = tExp.lLong
        Put #GP2FileNum, oData.Length(GP2V) + (CountExport * 7), tExp.iInt
    End If
End Sub

Public Sub ExportWare()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "Ware", TempFile)
    If Read <> "" Then
        tExp.lLong = Read
        If tExp.lLong > 32767 Then
            tExp.lLong = tExp.lLong - 65535
        End If
        tExp.iInt = tExp.lLong
        Put #GP2FileNum, oData.Ware + (CountExport * 2), tExp.iInt
    End If
End Sub

Public Sub ExportPoints()
    Read = oMisc.ReadINI("Misc", "Point", TempFile)
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
    tExp.bByte = oMisc.ReadINI("Misc", "0as1", TempFile)
    If tExp.bByte = "1" Then tExp.bByte = "255"
    If tExp.bByte = "0" Then tExp.bByte = "254"
    Put #GP2FileNum, oData.OneAsNull(GP2V), tExp.bByte
End Sub

Public Sub ExportQuickRace()
    tExp.bByte = oMisc.ReadINI("Misc", "Quick", TempFile)
    Put F1SaveFileNum, 648, tExp.bByte
End Sub

Public Sub ExportSaveLap()
    Read = oMisc.ReadINI("Misc", "SaveLap", TempFile)
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
    tExp.Year = oMisc.ReadINI("Misc", "Year", TempFile)
    If tExp.Year <> "" Then Put #GP2FileNum, oData.Level(GP2V), tExp.Year
End Sub

Public Sub ExportCarHelp()
Dim CountNr As Integer
    X = 1
    Count1 = 0
    Count2 = 0
    Read = oMisc.ReadINI("Misc", "Aids", TempFile)
    CountNr = 0
    Count3 = 1
    For CountNr = 0 To 4
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
        tExp.bByte = Count1
        Put #GP2FileNum, oData.Help + CountNr, tExp.bByte
        Count2 = 0
        Count1 = 0
        X = 1
    Next
End Sub

Public Sub ExportPQPower()
    Read = oMisc.ReadINI("Player", "QPower", TempFile)
    If Read <> "" Then
        tExp.iInt = Read
        Put #GP2FileNum, oData.PQPower, tExp.iInt
    End If
End Sub

Public Sub ExportPRPower()
    Read = oMisc.ReadINI("Player", "RPower", TempFile)
    If Read <> "" Then
        tExp.iInt = Read
        Put #GP2FileNum, oData.PRPower, tExp.iInt
    End If
End Sub

Public Sub ExportPGrip()
    Read = oMisc.ReadINI("Player", "Grip", TempFile)
    TempDouble = Read
    TempDouble = TempDouble / 256
    Read2 = Mid(TempDouble, 1, 1)
    X = Read - (Read2 * 256)
    Read = Chr(X) + Chr(Read2)
    Put #GP2FileNum, oData.PGrip, Read
End Sub

Public Sub ExportPWeight()
    Read = oMisc.ReadINI("Player", "Weight", TempFile)
    If Read <> "" Then
        tExp.iInt = Read
        Put #GP2FileNum, oData.PWeight, tExp.iInt
    End If
End Sub

Public Sub ExportCWeight()
    Read = oMisc.ReadINI("Misc", "CWeight", TempFile)
    If Read <> "" Then
        tExp.iInt = Read
        Put #GP2FileNum, oData.CWeight, tExp.iInt
    End If
End Sub

Public Sub ExportSpeed()
    Read = oMisc.ReadINI("Player", "NoSpeed", TempFile)
    If Read = 1 Then
        tExp.bByte = 235
        Put #GP2FileNum, oData.NoPitSpeed, tExp.bByte
    Else
        tExp.bByte = 116
        Put #GP2FileNum, oData.NoPitSpeed, tExp.bByte
    End If
    Read = oMisc.ReadINI("Player", "Speed", TempFile)
    tExp.lLong = Read
    tExp.lLong = (tExp.lLong * 324) + 392
    Put #GP2FileNum, oData.PitSpeed, tExp.lLong
End Sub

Public Sub ExportUseTeam()
    tExp.bByte = oMisc.ReadINI("Player", "UseTeam", TempFile)
    If tExp.bByte = 0 Then tExp.bByte = 255
    If tExp.bByte = 1 Then tExp.bByte = 0
    Put #GP2FileNum, oData.UseTeam, tExp.bByte
End Sub

Public Sub ExportPictures()
    FileNum = FreeFile
    Open ProgramDir & "\Bat\Export.bat" For Append As FileNum
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "BPic", TempFile)
    If Read <> "" Then
        Read = oMisc.GetShortName(Read)
        Read2 = oMisc.GetShortName(GP2Dir)
        Read = Read2 + "\gp2hipic.exe -q #" + Trim(Str(CountExport + 1)) + " " + Read2 + "\bitmaps\f1pcsvga.bin " + Read
        Print #FileNum, LCase(Read)
    End If
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "SPic", TempFile)
    If Read <> "" Then
        Read = oMisc.GetShortName(Read)
        Read2 = oMisc.GetShortName(GP2Dir)
        Read = Read2 + "\gp2hipic.exe -q #" + Trim(Str(CountExport + 17)) + " " + Read2 + "\bitmaps\f1pcsvga.bin " + Read
        Print #FileNum, LCase(Read)
    End If
    If CountExport + 1 = 16 Then
        Read2 = oMisc.GetShortName(GP2Dir)
        Read = Read2 + "\gp2hipic.exe -d " + Read2 + "\bitmaps\f1pcsvga.bin"
        Print #FileNum, LCase(Read)
    End If
    Close FileNum
End Sub

Public Sub ExportRTeam()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "RTeam", TempFile)
    If Read <> "" Then
        Dim TempInteger As Integer
        TempInteger = Len(Read)
        Read2 = Chr(0)
        Read2 = String(12 - TempInteger, Read2)
        Read = Read & Read2
        Put #F1SaveFileNum, 718 + (CountExport * 88), Read
    End If
End Sub

Public Sub ExportQTeam()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "QTeam", TempFile)
    If Read <> "" Then
        Dim TempInteger As Integer
        TempInteger = Len(Read)
        Read2 = Chr(0)
        Read2 = String(12 - TempInteger, Read2)
        Read = Read & Read2
        Put #F1SaveFileNum, 674 + (CountExport * 88), Read
    End If
End Sub

Public Sub ExportRName()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "RDriver", TempFile)
    If Read <> "" Then
        Dim TempInteger As Integer
        TempInteger = Len(Read)
        Read2 = Chr(0)
        Read2 = String(22 - TempInteger, Read2)
        Read = Read & Read2
        Put #F1SaveFileNum, 694 + (CountExport * 88), Read
    End If
End Sub

Public Sub ExportQName()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "QDriver", TempFile)
    If Read <> "" Then
        Dim TempInteger As Integer
        TempInteger = Len(Read)
        Read2 = Chr(0)
        Read2 = String(22 - TempInteger, Read2)
        Read = Read & Read2
        Put #F1SaveFileNum, 650 + (CountExport * 88), Read
    End If
End Sub

Public Sub ExportTimeToGP2(ByVal QR As ImpExpTime)
    If QR = Qual Then
        Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "QTime", TempFile)
    Else
        Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "RTime", TempFile)
    End If
    If Read <> "" Then
        Count1 = Mid(Read, 1, InStr(1, Read, ":") - 1)
        Count2 = Mid(Read, InStr(1, Read, ":") + 1)
        Count1 = Count1 * 60000
        tExp.lLong = Count2 + Count1
        If QR = Qual Then
            Put #F1SaveFileNum, 688 + (CountExport * 88), tExp.lLong
        Else
            Put #F1SaveFileNum, 732 + (CountExport * 88), tExp.lLong
        End If
    End If
End Sub

Public Sub ExportQDate()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "QDate", TempFile)
    If Read <> "" Then
        tExp.lLong = DateDiff("d", "1978-01-01", Read)
        If tExp.lLong < 0 Then tExp.lLong = 0
        If tExp.lLong > 32767 Then
            tExp.iInt = tExp.lLong - 65535
        Else
            tExp.iInt = tExp.lLong
        End If
        Put #F1SaveFileNum, 692 + (CountExport * 88), tExp.iInt
    End If
End Sub

Public Sub ExportRDate()
    Read = oMisc.ReadINI("Track " + Trim(Str(CountExport + 1)), "RDate", TempFile)
    If Read <> "" Then
        tExp.lLong = DateDiff("d", "1978-01-01", Read)
        If tExp.lLong < 0 Then tExp.lLong = 0
        If tExp.lLong > 32767 Then
            tExp.iInt = tExp.lLong - 65535
        Else
            tExp.iInt = tExp.lLong
        End If
        Put #F1SaveFileNum, 736 + (CountExport * 88), tExp.iInt
    End If
End Sub

Public Sub ExportDos()
    FileNum = FreeFile
    Open ProgramDir & "\Bat\Export.bat" For Append As FileNum
    Read = oMisc.ReadINI("Misc", "EXEPath", TempFile)
    Read2 = oMisc.GetShortName(Read)
    Read = oMisc.GetShortName(GP2Dir)
    Read3 = oMisc.ReadINI("Misc", "EXE", TempFile)
    If Read3 = "" Then
        Read = Read2 & " " & Read
        Print #FileNum, Read
    Else
        Read = Read2 & " " & Read & " " & Read3
        Print #FileNum, Read
    End If
    Close FileNum
End Sub
