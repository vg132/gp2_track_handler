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
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "Laps", TempFile)
    If Read <> "" Then
        tExp.bByte = Read
        If tExp.bByte > 126 Then tExp.bByte = 126
        If tExp.bByte < 3 Then tExp.bByte = 3
        Put #Exp.GP2FileNum, oData.Laps(GP2V) + Exp.TrackNr, tExp.bByte
    End If
End Sub

Public Sub ExportName()
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "Name", TempFile)
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + TrackName(Exp.TrackNr) + Chr(0)
    End If
End Sub

Public Sub ExportCountry()
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "Country", TempFile)
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + Country(Exp.TrackNr) + Chr(0)
    End If
End Sub

Public Sub ExportAdjectiv()
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "Adjective", TempFile)
    If Read <> "" Then
        GP2NameFile = GP2NameFile + Trim(Read) + Chr(32) + Chr(0)
    Else
        GP2NameFile = GP2NameFile + Adj(Exp.TrackNr) + Chr(0)
    End If
End Sub

Public Sub ExportTracks()
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "TPath", TempFile)
    If Read <> "" Then
        Read2 = oMisc.File_Exists(Read)
        If Read2 = True Then
            SourceFile = Read
            If Exp.TrackNr + 1 < 10 Then
                Read2 = "0" & Exp.TrackNr + 1
            Else
                Read2 = Exp.TrackNr + 1
            End If
            Read2 = GP2Dir & "\Circuits\F1ct" & Read2 & ".dat"
            TargetFile = Read2
            If UCase(TargetFile) <> UCase(SourceFile) Then FileCopy SourceFile, TargetFile
        Else
            MsgBox "Track Nr " & Exp.TrackNr + 1 & ", " & Read & ", was not found." & vbLf & vbLf & _
                    "If you have moved this track to a diffrent path then remove this" & vbLf & _
                    "track from this track set and add it once more from the new location.", vbCritical, TH
        End If
    End If
End Sub

Public Sub ExportLength()
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "Length", TempFile)
    If Read <> "" Then
        tExp.lLong = Read
        TempDouble = tExp.lLong * 3.28212677519917
        tExp.lLong = Round(TempDouble, 0)
        If tExp.lLong > 32767 Then tExp.lLong = tExp.lLong - 65535
        tExp.iInt = tExp.lLong
        Put #Exp.GP2FileNum, oData.Length(GP2V) + (Exp.TrackNr * 7), tExp.iInt
    End If
End Sub

Public Sub ExportWare()
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "Ware", TempFile)
    If Read <> "" Then
        tExp.lLong = Read
        If tExp.lLong > 32767 Then
            tExp.lLong = tExp.lLong - 65535
        End If
        tExp.iInt = tExp.lLong
        Put #Exp.GP2FileNum, oData.Ware + (Exp.TrackNr * 2), tExp.iInt
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
    Put #Exp.GP2FileNum, oData.Point(GP2V), Read3
End Sub

Public Sub ExportNullAsOne()
    tExp.bByte = oMisc.ReadINI("Misc", "0as1", TempFile)
    If tExp.bByte = "1" Then tExp.bByte = "255"
    If tExp.bByte = "0" Then tExp.bByte = "254"
    Put #Exp.GP2FileNum, oData.OneAsNull(GP2V), tExp.bByte
End Sub

Public Sub ExportQuickRace()
    tExp.bByte = oMisc.ReadINI("Misc", "Quick", TempFile)
    Put #Exp.F1FileNum, 648, tExp.bByte
End Sub

Public Sub ExportSaveLap()
    Read = oMisc.ReadINI("Misc", "SaveLap", TempFile)
    If Read = 1 Then
        Read = Chr(100) + Chr(144)
        Put #Exp.GP2FileNum, oData.SaveLapTime, Read
        Read = Chr(144) + Chr(144)
        Put #Exp.GP2FileNum, oData.SaveLapTime2, Read
    Else
        Read2 = ""
        Read = Chr(92)
        Read2 = Read
        Read = Chr(114)
        Read2 = Read2 + Read
        
        Put #Exp.GP2FileNum, oData.SaveLapTime, Read2
        Read2 = ""
        Read = Chr(114)
        Read2 = Read
        Read = Chr(92)
        Read2 = Read2 + Read
        Put #Exp.GP2FileNum, oData.SaveLapTime2, Read2
    End If
End Sub

Public Sub ExportLevel()
    tExp.Year = oMisc.ReadINI("Misc", "Year", TempFile)
    If tExp.Year <> "" Then Put #Exp.GP2FileNum, oData.Level(GP2V), tExp.Year
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
        Put #Exp.GP2FileNum, oData.Help + CountNr, tExp.bByte
        Count2 = 0
        Count1 = 0
        X = 1
    Next
End Sub

Public Sub ExportPQPower()
    Read = oMisc.ReadINI("Player", "QPower", TempFile)
    If Read <> "" Then
        tExp.iInt = Read
        Put #Exp.GP2FileNum, oData.PQPower, tExp.iInt
    End If
End Sub

Public Sub ExportPRPower()
    Read = oMisc.ReadINI("Player", "RPower", TempFile)
    If Read <> "" Then
        tExp.iInt = Read
        Put #Exp.GP2FileNum, oData.PRPower, tExp.iInt
    End If
End Sub

Public Sub ExportPGrip()
    Read = oMisc.ReadINI("Player", "Grip", TempFile)
    TempDouble = Read
    TempDouble = TempDouble / 256
    Read2 = Mid(TempDouble, 1, 1)
    X = Read - (Read2 * 256)
    Read = Chr(X) + Chr(Read2)
    Put #Exp.GP2FileNum, oData.PGrip, Read
End Sub

Public Sub ExportPWeight()
    Read = oMisc.ReadINI("Player", "Weight", TempFile)
    If Read <> "" Then
        tExp.iInt = Read
        Put #Exp.GP2FileNum, oData.PWeight, tExp.iInt
    End If
End Sub

Public Sub ExportCWeight()
    Read = oMisc.ReadINI("Misc", "CWeight", TempFile)
    If Read <> "" Then
        tExp.iInt = Read
        Put #Exp.GP2FileNum, oData.CWeight, tExp.iInt
    End If
End Sub

Public Sub ExportSpeed()
    Read = oMisc.ReadINI("Player", "NoSpeed", TempFile)
    If Read = 1 Then
        tExp.bByte = 235
        Put #Exp.GP2FileNum, oData.NoPitSpeed, tExp.bByte
    Else
        tExp.bByte = 116
        Put #Exp.GP2FileNum, oData.NoPitSpeed, tExp.bByte
    End If
    Read = oMisc.ReadINI("Player", "Speed", TempFile)
    tExp.lLong = Read
    tExp.lLong = (tExp.lLong * 324) + 392
    Put #Exp.GP2FileNum, oData.PitSpeed, tExp.lLong
End Sub

Public Sub ExportUseTeam()
    tExp.bByte = oMisc.ReadINI("Player", "UseTeam", TempFile)
    If tExp.bByte = 0 Then tExp.bByte = 255
    If tExp.bByte = 1 Then tExp.bByte = 0
    Put #Exp.GP2FileNum, oData.UseTeam, tExp.bByte
End Sub

Public Sub ExportPictures()
    FileNum = FreeFile
    Open ProgramDir & "\Bat\Export.bat" For Append As FileNum
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "BPic", TempFile)
    If Read <> "" Then
        Read = oMisc.GetShortName(Read)
        Read2 = oMisc.GetShortName(GP2Dir)
        Read = Read2 & "\gp2hipic.exe -q #" & Exp.TrackNr + 1 & " " & Read2 & "\bitmaps\f1pcsvga.bin " & Read
        Print #FileNum, LCase(Read)
    End If
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "SPic", TempFile)
    If Read <> "" Then
        Read = oMisc.GetShortName(Read)
        Read2 = oMisc.GetShortName(GP2Dir)
        Read = Read2 & "\gp2hipic.exe -q #" & Exp.TrackNr + 17 & " " & Read2 & "\bitmaps\f1pcsvga.bin " & Read
        Print #FileNum, LCase(Read)
    End If
    If Exp.TrackNr + 1 = 16 Then
        Read2 = oMisc.GetShortName(GP2Dir)
        Read = Read2 & "\gp2hipic.exe -d " & Read2 & "\bitmaps\f1pcsvga.bin"
        Print #FileNum, LCase(Read)
    End If
    Close FileNum
End Sub

Public Sub ExportRTeam(ByVal Rec As RecEnum)
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "RTeam", TempFile)
    If Read <> "" Then
        Var.iInt1 = Len(Read)
        Read2 = Chr(0)
        Read2 = String(12 - Var.iInt1, Read2)
        Read = Read & Read2
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 718 + (Exp.TrackNr * 88), Read
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 101 + (Exp.TrackNr * 88), Read
        End If
    End If
End Sub

Public Sub ExportQTeam(ByVal Rec As RecEnum)
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "QTeam", TempFile)
    If Read <> "" Then
        Var.iInt1 = Len(Read)
        Read2 = Chr(0)
        Read2 = String(12 - Var.iInt1, Read2)
        Read = Read & Read2
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 674 + (Exp.TrackNr * 88), Read
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 57 + (Exp.TrackNr * 88), Read
        End If
    End If
End Sub

Public Sub ExportRName(ByVal Rec As RecEnum)
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "RDriver", TempFile)
    If Read <> "" Then
        Var.iInt1 = Len(Read)
        Read2 = Chr(0)
        Read2 = String(22 - Var.iInt1, Read2)
        Read = Read & Read2
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 694 + (Exp.TrackNr * 88), Read
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 77 + (Exp.TrackNr * 88), Read
        End If
    End If
End Sub

Public Sub ExportQName(ByVal Rec As RecEnum)
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "QDriver", TempFile)
    If Read <> "" Then
        Var.iInt1 = Len(Read)
        Read2 = Chr(0)
        Read2 = String(22 - Var.iInt1, Read2)
        Read = Read & Read2
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 650 + (Exp.TrackNr * 88), Read
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 33 + (Exp.TrackNr * 88), Read
        End If
    End If
End Sub

Public Sub ExportTime(ByVal QR As ImpExpTime, ByVal Rec As RecEnum)
    If QR = Qual Then
        Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "QTime", TempFile)
    Else
        Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "RTime", TempFile)
    End If
    If Read <> "" Then
        Count1 = Mid(Read, 1, InStr(1, Read, ":") - 1)
        Count2 = Mid(Read, InStr(1, Read, ":") + 1)
        Count1 = Count1 * 60000
        tExp.lLong = Count2 + Count1
        If QR = Qual Then
            If Rec = F1gstate Then
                Put #Exp.F1FileNum, 688 + (Exp.TrackNr * 88), tExp.lLong
            ElseIf Rec = RecFile Then
                Put #Exp.F1FileNum, 71 + (Exp.TrackNr * 88), tExp.lLong
            End If
        Else
            If Rec = F1gstate Then
            Put #Exp.F1FileNum, 732 + (Exp.TrackNr * 88), tExp.lLong
            ElseIf Rec = RecFile Then
                Put #Exp.F1FileNum, 115 + (Exp.TrackNr * 88), tExp.lLong
            End If
        End If
    End If
End Sub

Public Sub ExportQDate(ByVal Rec As RecEnum)
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "QDate", TempFile)
    If Read <> "" Then
        tExp.lLong = DateDiff("d", "1978-01-01", Read)
        If tExp.lLong < 0 Then tExp.lLong = 0
        If tExp.lLong > 32767 Then
            tExp.iInt = tExp.lLong - 65535
        Else
            tExp.iInt = tExp.lLong
        End If
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 692 + (Exp.TrackNr * 88), tExp.iInt
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 75 + (Exp.TrackNr * 88), tExp.iInt
        End If
    End If
End Sub

Public Sub ExportRDate(ByVal Rec As RecEnum)
    Read = oMisc.ReadINI("Track " & Exp.TrackNr + 1, "RDate", TempFile)
    If Read <> "" Then
        tExp.lLong = DateDiff("d", "1978-01-01", Read)
        If tExp.lLong < 0 Then tExp.lLong = 0
        If tExp.lLong > 32767 Then
            tExp.iInt = tExp.lLong - 65535
        Else
            tExp.iInt = tExp.lLong
        End If
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 736 + (Exp.TrackNr * 88), tExp.iInt
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 119 + (Exp.TrackNr * 88), tExp.iInt
        End If
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
