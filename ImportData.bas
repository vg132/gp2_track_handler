Attribute VB_Name = "ImportTrack"
Public Sub ImportLength()
    CountNr = 0
    Do Until CountNr > 15
        Read = String(2, " ")
        Get #GP2FileNum, oData.Length(GP2V) + (CountNr * 7), Read
        Read2 = Mid(Read, 1, 1)
        Count1 = Asc(Read2)
        Read2 = Mid(Read, 2, 1)
        Count2 = Asc(Read2)
        Count2 = (Count1 / 3.33333) + (Count2 * 78)
        Count2 = oMisc.WriteINI("Track " + Trim(Str(CountNr + 1)), "Length", Trim(Str(Count2)), ProgramDir + "\WorkCopy.lda")
        CountNr = CountNr + 1
    Loop
End Sub

Public Sub ImportLaps()
    Count1 = 0
    Do Until Count1 > 15
        Read = String(1, " ")
        Get #GP2FileNum, oData.Laps(GP2V) + Count1, Read
        Count2 = Asc(Read)
        Count2 = oMisc.WriteINI("Track " + Trim(Str(Count1 + 1)), "Laps", Trim(Str(Count2)), ProgramDir + "\WorkCopy.lda")
        Count1 = Count1 + 1
    Loop
End Sub

Public Sub ImportWare()
    X = 0
    Y = 1
    Do Until X > 31
        Read = String(2, " ")
        Read2 = String(1, " ")
        Get #GP2FileNum, oData.Ware + X, Read
        CountNr = 0
        Read3 = Mid(Read, 1, 1)
        CountNr = Asc(Read3)
        Read3 = Mid(Read, 2, 1)
        Count2 = 0
        Count2 = Asc(Read3)
        Count2 = (((Count2 - 58) * 256) + 14848)
        CountNr = (CountNr) + Count2
        X = X + 2
        CountNr = oMisc.WriteINI("Track " + Trim(Str(Y)), "Ware", Trim(Str(CountNr)), ProgramDir + "\WorkCopy.lda")
        Y = Y + 1
    Loop
End Sub

Public Sub ImportText()
    MDIForm1.MousePointer = 11
    Count1 = 1
    CountNr = 1
    Read4 = String(3000, " ")
    Get #GP2FileNum, oData.Text(GP2V) + 1, Read4
    Do Until Count1 > 16
        Read = String(1, " ")
        Read2 = ""
        Do Until Read = Chr(0)
            Read = Mid(Read4, CountNr, 1)
            If Read <> Chr(0) Then Read2 = Read2 + Read
            CountNr = CountNr + 1
        Loop
        If Len(Trim(Str(Count1))) = 1 Then
            Read3 = "0" + Trim(Str(Count1))
        Else
            Read3 = Count1
        End If
        Read3 = Gp2Dir + "\Circuits\f1ct" + Read3 + ".dat"
        Read2 = oMisc.WriteINI("Track " + Trim(Str(Count1)), "Name", Read2, ProgramDir + "\WorkCopy.lda")
        Read3 = oMisc.WriteINI("Track " + Trim(Str(Count1)), "TPath", Read3, ProgramDir + "\WorkCopy.lda")
        Count1 = Count1 + 1
    Loop
    CountNr = CountNr + 16
    Count1 = 1
    Do Until Count1 > 16
        Read = String(1, " ")
        Read2 = ""
        Do Until Read = Chr(0)
            Read = Mid(Read4, CountNr, 1)
            If Read <> Chr(0) Then Read2 = Read2 + Read
            CountNr = CountNr + 1
        Loop
        Read2 = oMisc.WriteINI("Track " + Trim(Str(Count1)), "Country", Read2, ProgramDir + "\WorkCopy.lda")
        Count1 = Count1 + 1
    Loop
    X = 1
    Do Until X > 4
        CountNr = CountNr + 16
        Count1 = 1
        Do Until Count1 > 16
            Read = String(1, " ")
            Read2 = ""
            Do Until Read = Chr(0)
                Read = Mid(Read4, CountNr, 1)
                If Read <> Chr(0) Then Read2 = Read2 + Read
                CountNr = CountNr + 1
            Loop
            Count1 = Count1 + 1
        Loop
        X = X + 1
    Loop

    CountNr = CountNr + 16
    Count1 = 1
    Do Until Count1 > 16
        Read = String(1, " ")
        Read2 = ""
        Do Until Read = Chr(0)
            Read = Mid(Read4, CountNr, 1)
            If Read <> Chr(0) Then Read2 = Read2 + Read
            CountNr = CountNr + 1
        Loop
        Read2 = oMisc.WriteINI("Track " + Trim(Str(Count1)), "Adjective", Read2, ProgramDir + "\WorkCopy.lda")
        Count1 = Count1 + 1
    Loop
    MakeNewTree
    Read4 = ""
    MDIForm1.MousePointer = 0
End Sub

Public Sub ImportPoints()
    X = 0
    Read = String(1, " ")
    Read3 = ""
    Do Until X > 25
        Get #GP2FileNum, oData.Point(GP2V) + X, Read
        Read2 = Asc(Read)
        If Read2 = 101 Then Read2 = "00"
        If Len(Read2) = 1 Then Read2 = "0" + Read2
        Read3 = Read3 + Read2
        X = X + 1
    Loop
    Read3 = oMisc.WriteINI("Misc", "Point", Read3, ProgramDir + "\WorkCopy.lda")
    Exit Sub
    X = 0
    Read = ""
    Read2 = ""
    Do Until X > 25
        Read2 = frmPoint.txtPoint(X).Text
        If Len(Read2) < 2 Then Read2 = " " + Read2
        Read = Read + Read2
        X = X + 1
    Loop
End Sub

Public Sub ImportNullAsOne()
    Read = String(1, " ")
    Get #GP2FileNum, oData.OneAsNull(GP2V), Read
    Read = Asc(Read)
    If Read = "255" Then Read = "1"
    If Read = "254" Then Read = "0"
    Read2 = oMisc.WriteINI("Misc", "0as1", Read, ProgramDir + "\WorkCopy.lda")
    MDIForm1.chk0as1.Value = Read
End Sub

Public Sub ImportQuick()
    Read = String(1, " ")
    Get #F1SaveFileNum, 648, Read
    Read = Asc(Read)
    Read2 = oMisc.WriteINI("Misc", "Quick", Read, ProgramDir + "\WorkCopy.lda")
    MDIForm1.HScroll2.Value = Read
End Sub

Public Sub ImportSaveLap()
    Read = String(2, " ")
    Get #GP2FileNum, oData.SaveLapTime, Read
    Read2 = String(2, " ")
    Get #GP2FileNum, oData.SaveLapTime2, Read2
    Read = Read + Read2
    Read3 = Chr(144)
    Read3 = String(3, Read3)
    Read2 = Chr(100)
    If Mid(Read, 2, 3) = Read3 And Mid(Read, 1, 1) = Read2 Then
        MDIForm1.chkSave.Value = 1
    Else
        MDIForm1.chkSave.Value = 0
    End If
    Read = oMisc.WriteINI("Misc", "SaveLap", MDIForm1.chkSave.Value, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportGameSettings()
    Read = String(1, " ")
    Count1 = 0
    Read3 = ""
    Do Until Count1 > 4
        X = 0
        Read = String(1, " ")
        Get #GP2FileNum, oData.Help + Count1, Read
        X = Asc(Read) ' + 1
        Read = ""
        Read2 = ""
        Dec2Bin (X)
        If Len(Read2) < 7 Then
            Temp = 7 - Len(Read2)
            Read = String(Temp, "0")
            Read = Read2 & Read
        ElseIf Len(Read2) = 7 Then
            Read = Read2
        End If
        Count1 = Count1 + 1
        If Len(Read) <> 7 Then
            Read = "0000000"
        End If
        Read3 = Read3 + Read
        Read = ""
        Read2 = ""
    Loop
    Read2 = oMisc.WriteINI("Misc", "Aids", Read3, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportLevel()
    Read = String(4, " ")
    Get #GP2FileNum, oData.Level(GP2V), Read
    MDIForm1.Slider1.Value = Read
    Read2 = oMisc.WriteINI("Misc", "Year", Read, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportPRPower()
    Read = String(1, " ")
    Get #GP2FileNum, oData.PRPower, Read
    Read2 = Asc(Read)
    Get #GP2FileNum, oData.PRPower + 1, Read
    Read3 = Asc(Read)
    Read = (Read3 * 256) + Read2
    MDIForm1.hscPRPower.Value = Read
    Read = oMisc.WriteINI("Player", "RPower", Read, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportPQPower()
    Read = String(1, " ")
    Get #GP2FileNum, oData.PQPower, Read
    Read2 = Asc(Read)
    Get #GP2FileNum, oData.PQPower + 1, Read
    Read3 = Asc(Read)
    Read = (Read3 * 256) + Read2
    MDIForm1.hscPQPower.Value = Read
    Read = oMisc.WriteINI("Player", "QPower", Read, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportPGrip()
    Read = String(1, " ")
    Get #GP2FileNum, oData.PGrip, Read
    Read2 = String(1, " ")
    Get #GP2FileNum, oData.PGrip + 1, Read2
    X = Asc(Read)
    Read2 = Asc(Read2)
    X = X + (256 * Read2)
    Read = oMisc.WriteINI("Player", "Grip", Trim(Str(X)), ProgramDir + "\WorkCopy.lda")
    MDIForm1.hscPGrip.Value = X
End Sub

Public Sub ImportSpeed()
    Read = String(1, " ")
    Get #GP2FileNum, oData.NoPitSpeed, Read
    If Asc(Read) = 235 Then
        Read = oMisc.WriteINI("Player", "Speed", "1", ProgramDir + "\WorkCopy.lda")
        MDIForm1.chkNoLimit.Value = 1
    End If
    If Asc(Read) = 116 Then
        Read = oMisc.WriteINI("Player", "NoSpeed", "0", ProgramDir + "\WorkCopy.lda")
        MDIForm1.chkNoLimit.Value = 0
    End If
    Read = String(1, " ")
    Get #GP2FileNum, oData.PitSpeed, Read
    Read2 = String(1, " ")
    Get #GP2FileNum, oData.PitSpeed + 1, Read2
    CountNr = 204
    Count1 = 0
    Count2 = 2
    Do Until (CountNr = Asc(Read)) And (Count2 = Asc(Read2))
        CountNr = CountNr + 68
        If CountNr > 255 Then
            CountNr = CountNr - 256
            Count2 = Count2 + 1
        End If
        Count2 = Count2 + 1
        Count1 = Count1 + 1
    Loop
    Read = Count1 + 1
    MDIForm1.hscPitSpeed = Read
    Read = oMisc.WriteINI("Player", "Speed", Read, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportCWeight()
    Read = String(1, " ")
    Get #GP2FileNum, oData.CWeight, Read
    Read2 = String(1, " ")
    Get #GP2FileNum, oData.CWeight + 1, Read2
    X = Asc(Read)
    Read2 = Asc(Read2)
    X = X + (256 * Read2)
    Read = oMisc.WriteINI("Misc", "CWeight", Trim(Str(X)), ProgramDir + "\WorkCopy.lda")
    MDIForm1.HScroll1.Value = X
End Sub

Public Sub ImportPWeight()
    Read = String(1, " ")
    Get #GP2FileNum, oData.PWeight, Read
    Read2 = String(1, " ")
    Get #GP2FileNum, oData.PWeight + 1, Read2
    X = Asc(Read)
    Read2 = Asc(Read2)
    X = X + (256 * Read2)
    Read = oMisc.WriteINI("Player", "Weight", Trim(Str(X)), ProgramDir + "\WorkCopy.lda")
    MDIForm1.hscWeight.Value = X
End Sub

Public Sub ImportUseTeam()
    Read = String(1, " ")
    Get #GP2FileNum, oData.UseTeam, Read
    If Asc(Read) = 255 Then
        Read = oMisc.WriteINI("Player", "UseTeam", "0", ProgramDir + "\WorkCopy.lda")
        MDIForm1.chkSelectedTeam.Value = 0
        Exit Sub
    End If
    If Asc(Read) = 0 Then
        Read = oMisc.WriteINI("Player", "UseTeam", "1", ProgramDir + "\WorkCopy.lda")
        MDIForm1.chkSelectedTeam.Value = 1
        Exit Sub
    End If
End Sub

Public Sub ImportQName()
    Read2 = Chr(0)
    Temp = 650 + (CountExport * 88)
    Read3 = ""
    Read = ""
    Do Until Read = Read2
        Read = String(1, " ")
        Get #FileNum, Temp, Read
        If Read <> Read2 Then Read3 = Read3 + Read
        Temp = Temp + 1
    Loop
    Read3 = oMisc.WriteINI("Track " + Trim(Str(CountExport + 1)), "QDriver", Read3, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportRName()
    Read2 = Chr(0)
    Temp = 694 + (CountExport * 88)
    Read3 = ""
    Read = ""
    Do Until Read = Read2
        Read = String(1, " ")
        Get #FileNum, Temp, Read
        If Read <> Read2 Then Read3 = Read3 + Read
        Temp = Temp + 1
    Loop
    Read3 = oMisc.WriteINI("Track " + Trim(Str(CountExport + 1)), "RDriver", Read3, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportQTeam()
    Read2 = Chr(0)
    Temp = 674 + (CountExport * 88)
    Read3 = ""
    Read = ""
    Do Until Read = Read2
        Read = String(1, " ")
        Get #FileNum, Temp, Read
        If Read <> Read2 Then Read3 = Read3 + Read
        Temp = Temp + 1
    Loop
    Read3 = oMisc.WriteINI("Track " + Trim(Str(CountExport + 1)), "QTeam", Read3, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportRTeam()
    Read2 = Chr(0)
    Temp = 718 + (CountExport * 88)
    Read3 = ""
    Read = ""
    Do Until Read = Read2
        Read = String(1, " ")
        Get #FileNum, Temp, Read
        If Read <> Read2 Then Read3 = Read3 + Read
        Temp = Temp + 1
    Loop
    Read3 = oMisc.WriteINI("Track " + Trim(Str(CountExport + 1)), "RTeam", Read3, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportRaceTime()
    Temp = 732 + (CountExport * 88)
    TempDouble = 0
    Read = String(1, " ")
    Get #FileNum, Temp + 2, Read
    X = Asc(Read)
    TempDouble = X * 65.536
    Read = String(1, " ")
    Get #FileNum, Temp + 1, Read
    X = Asc(Read)
    TempDouble = TempDouble + (X * 0.256)
    Read = String(1, " ")
    Get #FileNum, Temp, Read
    X = Asc(Read)
    TempDouble = TempDouble + (X / 1000)
    Read = TempDouble
    Read2 = Mid(Read, 1, 1)
    If Mid(Read, 2, 1) <> "," Then
        Read2 = Read2 + Mid(Read, 2, 1)
        If Mid(Read, 3, 1) <> "," Then
            Read2 = Read2 + Mid(Read, 3, 1)
        End If
    End If
    X = Read2
    Count1 = 0
    Do Until X < 60
        X = X - 60
        TempDouble = TempDouble - 60
        Count1 = Count1 + 1
    Loop
    Read = X
    If Len(Read) = 1 Then Read = "0" + Read
    
    Dim TempInt As Integer
    TempInt = Len(Trim(Str(X)))
    X = Len(Trim(Str(TempDouble)))
    If X > 6 Then
        X = 6
    End If
    Read2 = TempDouble
    If TempInt = X - 4 Then Read2 = Mid(Read2, TempInt + 2, 3)
    If TempInt = X - 3 Then Read2 = Mid(Read2, TempInt + 2, 2) + "0"
    If TempInt = X - 2 Then Read2 = Mid(Read2, TempInt + 2, 1) + "00"
    If X = TempInt Then Read2 = "000"
    
    Read3 = Trim(Str(Count1)) + "." + Read + "." + Read2
    Read3 = oMisc.WriteINI("Track " + Trim(Str(CountExport + 1)), "RTime", Read3, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportQTime()
    Temp = 688 + (CountExport * 88)
    TempDouble = 0
    Read = String(1, " ")
    Get #FileNum, Temp + 2, Read
    X = Asc(Read)
    TempDouble = X * 65.536
    Read = String(1, " ")
    Get #FileNum, Temp + 1, Read
    X = Asc(Read)
    TempDouble = TempDouble + (X * 0.256)
    Read = String(1, " ")
    Get #FileNum, Temp, Read
    X = Asc(Read)
    TempDouble = TempDouble + (X / 1000)
    Read = TempDouble
    Read2 = Mid(Read, 1, 1)
    If Mid(Read, 2, 1) <> "," Then
        Read2 = Read2 + Mid(Read, 2, 1)
        If Mid(Read, 3, 1) <> "," Then
            Read2 = Read2 + Mid(Read, 3, 1)
        End If
    End If
    X = Read2
    Count1 = 0
    Do Until X < 60
        X = X - 60
        TempDouble = TempDouble - 60
        Count1 = Count1 + 1
    Loop
    Read = X
    If Len(Read) = 1 Then Read = "0" + Read
    
    Dim TempInt As Integer
    TempInt = Len(Trim(Str(X)))
    X = Len(Trim(Str(TempDouble)))
    If X > 6 Then
        X = 6
    End If
    Read2 = TempDouble
    If TempInt = X - 4 Then Read2 = Mid(Read2, TempInt + 2, 3)
    If TempInt = X - 3 Then Read2 = Mid(Read2, TempInt + 2, 2) + "0"
    If TempInt = X - 2 Then Read2 = Mid(Read2, TempInt + 2, 1) + "00"
    If X = TempInt Then Read2 = "000"
    
    Read3 = Trim(Str(Count1)) + "." + Read + "." + Read2

    Read3 = oMisc.WriteINI("Track " + Trim(Str(CountExport + 1)), "QTime", Read3, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportQDate()
    Temp = 692 + (CountExport * 88)
    Read = String(2, " ")
    Get #FileNum, Temp, Read
    
    Read2 = Mid(Read, 1, 1)
    X = Asc(Read2)
    Read2 = Mid(Read, 2, 1)
    Count1 = Asc(Read2)
    Count1 = (Count1 * 256) + X
    TheDate = "1978-01-01"
    Read2 = TheDate + Count1
    Read2 = oMisc.WriteINI("Track " + Trim(Str(CountExport + 1)), "QDate", Read2, ProgramDir + "\WorkCopy.lda")
End Sub

Public Sub ImportRDate()
    Temp = 736 + (CountExport * 88)
    Read = String(2, " ")
    Get #FileNum, Temp, Read
    Read2 = Mid(Read, 1, 1)
    X = Asc(Read2)
    Read2 = Mid(Read, 2, 1)
    Count1 = Asc(Read2)
    Count1 = (Count1 * 256) + X
    TheDate = "1978-01-01"
    Read2 = TheDate + Count1
    Read2 = oMisc.WriteINI("Track " + Trim(Str(CountExport + 1)), "RDate", Read2, ProgramDir + "\WorkCopy.lda")
End Sub
