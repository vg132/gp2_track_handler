Attribute VB_Name = "modImport"
Private Function Dec2Bin(MyNum As Byte) As String
Dim LoopCounter As Integer
Dim Bin As String
    Do Until 2 ^ LoopCounter > MyNum
        If (MyNum And 2 ^ LoopCounter) = 2 ^ LoopCounter Then
            Bin = Bin & "1"
        Else
            Bin = Bin & "0"
        End If
        LoopCounter = LoopCounter + 1
    Loop
    Dec2Bin = Bin
End Function

Public Sub ImportLength()
    Count1 = oData.Length(GP2V)
    For X = 0 To 15
        Get #GP2FileNum, Count1 + (X * 7), tImp.iInt
        If tImp.iInt < 0 Then
            tImp.lLong = tImp.iInt + 65535
        Else
            tImp.lLong = tImp.iInt
        End If
        TempDouble = tImp.lLong / 3.28212677519917
        oMisc.WriteINI "Track " & X + 1, "Length", Trim(Str(Round(TempDouble, 0))), TempFile
    Next
End Sub

Public Sub ImportLaps()
    Count1 = oData.Laps(GP2V)
    For X = 0 To 15
        Get #GP2FileNum, Count1 + X, tImp.bByte
        oMisc.WriteINI "Track " & X + 1, "Laps", Trim(Str(tImp.bByte)), TempFile
    Next
End Sub

Public Sub ImportWare()
    X = 0
    Y = 1
    For Y = 1 To 16
        Get #GP2FileNum, oData.Ware + X, tImp.iInt
        tImp.lLong = tImp.iInt
        If tImp.lLong < 0 Then tImp.lLong = tImp.lLong + 65535
        oMisc.WriteINI "Track " & Y, "Ware", Trim(Str(tImp.lLong)), TempFile
        X = X + 2
    Next
End Sub

Public Sub ImportText()
Read = String(3000, " ")
    Get #GP2FileNum, oData.Text(GP2V) + 1, Read
    
    For Count1 = 0 To 15
        Count2 = InStr(1, Read, Chr(0))
        Read2 = Mid(Read, 1, Count2 - 1)
        oMisc.WriteINI "Track " & Count1 + 1, "Name", Read2, TempFile
        If Len(Trim(Str(Count1 + 1))) = 1 Then
            Read2 = "0" & Count1 + 1
        Else
            Read2 = Count1 + 1
        End If
        Read2 = GP2Dir & "\Circuits\f1ct" & Read2 & ".dat"
        oMisc.WriteINI "Track " & Count1 + 1, "TPath", Read2, TempFile
        Read = Mid(Read, Count2 + 1)
    Next
    
    Read = Mid(Read, 17)
    For Count1 = 0 To 15
        Count2 = InStr(1, Read, Chr(0))
        Read2 = Mid(Read, 1, Count2 - 1)
        oMisc.WriteINI "Track " & Count1 + 1, "Country", Read2, TempFile
        Read = Mid(Read, Count2 + 1)
    Next

    Read = Mid(Read, 17)
    For Count1 = 0 To 3
        Read2 = String(17, Chr(0))
        Count2 = InStr(1, Read, Read2)
        Read = Mid(Read, Count2 + 1)
    Next
    
    Read = Mid(Read, 17)
    For Count1 = 0 To 15
        Count2 = InStr(1, Read, Chr(0))
        Read2 = Mid(Read, 1, Count2 - 1)
        oMisc.WriteINI "Track " & Count1 + 1, "Adjective", Read2, TempFile
        Read = Mid(Read, Count2 + 1)
    Next
    Read = ""
    LoadFile
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
    oMisc.WriteINI "Misc", "Point", Read3, TempFile
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
    Get #GP2FileNum, oData.OneAsNull(GP2V), tImp.bByte
    If tImp.bByte = 255 Then
        oMisc.WriteINI "Misc", "0as1", "1", TempFile
    Else
        oMisc.WriteINI "Misc", "0as1", "0", TempFile
    End If
End Sub

Public Sub ImportQuick()
    Get #F1SaveFileNum, 648, tImp.bByte
    oMisc.WriteINI "Misc", "Quick", Trim(Str(tImp.bByte)), TempFile
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
        frmMain.chkSave.Value = 1
    Else
        frmMain.chkSave.Value = 0
    End If
    oMisc.WriteINI "Misc", "SaveLap", frmMain.chkSave.Value, TempFile
End Sub

Public Sub ImportGameSettings()
    Read3 = ""
    For Count1 = 0 To 4
        Get #GP2FileNum, oData.Help + Count1, tImp.bByte
        Read2 = Dec2Bin(tImp.bByte)
        If Len(Read2) < 7 Then
            Temp = 7 - Len(Read2)
            Read = String(Temp, "0")
            Read = Read2 & Read
        ElseIf Len(Read2) = 7 Then
            Read = Read2
        End If
        If Len(Read) <> 7 Then
            Read = "0000000"
        End If
        Read3 = Read3 + Read
        Read = ""
        Read2 = ""
    Next
    oMisc.WriteINI "Misc", "Aids", Read3, TempFile
End Sub

Public Sub ImportLevel()
    Get #GP2FileNum, oData.Level(GP2V), tImp.Year
    oMisc.WriteINI "Misc", "Year", tImp.Year, TempFile
End Sub

Public Sub ImportPRPower()
    Get #GP2FileNum, oData.PRPower, tImp.iInt
    oMisc.WriteINI "Player", "RPower", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportPQPower()
    Get #GP2FileNum, oData.PQPower, tImp.iInt
    oMisc.WriteINI "Player", "QPower", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportPGrip()
    Get #GP2FileNum, oData.PGrip, tImp.iInt
    oMisc.WriteINI "Player", "Grip", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportSpeed()
    Get #GP2FileNum, oData.NoPitSpeed, tImp.bByte
    If tImp.bByte = 235 Then
        oMisc.WriteINI "Player", "Speed", "1", TempFile
    ElseIf tImp.bByte = 116 Then
        oMisc.WriteINI "Player", "NoSpeed", "0", TempFile
    End If
    Get #GP2FileNum, oData.PitSpeed, tImp.lLong
    tImp.lLong = (tImp.lLong - 392) / 324
    oMisc.WriteINI "Player", "Speed", Trim(Str(tImp.lLong)), TempFile
End Sub

Public Sub ImportCWeight()
    Get #GP2FileNum, oData.CWeight, tImp.iInt
    oMisc.WriteINI "Misc", "CWeight", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportPWeight()
    Get #GP2FileNum, oData.PWeight, tImp.iInt
    oMisc.WriteINI "Player", "Weight", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportUseTeam()
    Get #GP2FileNum, oData.UseTeam, tImp.bByte
    If tImp.bByte = 255 Then
        oMisc.WriteINI "Player", "UseTeam", "0", TempFile
    ElseIf tImp.bByte = 0 Then
        oMisc.WriteINI "Player", "UseTeam", "1", TempFile
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
    oMisc.WriteINI "Track " + Trim(Str(CountExport + 1)), "QDriver", Read3, TempFile
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
    oMisc.WriteINI "Track " + Trim(Str(CountExport + 1)), "RDriver", Read3, TempFile
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
    oMisc.WriteINI "Track " + Trim(Str(CountExport + 1)), "QTeam", Read3, TempFile
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
    oMisc.WriteINI "Track " + Trim(Str(CountExport + 1)), "RTeam", Read3, TempFile
End Sub

Public Function ImportTimeFromGP2(ByVal QR As ImpExpTime)
Dim M As Integer
    If QR = Qual Then
        Get #FileNum, 688 + (CountExport * 88), tImp.lLong
    Else
        Get #FileNum, 732 + (CountExport * 88), tImp.lLong
    End If
    TempDouble = tImp.lLong / 60000
    M = Fix(TempDouble)
    tImp.lLong = tImp.lLong - (M * 60000)
    TempDouble = tImp.lLong / 1000
    If TempDouble < 10 Then
        Read = "0" & Trim(Str(TempDouble))
    Else
        Read = TempDouble
    End If
    If Len(Read) = 2 Then
        Read = Read & ".000"
    ElseIf Len(Read) = 4 Then
        Read = Read & "00"
    ElseIf Len(Read) = 5 Then
        Read = Read & "0"
    End If
    
    Read = Trim(Str(M)) & ":" & Read
    X = InStr(1, Read, ",")
    If X > 0 Then
    Read = Mid(Read, 1, X - 1) & "." & Mid(Read, X + 1)
    End If
    If QR = Qual Then
        oMisc.WriteINI "Track " & CountExport + 1, "QTime", Read, TempFile
    Else
        oMisc.WriteINI "Track " & CountExport + 1, "RTime", Read, TempFile
    End If
End Function

Public Sub ImportQDate()
    Get #FileNum, 692 + (CountExport * 88), tImp.iInt
    TheDate = "1978-01-01"
    Read = TheDate + tImp.iInt
    oMisc.WriteINI "Track " + Trim(Str(CountExport + 1)), "QDate", Read, TempFile
End Sub

Public Sub ImportRDate()
    Get #FileNum, 736 + (CountExport * 88), tImp.iInt
    TheDate = "1978-01-01"
    Read = TheDate + tImp.iInt
    oMisc.WriteINI "Track " + Trim(Str(CountExport + 1)), "RDate", Read, TempFile
End Sub
