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
    Get #Exp.GP2FileNum, Count1 + (Exp.TrackNr * 7), tImp.iInt
    If tImp.iInt < 0 Then
        tImp.lLong = tImp.iInt + 65535
    Else
        tImp.lLong = tImp.iInt
    End If
    TempDouble = Round(tImp.lLong / 3.28212677519917, 0)
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "Length", Trim(Str(TempDouble)), TempFile
End Sub

Public Sub ImportLaps()
    Count1 = oData.Laps(GP2V)
    Get #Exp.GP2FileNum, Count1 + Exp.TrackNr, tImp.bByte
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "Laps", Trim(Str(tImp.bByte)), TempFile
End Sub

Public Sub ImportWare()
    Count1 = oData.Ware
    Get #Exp.GP2FileNum, Count1 + (Exp.TrackNr * 2), tImp.iInt
    tImp.lLong = tImp.iInt
    If tImp.lLong < 0 Then tImp.lLong = tImp.lLong + 65535
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "Ware", Trim(Str(tImp.lLong)), TempFile
End Sub

Public Sub ImportText()
Read = String(3000, " ")
    Get #Exp.GP2FileNum, oData.Text(GP2V) + 1, Read
    
    For Count1 = 0 To 15
        Count2 = InStr(1, Read, Chr(0))
        Read2 = Mid(Read, 1, Count2 - 1)
        TrackName(Count1) = Read2
        Read = Mid(Read, Count2 + 1)
    Next

    Read = Mid(Read, 17)
    For Count1 = 0 To 15
        Count2 = InStr(1, Read, Chr(0))
        Read2 = Mid(Read, 1, Count2 - 1)
        Country(Count1) = Read2
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
        Adj(Count1) = Read2
        Read = Mid(Read, Count2 + 1)
    Next
    Read = ""
End Sub

Public Sub ImportPoints()
    X = 0
    Read = String(1, " ")
    Read3 = ""
    Do Until X > 25
        Get #Exp.GP2FileNum, oData.Point(GP2V) + X, Read
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
    Get #Exp.GP2FileNum, oData.OneAsNull(GP2V), tImp.bByte
    If tImp.bByte = 255 Then
        oMisc.WriteINI "Misc", "0as1", "1", TempFile
    Else
        oMisc.WriteINI "Misc", "0as1", "0", TempFile
    End If
End Sub

Public Sub ImportQuick()
    Get #Exp.F1FileNum, 648, tImp.bByte
    oMisc.WriteINI "Misc", "Quick", Trim(Str(tImp.bByte)), TempFile
End Sub

Public Sub ImportSaveLap()
    Read = String(2, " ")
    Get #Exp.GP2FileNum, oData.SaveLapTime, Read
    Read2 = String(2, " ")
    Get #Exp.GP2FileNum, oData.SaveLapTime2, Read2
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
        Get #Exp.GP2FileNum, oData.Help + Count1, tImp.bByte
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
    Get #Exp.GP2FileNum, oData.Level(GP2V), tImp.Year
    oMisc.WriteINI "Misc", "Year", tImp.Year, TempFile
End Sub

Public Sub ImportPRPower()
    Get #Exp.GP2FileNum, oData.PRPower, tImp.iInt
    oMisc.WriteINI "Player", "RPower", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportPQPower()
    Get #Exp.GP2FileNum, oData.PQPower, tImp.iInt
    oMisc.WriteINI "Player", "QPower", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportPGrip()
    Get #Exp.GP2FileNum, oData.PGrip, tImp.iInt
    oMisc.WriteINI "Player", "Grip", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportSpeed()
    Get #Exp.GP2FileNum, oData.NoPitSpeed, tImp.bByte
    If tImp.bByte = 235 Then
        oMisc.WriteINI "Player", "Speed", "1", TempFile
    ElseIf tImp.bByte = 116 Then
        oMisc.WriteINI "Player", "NoSpeed", "0", TempFile
    End If
    Get #Exp.GP2FileNum, oData.PitSpeed, tImp.lLong
    tImp.lLong = (tImp.lLong - 392) / 324
    oMisc.WriteINI "Player", "Speed", Trim(Str(tImp.lLong)), TempFile
End Sub

Public Sub ImportCWeight()
    Get #Exp.GP2FileNum, oData.CWeight, tImp.iInt
    oMisc.WriteINI "Misc", "CWeight", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportPWeight()
    Get #Exp.GP2FileNum, oData.PWeight, tImp.iInt
    oMisc.WriteINI "Player", "Weight", Trim(Str(tImp.iInt)), TempFile
End Sub

Public Sub ImportUseTeam()
    Get #Exp.GP2FileNum, oData.UseTeam, tImp.bByte
    If tImp.bByte = 255 Then
        oMisc.WriteINI "Player", "UseTeam", "0", TempFile
    ElseIf tImp.bByte = 0 Then
        oMisc.WriteINI "Player", "UseTeam", "1", TempFile
    End If
End Sub

Public Sub ImportQName(ByVal Rec As RecEnum)
    If Rec = F1gstate Then
        Count1 = 650 + (Exp.TrackNr * 88)
    ElseIf Rec = RecFile Then
        Count1 = 33 + (Exp.TrackNr * 88)
    End If
    Read = String(22, " ")
    Get #Exp.F1FileNum, Count1, Read
    Read = Left(Read, InStr(Read, vbNullChar) - 1)
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "QDriver", Read, TempFile
End Sub

Public Sub ImportRName(ByVal Rec As RecEnum)
    If Rec = F1gstate Then
        Count1 = 694 + (Exp.TrackNr * 88)
    ElseIf Rec = RecFile Then
        Count1 = 77 + (Exp.TrackNr * 88)
    End If
    Read = String(22, " ")
    Get #Exp.F1FileNum, Count1, Read
    Read = Left(Read, InStr(Read, vbNullChar) - 1)
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "RDriver", Read, TempFile
End Sub

Public Sub ImportQTeam(ByVal Rec As RecEnum)
    If Rec = F1gstate Then
        Count1 = 674 + (Exp.TrackNr * 88)
    ElseIf Rec = RecFile Then
        Count1 = 57 + (Exp.TrackNr * 88)
    End If
    Read = String(12, " ")
    Get #Exp.F1FileNum, Count1, Read
    Read = Left(Read, InStr(Read, vbNullChar) - 1)
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "QTeam", Read, TempFile
End Sub

Public Sub ImportRTeam(ByVal Rec As RecEnum)
    If Rec = F1gstate Then
        Count1 = 718 + (Exp.TrackNr * 88)
    ElseIf Rec = RecFile Then
        Count1 = 101 + (Exp.TrackNr * 88)
    End If
    Read = String(12, " ")
    Get #Exp.F1FileNum, Count1, Read
    Read = Left(Read, InStr(Read, vbNullChar) - 1)
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "RTeam", Read, TempFile
End Sub

Public Function ImportTime(ByVal QR As ImpExpTime, ByVal Rec As RecEnum)
Dim M As Integer
Dim S As Integer
    If QR = Qual Then
        If Rec = F1gstate Then
            Get #Exp.F1FileNum, 688 + (Exp.TrackNr * 88), tImp.lLong
        ElseIf Rec = RecFile Then
            Get #Exp.F1FileNum, 71 + (Exp.TrackNr * 88), tImp.lLong
        End If
    Else
        If Rec = F1gstate Then
            Get #Exp.F1FileNum, 732 + (Exp.TrackNr * 88), tImp.lLong
        ElseIf Rec = RecFile Then
            Get #Exp.F1FileNum, 115 + (Exp.TrackNr * 88), tImp.lLong
        End If
    End If
    M = 0
    Do Until tImp.lLong < 60000
        M = M + 1
        tImp.lLong = tImp.lLong - 60000
    Loop
    Do Until tImp.lLong < 1000
        S = S + 1
        tImp.lLong = tImp.lLong - 1000
    Loop
    If S < 10 Then
        Read = M & ":0" & S
    Else
        Read = M & ":" & S
    End If
    If tImp.lLong < 10 Then
        Read = Read & ".00" & tImp.lLong
    ElseIf tImp.lLong < 100 Then
        Read = Read & ".0" & tImp.lLong
    Else
        Read = Read & "." & tImp.lLong
    End If
    If QR = Qual Then
        oMisc.WriteINI "Track " & Exp.TrackNr + 1, "QTime", Read, TempFile
    Else
        oMisc.WriteINI "Track " & Exp.TrackNr + 1, "RTime", Read, TempFile
    End If
End Function

Public Sub ImportQDate(ByVal Rec As RecEnum)
    If Rec = F1gstate Then
        Get #Exp.F1FileNum, 692 + (Exp.TrackNr * 88), tImp.iInt
    ElseIf Rec = RecFile Then
        Get #Exp.F1FileNum, 75 + (Exp.TrackNr * 88), tImp.iInt
    End If
    TheDate = "1978-01-01"
    Read = TheDate + tImp.iInt
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "QDate", Read, TempFile
End Sub

Public Sub ImportRDate(ByVal Rec As RecEnum)
    If Rec = F1gstate Then
        Get #Exp.F1FileNum, 736 + (Exp.TrackNr * 88), tImp.iInt
    ElseIf Rec = RecFile Then
        Get #Exp.F1FileNum, 119 + (Exp.TrackNr * 88), tImp.iInt
    End If
    TheDate = "1978-01-01"
    Read = TheDate + tImp.iInt
    oMisc.WriteINI "Track " & Exp.TrackNr + 1, "RDate", Read, TempFile
End Sub
