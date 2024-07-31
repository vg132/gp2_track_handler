Attribute VB_Name = "modExport"
Public Sub SetAttribut()
    On Error Resume Next
    For X = 1 To 16
        Read = Str(X)
        Read = Trim(Read)
        If Len(Read) < 2 Then Read = "0" + Read
        Read = Gp2Dir + "\Circuits\f1ct" + Read + ".dat"
        SetAttr Read, vbNormal
    Next
    Read = Gp2Dir + "\gp2.exe"
    SetAttr Read, vbNormal
End Sub

Public Sub ExportLaps()
Dim Lap As Byte
    Read = ReadINI("Track " & Exp.TrackNr + 1, "Laps", TempFile)
    If Read <> "" Then
        Lap = Read
        If Lap > 126 Then Lap = 126
        If Lap < 3 Then Lap = 3
        Put #Exp.Gp2FileNum, oData.Laps(Gp2V) + Exp.TrackNr, Lap
    End If
End Sub

Public Sub ExportName()
    Read = ReadINI("Track " & Exp.TrackNr + 1, "Name", TempFile)
    If Read <> "" Then
        Gp2NameFile = Gp2NameFile + Trim(Read) + Chr(0)
    Else
        Gp2NameFile = Gp2NameFile + TrackName(Exp.TrackNr) + Chr(0)
    End If
End Sub

Public Sub ExportCountry()
    Read = ReadINI("Track " & Exp.TrackNr + 1, "Country", TempFile)
    If Read <> "" Then
        Gp2NameFile = Gp2NameFile + Trim(Read) + Chr(0)
    Else
        Gp2NameFile = Gp2NameFile + Country(Exp.TrackNr) + Chr(0)
    End If
End Sub

Public Sub ExportAdjectiv()
    Read = ReadINI("Track " & Exp.TrackNr + 1, "Adjective", TempFile)
    If Read <> "" Then
        Gp2NameFile = Gp2NameFile + Trim(Read) + Chr(32) + Chr(0)
    Else
        Gp2NameFile = Gp2NameFile + Adj(Exp.TrackNr) + Chr(0)
    End If
End Sub

Public Sub ExportTracks()
    Read = ReadINI("Track " & Exp.TrackNr + 1, "TPath", TempFile)
    If Read <> "" Then
        Read2 = oFile.FileExists(Read)
        If Read2 = True Then
            SourceFile = Read
            If Exp.TrackNr + 1 < 10 Then
                Read2 = "0" & Exp.TrackNr + 1
            Else
                Read2 = Exp.TrackNr + 1
            End If
            Read2 = Gp2Dir & "\Circuits\F1ct" & Read2 & ".dat"
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
Dim l As Long
Dim d As Double
Dim i As Integer
    Read = ReadINI("Track " & Exp.TrackNr + 1, "Length", TempFile)
    If Read <> "" Then
        l = Read
        d = l * 3.28212677519917
        l = Round(d, 0)
        If l > 32767 Then l = l - 65535
        i = l
        Put #Exp.Gp2FileNum, oData.Length(Gp2V) + (Exp.TrackNr * 7), i
    End If
End Sub

Public Sub ExportWare()
Dim l As Long
Dim i As Long
    Read = ReadINI("Track " & Exp.TrackNr + 1, "Ware", TempFile)
    If Read <> "" Then
        l = Read
        If l > 32767 Then
            l = l - 65535
        End If
        i = l
        Put #Exp.Gp2FileNum, oData.Ware + (Exp.TrackNr * 2), i
    End If
End Sub

Public Sub ExportPoints()
    Read = ReadINI("Misc", "Point", TempFile)
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
    Put #Exp.Gp2FileNum, oData.Point(Gp2V), Read3
End Sub

Public Sub ExportNullAsOne()
Dim b As Byte
    b = ReadINI("Misc", "0as1", TempFile)
    If b = "1" Then
      b = "255"
    ElseIf b = "0" Then
      b = "254"
    Else
      b = 255
    End If
    Put #Exp.Gp2FileNum, oData.OneAsNull(Gp2V), b
End Sub

Public Sub ExportQuickRace()
Dim b As Byte
    b = ReadINI("Misc", "Quick", TempFile)
    Put #Exp.F1FileNum, 648, b
End Sub

Public Sub ExportSaveLap()
    Read = ReadINI("Misc", "SaveLap", TempFile)
    If Read = 1 Then
        Read = Chr(100) + Chr(144)
        Put #Exp.Gp2FileNum, oData.SaveLapTime, Read
        Read = Chr(144) + Chr(144)
        Put #Exp.Gp2FileNum, oData.SaveLapTime2, Read
    Else
        Read2 = ""
        Read = Chr(92)
        Read2 = Read
        Read = Chr(114)
        Read2 = Read2 + Read
        Put #Exp.Gp2FileNum, oData.SaveLapTime, Read2
        Read2 = ""
        Read = Chr(114)
        Read2 = Read
        Read = Chr(92)
        Read2 = Read2 + Read
        Put #Exp.Gp2FileNum, oData.SaveLapTime2, Read2
    End If
End Sub

Public Sub ExportLevel()
Dim sYear As String * 4
    sYear = ReadINI("Misc", "Year", TempFile)
    If sYear <> "" Then Put #Exp.Gp2FileNum, oData.Level(Gp2V), sYear
End Sub

Public Sub ExportCarHelp()
Dim CountNr As Integer
Dim iBigLoop As Integer
Dim iLowLoop
Dim bHelp As Byte

    Read = ReadINI("Misc", "Aids", TempFile)
    For iBigLoop = 0 To 4
        Read2 = Mid(Read, iBigLoop * 7 + 1, 7)
        bHelp = 0
        For iLowLoop = 0 To 6
            Read3 = Mid(Read2, iLowLoop + 1, 1)
            If Read3 = 1 Then
                bHelp = bHelp + 2 ^ iLowLoop
            End If
        Next
        Put #Exp.Gp2FileNum, oData.Help + iBigLoop, bHelp
    Next

End Sub

Public Sub ExportPQPower()
Dim i As Integer
    Read = ReadINI("Player", "QPower", TempFile)
    If Read <> "" Then
        i = Read
        Put #Exp.Gp2FileNum, oData.PQPower, i
    End If
End Sub

Public Sub ExportPRPower()
Dim i As Integer
    Read = ReadINI("Player", "RPower", TempFile)
    If Read <> "" Then
        i = Read
        Put #Exp.Gp2FileNum, oData.PRPower, i
    End If
End Sub

Public Sub ExportPGrip()
Dim i As Integer
    i = ReadINI("Player", "Grip", TempFile)
    Put #Exp.Gp2FileNum, oData.PGrip, i
End Sub

Public Sub ExportPWeight()
Dim i As Integer
    Read = ReadINI("Player", "Weight", TempFile)
    If Read <> "" Then
        i = Read
        Put #Exp.Gp2FileNum, oData.PWeight, i
    End If
End Sub

Public Sub ExportCWeight()
Dim i As Integer
    Read = ReadINI("Misc", "CWeight", TempFile)
    If Read <> "" Then
        i = Read
        Put #Exp.Gp2FileNum, oData.CWeight, i
    End If
End Sub

Public Sub ExportSpeed()
    Read = ReadINI("Player", "NoSpeed", TempFile)
    If Read = 1 Then
        tVar.bByte = 235
        Put #Exp.Gp2FileNum, oData.NoPitSpeed, tVar.bByte
    Else
        tVar.bByte = 116
        Put #Exp.Gp2FileNum, oData.NoPitSpeed, tVar.bByte
    End If
    Read = ReadINI("Player", "Speed", TempFile)
    tVar.lLong = Read
    tVar.lLong = (tVar.lLong * 324) + 392
    Put #Exp.Gp2FileNum, oData.PitSpeed, tVar.lLong
End Sub

Public Sub ExportUseTeam()
Dim b As Byte
    b = ReadINI("Player", "UseTeam", TempFile)
    If b = 0 Then
      b = 255
    ElseIf b = 1 Then
      b = 0
    Else
      b = 255
    End If
    Put #Exp.Gp2FileNum, oData.UseTeam, b
End Sub

Public Sub ExportPictures()
    FileNum = FreeFile
    Open ProgramDir & "\Bat\Export.bat" For Append As FileNum
    Read = ReadINI("Track " & Exp.TrackNr + 1, "BPic", TempFile)
    If Read <> "" Then
        Read = oFile.GetShortName(Read)
        Read2 = oFile.GetShortName(Gp2Dir)
        Read = Read2 & "\gp2hipic.exe -q #" & Exp.TrackNr + 1 & " " & Read2 & "\bitmaps\f1pcsvga.bin " & Read
        Print #FileNum, LCase(Read)
    End If
    Read = ReadINI("Track " & Exp.TrackNr + 1, "SPic", TempFile)
    If Read <> "" Then
        Read = oFile.GetShortName(Read)
        Read2 = oFile.GetShortName(Gp2Dir)
        Read = Read2 & "\gp2hipic.exe -q #" & Exp.TrackNr + 17 & " " & Read2 & "\bitmaps\f1pcsvga.bin " & Read
        Print #FileNum, LCase(Read)
    End If
    If Exp.TrackNr + 1 = 16 Then
        Read2 = oFile.GetShortName(Gp2Dir)
        Read = Read2 & "\gp2hipic.exe -d " & Read2 & "\bitmaps\f1pcsvga.bin"
        Print #FileNum, LCase(Read)
    End If
    Close FileNum
End Sub

Public Sub ExportRTeam(ByVal Rec As RecEnum)
    Read = ReadINI("Track " & Exp.TrackNr + 1, "RTeam", TempFile)
    If Read <> "" Then
        tVar.iInt = Len(Read)
        Read2 = Chr(0)
        Read2 = String(12 - tVar.iInt, Read2)
        Read = Read & Read2
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 718 + (Exp.TrackNr * 88), Read
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 101 + (Exp.TrackNr * 88), Read
        End If
    End If
End Sub

Public Sub ExportQTeam(ByVal Rec As RecEnum)
    Read = ReadINI("Track " & Exp.TrackNr + 1, "QTeam", TempFile)
    If Read <> "" Then
        tVar.iInt = Len(Read)
        Read2 = Chr(0)
        Read2 = String(12 - tVar.iInt, Read2)
        Read = Read & Read2
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 674 + (Exp.TrackNr * 88), Read
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 57 + (Exp.TrackNr * 88), Read
        End If
    End If
End Sub

Public Sub ExportRName(ByVal Rec As RecEnum)
    Read = ReadINI("Track " & Exp.TrackNr + 1, "RDriver", TempFile)
    If Read <> "" Then
        tVar.iInt = Len(Read)
        Read2 = Chr(0)
        Read2 = String(23 - tVar.iInt, Read2)
        Read = Read & Read2
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 694 + (Exp.TrackNr * 88), Read
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 77 + (Exp.TrackNr * 88), Read
        End If
    End If
End Sub

Public Sub ExportQName(ByVal Rec As RecEnum)
    Read = ReadINI("Track " & Exp.TrackNr + 1, "QDriver", TempFile)
    If Read <> "" Then
        tVar.iInt = Len(Read)
        Read2 = Chr(0)
        Read2 = String(23 - tVar.iInt, Read2)
        Read = Read & Read2
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 650 + (Exp.TrackNr * 88), Read
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 33 + (Exp.TrackNr * 88), Read
        End If
    End If
End Sub

Public Sub ExportTime(ByVal QR As QR, ByVal Rec As RecEnum)
    If QR = Qual Then
        Read = ReadINI("Track " & Exp.TrackNr + 1, "QTime", TempFile)
    Else
        Read = ReadINI("Track " & Exp.TrackNr + 1, "RTime", TempFile)
    End If
    If Read <> "" Then
        Count1 = Mid(Read, 1, InStr(1, Read, ":") - 1)
        Count2 = Mid(Read, InStr(1, Read, ":") + 1)
        Count1 = Count1 * 60000
        tVar.lLong = Count2 + Count1
        If QR = Qual Then
            If Rec = F1gstate Then
                Put #Exp.F1FileNum, 688 + (Exp.TrackNr * 88), tVar.lLong
            ElseIf Rec = RecFile Then
                Put #Exp.F1FileNum, 71 + (Exp.TrackNr * 88), tVar.lLong
            End If
        Else
            If Rec = F1gstate Then
            Put #Exp.F1FileNum, 732 + (Exp.TrackNr * 88), tVar.lLong
            ElseIf Rec = RecFile Then
                Put #Exp.F1FileNum, 115 + (Exp.TrackNr * 88), tVar.lLong
            End If
        End If
    End If
End Sub

Public Sub ExportQDate(ByVal Rec As RecEnum)
    Read = ReadINI("Track " & Exp.TrackNr + 1, "QDate", TempFile)
    If Read <> "" Then
        tVar.lLong = DateDiff("d", "1978-01-01", Read)
        If tVar.lLong < 0 Then tVar.lLong = 0
        If tVar.lLong > 32767 Then
            tVar.iInt = tVar.lLong - 65535
        Else
            tVar.iInt = tVar.lLong
        End If
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 692 + (Exp.TrackNr * 88), tVar.iInt
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 75 + (Exp.TrackNr * 88), tVar.iInt
        End If
    End If
End Sub

Public Sub ExportRDate(ByVal Rec As RecEnum)
    Read = ReadINI("Track " & Exp.TrackNr + 1, "RDate", TempFile)
    If Read <> "" Then
        tVar.lLong = DateDiff("d", "1978-01-01", Read)
        If tVar.lLong < 0 Then tVar.lLong = 0
        If tVar.lLong > 32767 Then
            tVar.iInt = tVar.lLong - 65535
        Else
            tVar.iInt = tVar.lLong
        End If
        If Rec = F1gstate Then
            Put #Exp.F1FileNum, 736 + (Exp.TrackNr * 88), tVar.iInt
        ElseIf Rec = RecFile Then
            Put #Exp.F1FileNum, 119 + (Exp.TrackNr * 88), tVar.iInt
        End If
    End If
End Sub

Public Sub ExportDos()
    FileNum = FreeFile
    Open ProgramDir & "\Bat\Export.bat" For Append As FileNum
    Read = ReadINI("Misc", "EXEPath", TempFile)
    Read2 = oFile.GetShortName(Read)
    Read = oFile.GetShortName(Gp2Dir)
    Read3 = ReadINI("Misc", "EXE", TempFile)
    If Read3 = "" Then
        Read = Read2 & " " & Read
        Print #FileNum, Read
    Else
        Read = Read2 & " " & Read & " " & Read3
        Print #FileNum, Read
    End If
    Close FileNum
End Sub

Public Sub ExportCCFuel()
'*************************************
'Function Name: ExportCCFuel
'Use: Export CC Fuel
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-10-29
'*************************************
Dim b As Integer
    Read = ReadINI("Misc", "CCFuel", TempFile)
    If Read = 1 Then
        b = 235
    ElseIf Read = 0 Then
        b = 116
    End If
    Put #Exp.Gp2FileNum, oData.CCFuel, b
End Sub
