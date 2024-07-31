Attribute VB_Name = "modRecFile"
Private Sub ExportRTeam()
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

Private Sub ExportQTeam()
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

Private Sub ExportRName()
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

Private Sub ExportQName()
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

Private Sub ExportTimeToGP2(ByVal QR As ImpExpTime)
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

Private Sub ExportQDate()
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

Private Sub ExportRDate()
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
