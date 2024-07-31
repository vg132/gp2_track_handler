Attribute VB_Name = "modConvert"
Private S1 As SaveFileInfo
Private S2 As SaveFileInfo2
Private LapTime As FastLap
Private RecordLen

'--Fil Struktur 1.0/1.1
Type SaveFileInfo
    Path As String * 200
    Country As String * 20
    Country2 As String * 20
    Track As String * 30
    Laps As String * 3
    Pic As String * 200
    Pic2 As String * 200
    Length As String * 4
    EXE As String * 16
    CarSet As String * 200
End Type

'--Fil Struktur 1.2
Type SaveFileInfo2
    Track As String * 22
    Country As String * 22
    Country2 As String * 22
    Laps As String * 3
    Ware As String * 5
    Pic As String * 100
    Pic2 As String * 100
    Path As String * 100
    CarSet As String * 100
    Length As String * 4
    Points As String * 52
    EXE As String * 48
End Type

'--Fil struktur för 1.3 time database
Type FastLap
    Track As String * 22
    QTime As String * 8
    RTime As String * 8
    QTeam As String * 12
    RTeam As String * 12
    QDriver As String * 22
    RDriver As String * 22
    QDate As String * 10
    RDate As String * 10
End Type

Public Sub WinTrack2TH(ByVal WinPath As String, ByVal OutFile As String)
Dim Path As Integer
Dim Name As Integer
Dim Adj As Integer
Dim Laps As Integer
Dim Ware As Integer
Dim TLen As Integer
Dim Land As Integer
Dim BPic As Integer
Dim SPic As Integer

    Read = ""
    Read2 = ""
    Path = 1
    Name = 257
    Land = 284
    Adj = 338
    TLen = 365 '2
    Laps = 367 '2
    Ware = 369 '2
    BPic = 381
    SPic = 641

    FileNum = FreeFile
    Open WinPath For Binary As FileNum
    Read = String(1, " ")
    For X = 1 To 16
        Get #FileNum, Path, Read
        Read2 = String(Asc(Read), " ")
        Get #FileNum, Path + 1, Read2
        oMisc.WriteINI "Track " & X, "TPath", Read2, OutFile

        Get #FileNum, Name, Read
        Read2 = String(Asc(Read), " ")
        Get #FileNum, Name + 1, Read2
        oMisc.WriteINI "Track " & X, "Name", Read2, OutFile
        
        Get #FileNum, Adj, Read
        Read2 = String(Asc(Read), " ")
        Get #FileNum, Adj + 1, Read2
        oMisc.WriteINI "Track " & X, "Adjective", Read2, OutFile
        
        Get #FileNum, Land, Read
        Read2 = String(Asc(Read), " ")
        Get #FileNum, Land + 1, Read2
        oMisc.WriteINI "Track " & X, "Country", Read2, OutFile

        Get #FileNum, Laps, Read
        Read2 = Asc(Read)
        Get #FileNum, Laps + 1, Read
        Read3 = Asc(Read)
        Read2 = (Read3 * 256) + Read2
        oMisc.WriteINI "Track " & X, "Laps", Read2, OutFile

        Get #FileNum, Ware, Read
        Read2 = Asc(Read)
        Get #FileNum, Ware + 1, Read
        Read3 = Asc(Read)
        Read2 = (Read3 * 256) + Read2
        oMisc.WriteINI "Track " & X, "Ware", Read2, OutFile

        Get #FileNum, BPic, Read
        Read2 = String(Asc(Read), " ")
        Get #FileNum, BPic + 1, Read2
        oMisc.WriteINI "Track " & X, "BPic", Read2, OutFile

        Get #FileNum, SPic, Read
        Read2 = String(Asc(Read), " ")
        Get #FileNum, SPic + 1, Read2
        oMisc.WriteINI "Track " & X, "SPic", Read2, OutFile

        Get #FileNum, TLen, Read
        Count1 = Asc(Read)
        Get #FileNum, TLen + 1, Read
        Count2 = Asc(Read)
        Read = (Count1 / 3.33333) + (Count2 * 78)
        oMisc.WriteINI "Track " & X, "Length", Read, OutFile
        
        Path = Path + 896
        Name = Name + 896
        Adj = Adj + 896
        Laps = Laps + 896
        Ware = Ware + 896
        TLen = TLen + 896
        Land = Land + 896
        BPic = BPic + 896
        SPic = SPic + 896
    Next
    Close FileNum
End Sub

Public Sub Conv1(ByVal OldPath As String, ByVal NewPath As String)
    RecordLen = Len(S1)
    FileNum = FreeFile
    Open OldPath For Random As FileNum Len = RecordLen

    For X = 1 To 16
        Get #FileNum, X, S1
        If Trim(S1.Country2) <> "No Data" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Adjective", Trim(S1.Country2), NewPath
        If Trim(S1.Country) <> "No Data" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Country", Trim(S1.Country), NewPath
        If Trim(S1.Laps) <> "No" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Laps", Trim(S1.Laps), NewPath
        If Trim(S1.Track) <> "No Data" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Name", Trim(S1.Track), NewPath
        If Trim(S1.Pic) <> "" Then oMisc.WriteINI "Track " + Trim(Str(X)), "SPic", Trim(S1.Pic), NewPath
        If Trim(S1.Pic2) <> "" Then oMisc.WriteINI "Track " + Trim(Str(X)), "BPic", Trim(S1.Pic2), NewPath
        If Trim(S1.Path) <> "No Data" Then oMisc.WriteINI "Track " + Trim(Str(X)), "TPath", Trim(S1.Path), NewPath
        If Trim(S1.Length) <> "No" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Length", Trim(S1.Length), NewPath
    Next
    Close FileNum
End Sub

Public Sub Conv2(ByVal OldPath As String, ByVal NewPath As String)
    RecordLen = Len(S2)
    FileNum = FreeFile
    Open OldPath For Random As FileNum Len = RecordLen

    For X = 1 To 16
        Get #FileNum, X, S2
        If Trim(S2.Track) <> "NoTh" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Name", Trim(S2.Track), NewPath
        If Trim(S2.Country) <> "NoTh" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Country", Trim(S2.Country), NewPath
        If Trim(S2.Country2) <> "NoTh" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Adjective", Trim(S2.Country2), NewPath
        If Trim(S2.Laps) <> "NoT" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Laps", Trim(S2.Laps), NewPath
        If Trim(S2.Ware) <> "NoTh" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Ware", Trim(S2.Ware), NewPath
        If Trim(S2.Pic) <> "NoTh" Then oMisc.WriteINI "Track " + Trim(Str(X)), "SPic", Trim(S2.Pic), NewPath
        If Trim(S2.Pic2) <> "NoTh" Then oMisc.WriteINI "Track " + Trim(Str(X)), "SPic", Trim(S2.Pic2), NewPath
        If Trim(S2.Path) <> "NoTh" Then oMisc.WriteINI "Track " + Trim(Str(X)), "TPath", Trim(S2.Path), NewPath
        If Trim(S2.Length) <> "NoTh" Then oMisc.WriteINI "Track " + Trim(Str(X)), "Length", Trim(S2.Length), NewPath
    Next
    Close FileNum
End Sub

Public Sub LapTimeCon(ByVal OldPath As String)
    RecordLen = Len(LapTime)
    FileNum = FreeFile
    Open OldPath For Random As FileNum Len = RecordLen
    X = FileLen(OldPath) / RecordLen
    With LapTime
        For Count1 = 1 To X
            Get #FileNum, Count1, LapTime
            Read = .QTime & ";" & .QDate & ";Qual;" & Trim(.QDriver) & ";" & Trim(.QTeam) & ";" & Trim(.Track)
            oDB.SaveNew dbFile, Read
            Read = .RTime & ";" & .RDate & ";Race;" & Trim(.RDriver) & ";" & Trim(.RTeam) & ";" & Trim(.Track)
            oDB.SaveNew dbFile, Read
        Next
    End With
End Sub
