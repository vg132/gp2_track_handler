Attribute VB_Name = "modConvert"

Private S1 As SaveFileInfo
Private S2 As SaveFileInfo2
Private RecordLen

'--Fil Struktur 1.0/1.1
Private Type SaveFileInfo
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
Private Type SaveFileInfo2
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

Public Sub WinTrack2TH(ByVal WinPath As String, ByVal NewPath As String)
Dim tmpInt As Integer
Dim MemFile As String

    FileNum = FreeFile
    Open WinPath For Binary As FileNum
    For X = 1 To 16
        MemFile = String(896, " ")
        Get #FileNum, 1 + (896 * (X - 1)), MemFile
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        WriteINI "Track " + Trim(Str(X)), "TPath", Read, NewPath

        MemFile = Mid(MemFile, 257)

        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        WriteINI "Track " + Trim(Str(X)), "Name", Read, NewPath

        MemFile = Mid(MemFile, 28)

        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        WriteINI "Track " + Trim(Str(X)), "Country", Read, NewPath
        
        MemFile = Mid(MemFile, 55)
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        WriteINI "Track " + Trim(Str(X)), "Adjective", Read, NewPath

        MemFile = Mid(MemFile, 28)

        tmpInt = Asc(Mid(MemFile, 2, 1)) * 256
        tmpInt = tmpInt + Asc(Mid(MemFile, 1, 1))
        If tmpInt < 0 Then
            tVar.lLong = tmpInt + 65535
        Else
            tVar.lLong = tmpInt
        End If
        tVar.dDouble = Round(tVar.lLong / 3.28212677519917, 0)
        WriteINI "Track " + Trim(Str(X)), "Length", tVar.dDouble, NewPath
        
        MemFile = Mid(MemFile, 3)
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        WriteINI "Track " + Trim(Str(X)), "Laps", tmpInt, NewPath

        MemFile = Mid(MemFile, 3)

        tmpInt = tmpInt + Asc(Mid(MemFile, 2, 1)) * 256
        tmpInt = tmpInt + Asc(Mid(MemFile, 1, 1))
        WriteINI "Track " + Trim(Str(X)), "Ware", tmpInt, NewPath
        
        MemFile = Mid(MemFile, 13)
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        WriteINI "Track " + Trim(Str(X)), "BPic", Read, NewPath
        
        MemFile = Mid(MemFile, 261)
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        WriteINI "Track " + Trim(Str(X)), "SPic", Read, NewPath
    Next
    Close FileNum
End Sub

Public Sub Conv1(ByVal OldPath As String, ByVal NewPath As String)
    RecordLen = Len(S1)
    FileNum = FreeFile
    Open OldPath For Random As FileNum Len = RecordLen
    For X = 1 To 16
        Get #FileNum, X, S1
        If Trim(S1.Country2) <> "No Data" Then WriteINI "Track " + Trim(Str(X)), "Adjective", Trim(S1.Country2), NewPath
        If Trim(S1.Country) <> "No Data" Then WriteINI "Track " + Trim(Str(X)), "Country", Trim(S1.Country), NewPath
        If Trim(S1.Laps) <> "No" Then WriteINI "Track " + Trim(Str(X)), "Laps", Trim(S1.Laps), NewPath
        If Trim(S1.Track) <> "No Data" Then WriteINI "Track " + Trim(Str(X)), "Name", Trim(S1.Track), NewPath
        If Trim(S1.Pic) <> "" Then WriteINI "Track " + Trim(Str(X)), "SPic", Trim(S1.Pic), NewPath
        If Trim(S1.Pic2) <> "" Then WriteINI "Track " + Trim(Str(X)), "BPic", Trim(S1.Pic2), NewPath
        If Trim(S1.Path) <> "No Data" Then WriteINI "Track " + Trim(Str(X)), "TPath", Trim(S1.Path), NewPath
        If Trim(S1.Length) <> "No" Then WriteINI "Track " + Trim(Str(X)), "Length", Trim(S1.Length), NewPath
    Next
    Close FileNum
End Sub

Public Sub Conv2(ByVal OldPath As String, ByVal NewPath As String)
    RecordLen = Len(S2)
    FileNum = FreeFile
    Open OldPath For Random As FileNum Len = RecordLen
    For X = 1 To 16
        Get #FileNum, X, S2
        If Trim(S2.Track) <> "NoTh" Then WriteINI "Track " + Trim(Str(X)), "Name", Trim(S2.Track), NewPath
        If Trim(S2.Country) <> "NoTh" Then WriteINI "Track " + Trim(Str(X)), "Country", Trim(S2.Country), NewPath
        If Trim(S2.Country2) <> "NoTh" Then WriteINI "Track " + Trim(Str(X)), "Adjective", Trim(S2.Country2), NewPath
        If Trim(S2.Laps) <> "NoT" Then WriteINI "Track " + Trim(Str(X)), "Laps", Trim(S2.Laps), NewPath
        If Trim(S2.Ware) <> "NoTh" Then WriteINI "Track " + Trim(Str(X)), "Ware", Trim(S2.Ware), NewPath
        If Trim(S2.Pic) <> "NoTh" Then WriteINI "Track " + Trim(Str(X)), "SPic", Trim(S2.Pic), NewPath
        If Trim(S2.Pic2) <> "NoTh" Then WriteINI "Track " + Trim(Str(X)), "SPic", Trim(S2.Pic2), NewPath
        If Trim(S2.Path) <> "NoTh" Then WriteINI "Track " + Trim(Str(X)), "TPath", Trim(S2.Path), NewPath
        If Trim(S2.Length) <> "NoTh" Then WriteINI "Track " + Trim(Str(X)), "Length", Trim(S2.Length), NewPath
    Next
    Close FileNum
End Sub

