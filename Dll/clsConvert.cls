VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Private S1 As SaveFileInfo
Private S2 As SaveFileInfo2
Private LapTime As FastLap
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

'--Fil struktur för 1.3 time database
Private Type FastLap
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
        oMisc.WriteINI "Track " + Trim(Str(X)), "TPath", Read, NewPath

        MemFile = Mid(MemFile, 257)

        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        oMisc.WriteINI "Track " + Trim(Str(X)), "Name", Read, NewPath

        MemFile = Mid(MemFile, 28)

        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        oMisc.WriteINI "Track " + Trim(Str(X)), "Country", Read, NewPath
        
        MemFile = Mid(MemFile, 55)
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        oMisc.WriteINI "Track " + Trim(Str(X)), "Adjective", Read, NewPath

        MemFile = Mid(MemFile, 28)

        tmpInt = Asc(Mid(MemFile, 2, 1)) * 256
        tmpInt = tmpInt + Asc(Mid(MemFile, 1, 1))
        If tmpInt < 0 Then
            tImp.lLong = tmpInt + 65535
        Else
            tImp.lLong = tmpInt
        End If
        TempDouble = Round(tImp.lLong / 3.28212677519917, 0)
        oMisc.WriteINI "Track " + Trim(Str(X)), "Length", TempDouble, NewPath
        
        MemFile = Mid(MemFile, 3)
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        oMisc.WriteINI "Track " + Trim(Str(X)), "Laps", tmpInt, NewPath

        MemFile = Mid(MemFile, 3)

        tmpInt = tmpInt + Asc(Mid(MemFile, 2, 1)) * 256
        tmpInt = tmpInt + Asc(Mid(MemFile, 1, 1))
        oMisc.WriteINI "Track " + Trim(Str(X)), "Ware", tmpInt, NewPath
        
        MemFile = Mid(MemFile, 13)
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        oMisc.WriteINI "Track " + Trim(Str(X)), "BPic", Read, NewPath
        
        MemFile = Mid(MemFile, 261)
        
        tmpInt = Asc(Mid(MemFile, 1, 1))
        Read = Mid(MemFile, 2, tmpInt)
        oMisc.WriteINI "Track " + Trim(Str(X)), "SPic", Read, NewPath
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
