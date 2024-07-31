Attribute VB_Name = "EditTrackInfo"
Option Explicit
Dim Read4 As String
Dim Read3 As String
Dim Read2 As String
Dim Read As String
Dim Path As String
Dim FileNum As Integer
Dim X As Long
Dim PixX As Long

Public Sub GetTrackText(Path As String)
    Read = ""
    Read2 = ""
    FileNum = FreeFile
    Open Path For Binary As FileNum
    Read = String(14, " ")
    Read2 = UCase("#GP2INFO|Name|")
    Get #FileNum, 1, Read
    If UCase(Read) = UCase(Read2) Then
        Read4 = String(2000, " ")
        Get #FileNum, 1, Read4
        Close FileNum

        'Get Track Name
        Read2 = "|Name|"
        X = 8
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 6)
            X = X + 1
        Loop
        X = X + 5
        Read2 = ""
        Do Until Read = "|"
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        MDIForm1.txteTrackName = Read2
        
        'Get Country
        If X > 1000 Then X = 1
        PicX = X
        Read2 = "Country|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 8)
            X = X + 1
        Loop
        X = X + 7
        Read2 = ""
        Do Until Read = "|"
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        MDIForm1.txteCountry.Text = Read2
        
        'Get Author
        If X > 1000 Then X = PicX
        PicX = X
        Read2 = "Author|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 7)
            X = X + 1
        Loop
        X = X + 6
        Read2 = ""
        Do Until Read = "|"
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        TrackEdit.Mid1 = "|Author|" & Read2
        
        'Get year
        If X > 1000 Then X = PicX
        PicX = X
        Read2 = "Year|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 5)
            X = X + 1
        Loop
        X = X + 4
        Read2 = ""
        Do Until Read = "|"
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        TrackEdit.Mid1 = Trim(TrackEdit.Mid1) & "|Year|" & Read2
        
        'Get event
        If X > 1000 Then X = PicX
        PicX = X
        Read2 = "Event|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 6)
            X = X + 1
        Loop
        X = X + 5
        Read2 = ""
        Do Until Read = "|"
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        TrackEdit.Mid1 = Trim(TrackEdit.Mid1) & "|Event|" & Read2
        
        'Get Description
        If X > 1000 Then X = PicX
        PicX = X
        Read2 = "Desc|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 5)
            X = X + 1
        Loop
        X = X + 4
        Read2 = ""
        Do Until Read = "|"
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        TrackEdit.Mid1 = Trim(TrackEdit.Mid1) & "|Desc|" & Read2
        
        'Get Nr of Laps
        If X > 1000 Then X = PicX
        PicX = X
        Read2 = "Laps|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 5)
            X = X + 1
        Loop
        X = X + 4
        Read2 = ""
        Do Until Read = "|"
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        MDIForm1.txteLaps.Text = Read2

        'Get Slot Nr
        If X > 1000 Then X = PicX
        PicX = X
        Read2 = "Slot|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 5)
            X = X + 1
        Loop
        X = X + 4
        Read2 = ""
        Do Until Read = "|"
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        TrackEdit.Mid2 = "|Slot|" & Read2

        'Get Tyre Ware
        If X > 1000 Then X = PicX
        PicX = X
        Read2 = "Tyre|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 5)
            X = X + 1
        Loop
        X = X + 4
        Read2 = ""
        Do Until Read = "|" Or X > 1000
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        MDIForm1.txteWare.Text = Read2
        
        'Get Track Length
        If X > 1000 Then X = PicX
        PicX = X
        Read2 = "LengthMeters|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 13)
            X = X + 1
        Loop
        X = X + 12
        Read2 = ""
        Do Until Read = "|" Or X > 1000
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        MDIForm1.txteLen = Read2

        'Get Race Lap Time
        Read2 = "LapRecord|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 10)
            X = X + 1
        Loop
        X = X + 9
        Read = String(1, " ")
        Read2 = ""
        Do Until Read = "|" Or X > 1000
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        If Read2 <> "None Entered" Then
            MDIForm1.txteRLap.Text = Read2
        Else
            MDIForm1.txteRLap.Text = ""
        End If
        
        'Get Qual Lap Time
        Read2 = "LapRecordQualify|"
        Do Until Read = Read2 Or X > 1000
            Read = Mid(Read4, X, 17)
            X = X + 1
        Loop
        X = X + 16
        Read = String(1, " ")
        Read2 = ""
        Do Until Read = "|" Or X > 1000
            Read = Mid(Read4, X, 1)
            If Read <> "|" Then Read2 = Read2 + Read
            X = X + 1
        Loop
        If Read2 <> "None Entered" Then
            MDIForm1.txteQLap.Text = Read2
        Else
            MDIForm1.txteQLap.Text = ""
        End If
    Else
        MsgBox "This track file is not supported by GP2 Track Handler.", vbInformation, GP2TH
    End If
End Sub

Public Sub SaveTrackText(ByVal SaveFileName As String, ByVal SaveFilePath As String)
Dim FileText As String
Dim RLap As String
Dim QLap As String
Dim RetVal

    If MDIForm1.txteQLap = "" Then
        QLap = "|LapRecordQualify|None Entered"
    Else
        QLap = "|LapRecordQualify|" & MDIForm1.txteQLap
    End If
    If MDIForm1.txteRLap = "" Then
        RLap = "|LapRecord|None Entered"
    Else
        RLap = "|LapRecord|" & MDIForm1.txteRLap
    End If

    FileText = ""
    
    FileText = "#GP2INFO|Name|" & MDIForm1.txteTrackName.Text & "|Country|" & MDIForm1.txteCountry.Text & "|Created|Created by Track Editor written by Paul Hoad see (License.txt about distributing this track)" & TrackEdit.Mid1 & "|Laps|" & MDIForm1.txteLaps.Text & TrackEdit.Mid2 & "|Tyre|" & MDIForm1.txteWare.Text & "|LengthMeters|" & MDIForm1.txteLen.Text & QLap & RLap & "|"
    
    X = 4000 - Len(FileText)
    Read = String(X, Chr(0))
    FileText = FileText & Read
    
    Read = oMisc.File_Exists("c:\command.com")
    If Read = True Then
        FileNum = FreeFile
        Open SaveFilePath & "\" & SaveFileName For Binary As FileNum
        Put #FileNum, 1, FileText
        Close FileNum
        FileCopy ProgramDir & "\gp2utils\check.exe", SaveFilePath & "\abcd.exe"
        RetVal = Shell("c:\command.com/c " & SaveFilePath & "\abcd.exe " & SaveFileName)
    End If
    FileText = ""
End Sub
