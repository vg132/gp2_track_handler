Attribute VB_Name = "TrackData"
Private Type TypeInfoFile
    lblTrackName As String * 17
    lblCountry As String * 13
    lblLaps As String * 2
    lblLen As String * 4
    lblWare As String * 5
    lblQual As String * 8
    lblRace As String * 8
End Type

Public Function ReadGp2Info(ByVal Path As String) As Boolean
Dim Data As String
Dim Start As Integer
Dim Stopp As Integer

    If FileLen(Path) < 4000 Then GoTo NoTrack
    FileNum = FreeFile
    Open Path For Binary As FileNum
    Read = String(4000, " ")
    Get #FileNum, 1, Read
    Close FileNum
    If InStr(1, Read, "#GP2INFO") Then
        With TrackInfo
            ReadGp2Info = True
            X = InStr(1, Read, "|Name|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .Name = Mid(Read, Start, Stopp - Start)
            Else
                .Name = ""
            End If
    
            X = InStr(1, Read, "|Country|")
            If X > 0 Then
                Start = X + 9
                Stopp = InStr(Start, Read, "|")
                .Country = Mid(Read, Start, Stopp - Start)
            Else
                .Country = ""
            End If
    
            X = InStr(1, Read, "|Year|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .Year = Mid(Read, Start, Stopp - Start)
            Else
                .Year = ""
            End If
    
            X = InStr(1, Read, "|Author|")
            If X > 0 Then
                Start = X + 8
                Stopp = InStr(Start, Read, "|")
                .Author = Mid(Read, Start, Stopp - Start)
            Else
                .Author = ""
            End If
    
            X = InStr(1, Read, "|Laps|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .Laps = Mid(Read, Start, Stopp - Start)
            Else
                .Laps = ""
            End If
    
            X = InStr(1, Read, "|Tyre|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .Tyre = Mid(Read, Start, Stopp - Start)
            Else
                .Tyre = ""
            End If
    
            X = InStr(1, Read, "|LengthMeters|")
            If X > 0 Then
                Start = X + 14
                Stopp = InStr(Start, Read, "|")
                .LengthMeters = Mid(Read, Start, Stopp - Start)
            Else
                .LengthMeters = ""
            End If
    
            X = InStr(1, Read, "|LapRecord|")
            If X > 0 Then
                Start = X + 11
                Stopp = InStr(Start, Read, "|")
                Read2 = Mid(Read, Start, Stopp - Start)
                If Read2 <> "None Entered" Then
                    .LapRecord = Read2
                Else
                    .LapRecord = ""
                End If
            Else
                .LapRecord = ""
            End If
    
            X = InStr(1, Read, "|LapRecordQualify|")
            If X > 0 Then
                Start = X + 18
                Stopp = InStr(Start, Read, "|")
                Read2 = Mid(Read, Start, Stopp - Start)
                If Read2 <> "None Entered" Then
                    .LapRecordQualify = Read2
                Else
                    .LapRecordQualify = ""
                End If
            Else
                .LapRecordQualify = ""
            End If
    
            X = InStr(1, Read, "|Slot|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .Slot = Mid(Read, Start, Stopp - Start)
            Else
                .Slot = ""
            End If
    
            X = InStr(1, Read, "|Event|")
            If X > 0 Then
                Start = X + 7
                Stopp = InStr(Start, Read, "|")
                .Event = Mid(Read, Start, Stopp - Start)
            Else
                .Event = ""
            End If

            X = InStr(1, Read, "|Desc|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .Desc = Mid(Read, Start, Stopp - Start)
            Else
                .Desc = ""
            End If
            frmMain.lstFile.OLEDragMode = ccOLEDragAutomatic
        End With
    Else
        If OldTrackFile(Path) = False Then
            FileNum = FreeFile
            Open Path For Binary As FileNum
            Read = String(3, " ")
            X = FileLen(Path) - 7
            Get #FileNum, X, Read
            Close FileNum
            If Read = "jam" Then
                frmMain.lstFile.OLEDragMode = ccOLEDragManual
                GoTo TrackFile
            Else
NoTrack:
                frmMain.mnuCCCarSetup.Enabled = False
                frmMain.mnuTrackSettings.Enabled = False
                frmMain.lstFile.OLEDragMode = ccOLEDragManual
                ReadGp2Info = False
            End If
        Else
TrackFile:
            frmMain.mnuCCCarSetup.Enabled = True
            frmMain.mnuTrackSettings.Enabled = True
            frmMain.lstFile.OLEDragMode = ccOLEDragAutomatic
            ReadGp2Info = True
            Exit Function
        End If
    End If
End Function

Private Function OldTrackFile(ByVal FilePath As String) As Boolean
Dim TrackSize As String

    TrackSize = FileLen(FilePath)

    If TrackSize = 32406 Then
        GetTrackInfo (0)
    ElseIf TrackSize = 32506 Then
        GetTrackInfo (1)
    ElseIf TrackSize = 37678 Then
        GetTrackInfo (2)
    ElseIf TrackSize = 58290 Then
        GetTrackInfo (3)
    ElseIf TrackSize = 34061 Then
        GetTrackInfo (4)
    ElseIf TrackSize = 34392 Then
        GetTrackInfo (5)
    ElseIf TrackSize = 31767 Then
        GetTrackInfo (6)
    ElseIf TrackSize = 38617 Then
        GetTrackInfo (7)
    ElseIf TrackSize = 31876 Then
        GetTrackInfo (8)
    ElseIf TrackSize = 34956 Then
        GetTrackInfo (9)
    ElseIf TrackSize = 45598 Then
        GetTrackInfo (10)
    ElseIf TrackSize = 41038 Then
        GetTrackInfo (11)
    ElseIf TrackSize = 37263 Then
        GetTrackInfo (12)
    ElseIf TrackSize = 33059 Then
        GetTrackInfo (13)
    ElseIf TrackSize = 35730 Then
        GetTrackInfo (14)
    ElseIf TrackSize = 44586 Then
        GetTrackInfo (15)
    Else
        OldTrackFile = False
        Exit Function
    End If
    OldTrackFile = True
Exit Function

ErrorTrap:
    MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
        "Error Desctiption: " & Err.Description & vbLf & _
        "Error Source: OldTrackFile()", vbCritical, TH & " - Error"
End Function

Public Function GetAdjectiv(ByVal sCountry As String) As String
    GetAdjectiv = ReadINI("Adjectiv", Trim(sCountry), ProgramDir & "\Adjectiv.ini")
End Function

Private Sub GetTrackInfo(ByVal Nr As Integer)
Dim GetData As TypeInfoFile
Dim RecLen As Integer
    RecLen = Len(GetData)
    FileNum = FreeFile
    Open ProgramDir & "\Org.lda" For Random As FileNum Len = RecLen
    Get #FileNum, Nr + 1, GetData
    Close FileNum
    
    With TrackInfo
        .Author = "Microproce"
        .Country = GetData.lblCountry
        .Desc = "Original Gp2 Track"
        .Event = "Formula 1"
        .LapRecord = GetData.lblRace
        .LapRecordQualify = GetData.lblQual
        .Laps = GetData.lblLaps
        .LengthMeters = GetData.lblLen
        .Name = GetData.lblTrackName
        .Slot = Nr + 1
        .Tyre = GetData.lblWare
        .Year = "1994"
    End With
End Sub
