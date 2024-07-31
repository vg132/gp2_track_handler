Attribute VB_Name = "TrackData"
Private Found As Boolean

Public Function ReadGP2Info(ByVal Path As String) As Boolean
Dim Data As String
Dim Start As Integer
Dim Stopp As Integer

    FileNum = FreeFile
    Open Path For Binary As FileNum
    Read = String(4000, " ")
    Get #FileNum, 1, Read
    Close FileNum
    If InStr(1, Read, "#GP2INFO") Then
        With frmMain
            ReadGP2Info = True
    
            X = InStr(1, Read, "|Name|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .lblTrackName = Mid(Read, Start, Stopp - Start)
            Else
                .lblTrackName = ""
            End If
    
            X = InStr(1, Read, "|Country|")
            If X > 0 Then
                Start = X + 9
                Stopp = InStr(Start, Read, "|")
                .lblCountry = Mid(Read, Start, Stopp - Start)
            Else
                .lblCountry = ""
            End If
    
            X = InStr(1, Read, "|Year|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .lblInfoYear = Mid(Read, Start, Stopp - Start)
            Else
                .lblInfoYear = ""
            End If
    
            X = InStr(1, Read, "|Author|")
            If X > 0 Then
                Start = X + 8
                Stopp = InStr(Start, Read, "|")
                .lblAuthor = Mid(Read, Start, Stopp - Start)
            Else
                .lblAuthor = ""
            End If
    
            X = InStr(1, Read, "|Laps|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .lblLaps = Mid(Read, Start, Stopp - Start)
            Else
                .lblLaps = ""
            End If
    
            X = InStr(1, Read, "|Tyre|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .lblWare = Mid(Read, Start, Stopp - Start)
            Else
                .lblWare = ""
            End If
    
            X = InStr(1, Read, "|LengthMeters|")
            If X > 0 Then
                Start = X + 14
                Stopp = InStr(Start, Read, "|")
                .lblLen = Mid(Read, Start, Stopp - Start)
            Else
                .lblLen = ""
            End If
    
            X = InStr(1, Read, "|LapRecord|")
            If X > 0 Then
                Start = X + 11
                Stopp = InStr(Start, Read, "|")
                Read2 = Mid(Read, Start, Stopp - Start)
                If Read2 <> "None Entered" Then
                    .lblRace = Read2
                Else
                    .lblRace = ""
                End If
            Else
                .lblRace = ""
            End If
    
            X = InStr(1, Read, "|LapRecordQualify|")
            If X > 0 Then
                Start = X + 18
                Stopp = InStr(Start, Read, "|")
                Read2 = Mid(Read, Start, Stopp - Start)
                If Read2 <> "None Entered" Then
                    .lblQual = Read2
                Else
                    .lblQual = ""
                End If
            Else
                .lblQual = ""
            End If
    
            X = InStr(1, Read, "|Slot|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .lblSlot = Mid(Read, Start, Stopp - Start)
            Else
                .lblSlot = ""
            End If
    
            X = InStr(1, Read, "|Event|")
            If X > 0 Then
                Start = X + 7
                Stopp = InStr(Start, Read, "|")
                .lblEvent = Mid(Read, Start, Stopp - Start)
            Else
                .lblEvent = ""
            End If

            X = InStr(1, Read, "|Desc|")
            If X > 0 Then
                Start = X + 6
                Stopp = InStr(Start, Read, "|")
                .lblMisc = Mid(Read, Start, Stopp - Start)
            Else
                .lblMisc = ""
            End If

            .lstFile.OLEDragMode = ccOLEDragAutomatic
        End With
    Else
        OldTrackFile Path
        If Found = False Then
            FileNum = FreeFile
            Open Path For Binary As FileNum
            Read = String(3, " ")
            X = FileLen(Path) - 7
            Get #FileNum, X, Read
            Close FileNum
            If Read = "jam" Then
                Var.iInt1 = MsgBox(LoadResString(129) & vbLf & LoadResString(130), vbYesNo, TH)
                If Var.iInt1 = vbYes Then frmMain.cmdSaveGP2Info_Click
                frmMain.lstFile.OLEDragMode = ccOLEDragManual
                GoTo TrackFile
            Else
                frmMain.mnuCCCarSetup.Enabled = False
                frmMain.mnuTrackSettings.Enabled = False
                frmMain.lstFile.OLEDragMode = ccOLEDragManual
                ReadGP2Info = False
            End If
        Else
TrackFile:
            frmMain.mnuCCCarSetup.Enabled = True
            frmMain.mnuTrackSettings.Enabled = True
            frmMain.lstFile.OLEDragMode = ccOLEDragAutomatic
            ReadGP2Info = True
            Exit Function
        End If
    End If
End Function


Private Function OldTrackFile(ByVal FilePath As String)
Dim TrackSize As String
    TrackSize = FileLen(FilePath)

    If TrackSize = 32406 Then
        frmMain.lblTrackName = "Interlagos"
        frmMain.lblCountry = "Brazil"
        Found = True
        frmMain.lblLaps = "71"
        frmMain.lblLen = "4325"
        frmMain.lblWare = "20140"
        frmMain.lblQual = "1:15.962"
        frmMain.lblRace = "1:18.455"
        Samma
    ElseIf TrackSize = 32506 Then
        frmMain.lblTrackName = "TI Circuit Aida"
        frmMain.lblCountry = "Japan"
        Found = True
        frmMain.lblLaps = "83"
        frmMain.lblLen = "3723"
        frmMain.lblWare = "30746"
        frmMain.lblQual = "1:10.218"
        frmMain.lblRace = "1:14.023"
        Samma
    ElseIf TrackSize = 37678 Then
        frmMain.lblTrackName = "Imola"
        frmMain.lblCountry = "San Marino"
        Found = True
        frmMain.lblLaps = "61"
        frmMain.lblLen = "5040"
        frmMain.lblWare = "23496"
        frmMain.lblQual = "1:21.548"
        frmMain.lblRace = "1:24.438"
        Samma
    ElseIf TrackSize = 58290 Then
        frmMain.lblTrackName = "Monte-Carlo"
        frmMain.lblCountry = "Monaco"
        Found = True
        frmMain.lblLaps = "78"
        frmMain.lblLen = "3328"
        frmMain.lblWare = "32384"
        frmMain.lblQual = "1:18.560"
        frmMain.lblRace = "1:21.078"
        Samma
    ElseIf TrackSize = 34061 Then
        frmMain.lblTrackName = "Barcelona"
        frmMain.lblCountry = "Spain"
        Found = True
        frmMain.lblLaps = "65"
        frmMain.lblLen = "4747"
        frmMain.lblWare = "21237"
        frmMain.lblQual = "1:21.908"
        frmMain.lblRace = "1:25.155"
        Samma
    ElseIf TrackSize = 34392 Then
        frmMain.lblTrackName = "Gilles Villeneuve"
        frmMain.lblCountry = "Canada"
        Found = True
        frmMain.lblLaps = "69"
        frmMain.lblLen = "4430"
        frmMain.lblWare = "22009"
        frmMain.lblQual = "1:26.178"
        frmMain.lblRace = "1:28.927"
        Samma
    ElseIf TrackSize = 31767 Then
        frmMain.lblTrackName = "Magny Cours"
        frmMain.lblCountry = "France"
        Found = True
        frmMain.lblLaps = "72"
        frmMain.lblLen = "4271"
        frmMain.lblWare = "21994"
        frmMain.lblQual = "1:16.282"
        frmMain.lblRace = "1:19.678"
        Samma
    ElseIf TrackSize = 38617 Then
        frmMain.lblTrackName = "Silverstone"
        frmMain.lblCountry = "Great Britain"
        Found = True
        frmMain.lblLaps = "60"
        frmMain.lblLen = "5153"
        frmMain.lblWare = "21012"
        frmMain.lblQual = "1:24.960"
        frmMain.lblRace = "1:27.100"
        Samma
    ElseIf TrackSize = 31876 Then
        frmMain.lblTrackName = "Hockenheim"
        frmMain.lblCountry = "Germany"
        Found = True
        frmMain.lblLaps = "45"
        frmMain.lblLen = "6802"
        frmMain.lblWare = "15215"
        frmMain.lblQual = "1:43.582"
        frmMain.lblRace = "1:46.211"
        Samma
    ElseIf TrackSize = 34956 Then
        frmMain.lblTrackName = "Hungaroring"
        frmMain.lblCountry = "Hungary"
        Found = True
        frmMain.lblLaps = "77"
        frmMain.lblLen = "3968"
        frmMain.lblWare = "21310"
        frmMain.lblQual = "1:18.258"
        frmMain.lblRace = "1:20.881"
        Samma
    ElseIf TrackSize = 45598 Then
        frmMain.lblTrackName = "Spa-Francorchamps"
        frmMain.lblCountry = "Belgium"
        Found = True
        frmMain.lblLaps = "44"
        frmMain.lblLen = "6940"
        frmMain.lblWare = "25892"
        frmMain.lblQual = "2:21.163"
        frmMain.lblRace = "1:57.117"
        Samma
    ElseIf TrackSize = 41038 Then
        frmMain.lblTrackName = "Monza"
        frmMain.lblCountry = "Italy"
        Found = True
        frmMain.lblLaps = "53"
        frmMain.lblLen = "5800"
        frmMain.lblWare = "16570"
        frmMain.lblQual = "1:23.844"
        frmMain.lblRace = "1:25.930"
        Samma
    ElseIf TrackSize = 37263 Then
        frmMain.lblTrackName = "Estoril"
        frmMain.lblCountry = "Portugal"
        Found = True
        frmMain.lblLaps = "71"
        frmMain.lblLen = "4350"
        frmMain.lblWare = "17000"
        frmMain.lblQual = "1:20.608"
        frmMain.lblRace = "1:22.446"
        Samma
    ElseIf TrackSize = 33059 Then
        frmMain.lblTrackName = "Jerez"
        frmMain.lblCountry = "Europe"
        Found = True
        frmMain.lblLaps = "69"
        frmMain.lblLen = "4428"
        frmMain.lblWare = "24952"
        frmMain.lblQual = "1:22.762"
        frmMain.lblRace = "1:25.040"
        Samma
    ElseIf TrackSize = 35730 Then
        frmMain.lblTrackName = "Suzuka"
        frmMain.lblCountry = "Japan"
        Found = True
        frmMain.lblLaps = "53"
        frmMain.lblLen = "5859"
        frmMain.lblWare = "23703"
        frmMain.lblQual = "1:37.209"
        frmMain.lblRace = "1:56.597"
        Samma
    ElseIf TrackSize = 44586 Then
        frmMain.lblTrackName = "Adelaide"
        frmMain.lblCountry = "Australia"
        Found = True
        frmMain.lblLaps = "81"
        frmMain.lblLen = "3780"
        frmMain.lblWare = "20054"
        frmMain.lblQual = "1:16.179"
        frmMain.lblRace = "1:17.140"
        Samma
    ElseIf TrackSize = 40888 Then
        frmMain.lblTrackName = "Brands Hatch"
        frmMain.lblCountry = "England"
        Found = True
        frmMain.lblLaps = ""
        frmMain.lblLen = "4216"
    ElseIf TrackSize = 40844 Then
        frmMain.lblTrackName = "Buenos Aires"
        frmMain.lblCountry = "Argentina"
        frmMain.lblYear = "Year: 1996"
        Found = True
        frmMain.lblLaps = "72"
        frmMain.lblLen = "4282"
    ElseIf TrackSize = 37406 Then
        frmMain.lblTrackName = "Imola"
        frmMain.lblCountry = "San Marino"
        frmMain.lblYear = "Year: 1996"
        Found = True
        frmMain.lblLaps = "63"
        frmMain.lblLen = "5150"
    ElseIf TrackSize = 33789 Then
        frmMain.lblTrackName = "Barcelona"
        frmMain.lblCountry = "Spain"
        frmMain.lblYear = "Year: 1996"
        Found = True
        frmMain.lblLaps = "65"
        frmMain.lblLen = "4701"
    ElseIf TrackSize = 40812 Then
        FileNum2 = FreeFile
        Open Path For Binary As FileNum2
        Read = String(1, " ")
        Get #FileNum2, 40810, Read
        Close FileNum2
        If Read = "Z" Then
            frmMain.lblTrackName = "Paul Ricard"
            frmMain.lblCountry = "France"
            Found = True
            frmMain.lblLaps = ""
            frmMain.lblLen = "3798"
        End If
    ElseIf TrackSize = 38127 Then
        frmMain.lblTrackName = "Silverstone"
        frmMain.lblCountry = "England"
        Found = True
        frmMain.lblLaps = "61"
        frmMain.lblLen = "5238"
    ElseIf TrackSize = 40882 Then
        frmMain.lblTrackName = "A1-Ring"
        frmMain.lblCountry = "Austria"
        Found = True
        frmMain.lblLaps = "71"
        frmMain.lblLen = "4267"
    ElseIf TrackSize = 40812 Then
        frmMain.lblTrackName = "Zandvoort"
        frmMain.lblCountry = "Nederlands"
        Found = True
        frmMain.lblLaps = ""
        frmMain.lblLen = "2487"
    ElseIf TrackSize = 41018 Then
        frmMain.lblTrackName = "Nürburgring"
        frmMain.lblCountry = "Germany"
        Found = True
        frmMain.lblLaps = "68"
        frmMain.lblLen = "4214"
    ElseIf TrackSize = 32843 Then
        frmMain.lblTrackName = "Donington"
        frmMain.lblCountry = "England"
        Found = True
        frmMain.lblLaps = ""
        frmMain.lblLen = "4023"
    ElseIf TrackSize = 40848 Then
        frmMain.lblTrackName = "Melbourne"
        frmMain.lblCountry = "Australia"
        Found = True
        frmMain.lblLaps = "58"
        frmMain.lblLen = "5218"
    Else
        Found = False
    End If
Exit Function

ErrorTrap:
    MsgBox "Error # " + Str(Err.Number) + Err.Description
End Function

Public Function GetAdjectiv(ByVal TrackName As String) As String
    TrackName = LCase(Trim(TrackName))
    If TrackName = "australia" Then
        GetAdjectiv = "Australian"
    ElseIf TrackName = "brazil" Then GetAdjectiv = "Brazilian"
    ElseIf TrackName = "japan" Then GetAdjectiv = "Japanese"
    ElseIf TrackName = "austria" Then GetAdjectiv = "Austrian"
    ElseIf TrackName = "sweden" Then GetAdjectiv = "Swedish"
    ElseIf TrackName = "china" Then GetAdjectiv = "Chinese"
    ElseIf TrackName = "america" Then GetAdjectiv = "American"
    ElseIf TrackName = "europe" Then GetAdjectiv = "European"
    ElseIf TrackName = "portugal" Then GetAdjectiv = "Portuguese"
    ElseIf TrackName = "italy" Then GetAdjectiv = "Italian"
    ElseIf TrackName = "belgium" Then GetAdjectiv = "Belgian"
    ElseIf TrackName = "hungary" Then GetAdjectiv = "Hungarian"
    ElseIf TrackName = "germany" Then GetAdjectiv = "German"
    ElseIf TrackName = "great Britain" Then GetAdjectiv = "British"
    ElseIf TrackName = "canada" Then GetAdjectiv = "Canadian"
    ElseIf TrackName = "france" Then GetAdjectiv = "French"
    ElseIf TrackName = "spain" Then GetAdjectiv = "Spanish"
    ElseIf TrackName = "monaco" Then GetAdjectiv = "Monaco"
    ElseIf TrackName = "san marino" Then GetAdjectiv = "San Marino"
    ElseIf TrackName = "pacific" Then GetAdjectiv = "Pacific"
    ElseIf TrackName = "mexico" Then GetAdjectiv = "Mexican"
    ElseIf TrackName = "finland" Then GetAdjectiv = "Finnish"
    ElseIf TrackName = "holland" Then GetAdjectiv = "Dutch"
    ElseIf TrackName = "netherlands" Then GetAdjectiv = "Dutch"
    ElseIf TrackName = "netherland" Then GetAdjectiv = "Dutch"
    ElseIf TrackName = "south africa" Then GetAdjectiv = "South African"
    ElseIf TrackName = "usa" Then GetAdjectiv = "USA"
    ElseIf TrackName = "england" Then GetAdjectiv = "British"
    ElseIf TrackName = "great britain" Then GetAdjectiv = "British"
    ElseIf TrackName = "argentine" Then GetAdjectiv = "Argentine"
    ElseIf TrackName = "argentina" Then GetAdjectiv = "Argentine"
    ElseIf TrackName = "schweiz" Then GetAdjectiv = "Swiss"
    ElseIf TrackName = "mother earth" Then GetAdjectiv = "Mother Earth"
    ElseIf TrackName = "luxembourg (germany)" Then GetAdjectiv = "Luxembourg"
    ElseIf TrackName = "luxembourg" Then GetAdjectiv = "Luxembourg"
    ElseIf TrackName = "malysia" Then GetAdjectiv = "Malysian"
    ElseIf TrackName = "malaysia" Then GetAdjectiv = "Malysian"
    ElseIf TrackName = "switzerland" Then GetAdjectiv = "Swiss"
    ElseIf TrackName = "valhalla" Then GetAdjectiv = "Valhalla"
    End If
End Function

Private Sub Samma()
    frmMain.lblInfoYear = "1994"
    frmMain.lblEvent = "Formula 1"
    frmMain.lblAuthor = "Microprose"
End Sub
