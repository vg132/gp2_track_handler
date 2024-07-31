Attribute VB_Name = "TrackData"

Public Sub OldTrackFile()
        If X17 <> 32000 Then
            If (Mid(MDIForm1.Dir1.Path, 3, 1) = "\") And (Len(MDIForm1.Dir1.Path) = 3) Then
                FileSize = FileLen(MDIForm1.Dir1.Path + MDIForm1.File1.FileName)
            Else
                FileSize = FileLen(MDIForm1.Dir1.Path + "\" + MDIForm1.File1.FileName)
            End If
        End If
        If X17 = 32000 Then FileSize = FileLen(Read3)
        If FileSize = 32406 Then
            MDIForm1.lblName = "Interlagos"
            MDIForm1.lblCountry = "Brazil"
            Samma
            MDIForm1.lblLaps = "71"
            MDIForm1.lblLength = "4325"
            MDIForm1.lblWare = "20140"
            Exit Sub
        End If
        If FileSize = 37678 Then
            MDIForm1.lblName = "Imola"
            MDIForm1.lblCountry = "San Marino"
            Samma
            MDIForm1.lblLaps = "61"
            MDIForm1.lblLength = "5040"
            MDIForm1.lblWare = "23496"
            Exit Sub
        End If
        If FileSize = 58290 Then
            MDIForm1.lblName = "Monte-Carlo"
            MDIForm1.lblCountry = "Monaco"
            Samma
            MDIForm1.lblLaps = "78"
            MDIForm1.lblLength = "3328"
            MDIForm1.lblWare = "32384"
            Exit Sub
        End If
        If FileSize = 34061 Then
            MDIForm1.lblName = "Barcelona"
            MDIForm1.lblCountry = "Spain"
            Samma
            MDIForm1.lblLaps = "65"
            MDIForm1.lblLength = "4747"
            MDIForm1.lblWare = "21237"
            Exit Sub
        End If
        If FileSize = 34392 Then
            MDIForm1.lblName = "Gilles Villeneuve"
            MDIForm1.lblCountry = "Canada"
            Samma
            MDIForm1.lblLaps = "69"
            MDIForm1.lblLength = "4430"
            MDIForm1.lblWare = "22009"
            Exit Sub
        End If
        If FileSize = 31767 Then
            MDIForm1.lblName = "Magny Cours"
            MDIForm1.lblCountry = "France"
            Samma
            MDIForm1.lblLaps = "72"
            MDIForm1.lblLength = "4271"
            MDIForm1.lblWare = "21994"
            Exit Sub
        End If
        If FileSize = 38617 Then
            MDIForm1.lblName = "Silverstone"
            MDIForm1.lblCountry = "Great Britain"
            Samma
            MDIForm1.lblLaps = "60"
            MDIForm1.lblLength = "5153"
            MDIForm1.lblWare = "21012"
            Exit Sub
        End If
        If FileSize = 31876 Then
            MDIForm1.lblName = "Hockenheim"
            MDIForm1.lblCountry = "Germany"
            Samma
            MDIForm1.lblLaps = "45"
            MDIForm1.lblLength = "6802"
            MDIForm1.lblWare = "15215"
            Exit Sub
        End If
        If FileSize = 34956 Then
            MDIForm1.lblName = "Hungaroring"
            MDIForm1.lblCountry = "Hungary"
            Samma
            MDIForm1.lblLaps = "77"
            MDIForm1.lblLength = "3968"
            MDIForm1.lblWare = "21310"
            Exit Sub
        End If
        If FileSize = 45598 Then
            MDIForm1.lblName = "Spa-Francorchamps"
            MDIForm1.lblCountry = "Belgium"
            Samma
            MDIForm1.lblLaps = "44"
            MDIForm1.lblLength = "6940"
            MDIForm1.lblWare = "25892"
            Exit Sub
        End If
        If FileSize = 41038 Then
            MDIForm1.lblName = "Monza"
            MDIForm1.lblCountry = "Italy"
            Samma
            MDIForm1.lblLaps = "53"
            MDIForm1.lblLength = "5800"
            MDIForm1.lblWare = "16570"
            Exit Sub
        End If
        If FileSize = 37263 Then
            MDIForm1.lblName = "Estoril"
            MDIForm1.lblCountry = "Portugal"
            Samma
            MDIForm1.lblLaps = "71"
            MDIForm1.lblLength = "4350"
            MDIForm1.lblWare = "17000"
            Exit Sub
        End If
        If FileSize = 33059 Then
            MDIForm1.lblName = "Jerez"
            MDIForm1.lblCountry = "Europe"
            Samma
            MDIForm1.lblLaps = "69"
            MDIForm1.lblLength = "4428"
            MDIForm1.lblWare = "24952"
            Exit Sub
        End If
        If FileSize = 35730 Then
            MDIForm1.lblName = "Suzuka"
            MDIForm1.lblCountry = "Japan"
            Samma
            MDIForm1.lblLaps = "53"
            MDIForm1.lblLength = "5859"
            MDIForm1.lblWare = "23703"
            Exit Sub
        End If
        If FileSize = 44586 Then
            MDIForm1.lblName = "Adelaide"
            MDIForm1.lblCountry = "Australia"
            Samma
            MDIForm1.lblLaps = "81"
            MDIForm1.lblLength = "3780"
            MDIForm1.lblWare = "20054"
            Exit Sub
        End If
        If FileSize = 32506 Then
            MDIForm1.lblName = "TI Circuit Aida"
            MDIForm1.lblCountry = "Japan"
            Samma
            MDIForm1.lblLaps = "83"
            MDIForm1.lblLength = "3723"
            MDIForm1.lblWare = "30746"
            Exit Sub
        End If
        If FileSize = 40888 Then
            MDIForm1.lblName = "Brands Hatch"
            MDIForm1.lblCountry = "England"
            MDIForm1.lblYear = "Year: 1986"
            SammaIA
            MDIForm1.lblLaps = ""
            MDIForm1.lblLength = "4216"
            Exit Sub
        End If
        If FileSize = 40844 Then
            MDIForm1.lblName = "Buenos Aires"
            MDIForm1.lblCountry = "Argentina"
            MDIForm1.lblYear = "Year: 1996"
            SammaIA
            MDIForm1.lblLaps = "72"
            MDIForm1.lblLength = "4282"
            Exit Sub
        End If
        If FileSize = 37406 Then
            MDIForm1.lblName = "Imola"
            MDIForm1.lblCountry = "San Marino"
            MDIForm1.lblYear = "Year: 1996"
            SammaIA
            MDIForm1.lblLaps = "63"
            MDIForm1.lblLength = "5150"
            Exit Sub
        End If
        If FileSize = 33789 Then
            MDIForm1.lblName = "Barcelona"
            MDIForm1.lblCountry = "Spain"
            MDIForm1.lblYear = "Year: 1996"
            SammaIA
            MDIForm1.lblLaps = "65"
            MDIForm1.lblLength = "4701"
            Exit Sub
        End If
        If FileSize = 40812 Then
            FileNum2 = FreeFile
            Open MDIForm1.Dir1.Path + "\" + MDIForm1.File1.FileName For Binary As FileNum2
            Read = String(1, " ")
            Get #FileNum2, 40810, Read
            Close FileNum2
            If Read = "Z" Then
                MDIForm1.lblName = "Paul Ricard"
                MDIForm1.lblCountry = "France"
                MDIForm1.lblYear = "Year: 1988"
                SammaIA
                MDIForm1.lblLaps = ""
                MDIForm1.lblLength = "3798"
                Exit Sub
            End If
        End If
        If FileSize = 38127 Then
            MDIForm1.lblName = "Silverstone"
            MDIForm1.lblCountry = "England"
            MDIForm1.lblYear = "Year: 1996"
            SammaIA
            MDIForm1.lblLaps = "61"
            MDIForm1.lblLength = "5238"
            Exit Sub
        End If

        If FileSize = 40882 Then
            MDIForm1.lblName = "A1-Ring"
            MDIForm1.lblCountry = "Austria"
            MDIForm1.lblYear = "Year: 1997"
            SammaIA
            MDIForm1.lblLaps = "71"
            MDIForm1.lblLength = "4267"
            Exit Sub
        End If
        If FileSize = 40812 Then
            MDIForm1.lblName = "Zandvoort"
            MDIForm1.lblCountry = "Nederlands"
            MDIForm1.lblYear = "Year: 1996"
            SammaIA
            MDIForm1.lblLaps = ""
            MDIForm1.lblLength = "2487"
            Exit Sub
        End If
        If FileSize = 41018 Then
            MDIForm1.lblName = "Nürburgring"
            MDIForm1.lblCountry = "Germany"
            MDIForm1.lblYear = "Year: 1996"
            SammaIA
            MDIForm1.lblLaps = "68"
            MDIForm1.lblLength = "4214"
            Exit Sub
        End If

        If FileSize = 32843 Then
            MDIForm1.lblName = "Donington"
            MDIForm1.lblCountry = "England"
            MDIForm1.lblYear = "Year: "
            SammaIA
            MDIForm1.lblLaps = ""
            MDIForm1.lblLength = "4023"
            Exit Sub
        End If
        If FileSize = 40848 Then
            MDIForm1.lblName = "Melbourne"
            MDIForm1.lblCountry = "Australia"
            MDIForm1.lblYear = "Year: 1996"
            SammaIA
            MDIForm1.lblLaps = "58"
            MDIForm1.lblLength = "5218"
            Exit Sub
        End If
    Exit Sub
ErrorTrap:
    MsgBox "Error # " + Str(Err.Number) + Err.Description
End Sub

Public Sub Samma()
    MDIForm1.lblEvent = "Event: Formula 1"
    MDIForm1.lblYear.Caption = "Year: 1994"
    MDIForm1.lblAuthor.Caption = "Author: Microprose - Orginal Track from GP2"
    MDIForm1.lblMisc = "Misc Info: This data is based on the size of the file, and I hope that the data is right."
    Count2 = 1234
End Sub

Public Sub SammaIA()
    MDIForm1.lblEvent = "Event: Formula 1"
    MDIForm1.lblAuthor.Caption = "Author: Instant Access"
    MDIForm1.lblMisc = "Misc Info: This data is based on the size of the file, and I hope that the data is right."
    Count2 = 1234
End Sub

Public Sub GetDataFromLabel()
    Read3 = "Track " + Trim(Str(Count3))
    Read2 = ProgramDir + "\WorkCopy.lda"
    Read = oMisc.WriteINI(Read3, "Country", MDIForm1.lblCountry, Read2)
    Read = oMisc.WriteINI(Read3, "Ware", MDIForm1.lblWare, Read2)
    Read = oMisc.WriteINI(Read3, "Name", MDIForm1.lblName, Read2)
    Read = oMisc.WriteINI(Read3, "Laps", MDIForm1.lblLaps, Read2)
    Read = oMisc.WriteINI(Read3, "RTime", MDIForm1.lblRLap, Read2)
    Read = oMisc.WriteINI(Read3, "QTime", MDIForm1.lblQLap, Read2)
    Read = oMisc.WriteINI(Read3, "Length", MDIForm1.lblLength, Read2)
    If (Mid(MDIForm1.Dir1.Path, 3, 1) = "\") And (Len(MDIForm1.Dir1.Path) = 3) Then
        Read = oMisc.WriteINI(Read3, "TPath", MDIForm1.Dir1.Path + MDIForm1.File1.FileName, Read2)
    Else
        Read = oMisc.WriteINI(Read3, "TPath", MDIForm1.Dir1.Path + "\" + MDIForm1.File1.FileName, Read2)
    End If
    Read = MDIForm1.lblCountry
    GetAdjectiv
    Read = oMisc.WriteINI(Read3, "Adjective", Read, Read2)
End Sub

Public Sub NewTrack()
Dim oCtl As Control

    For Each oCtl In MDIForm1
        If TypeOf oCtl Is TextBox Then
            oCtl.Text = ""
        End If
    Next
    Set oCtl = Nothing

    Read = ""
    Read2 = ""
    Unload frmExport
    Unload frmImport
    Unload frmDosPath
    Unload frmPoint
    Unload frmAbout
    If Gp2Dir <> "" Then ImportPoints
End Sub

Public Sub GetAdjectiv()
Dim TempString2 As String
    TempString2 = ""
    If Trim(Read) = "Australia" Then TempString2 = "Australian"
    If Trim(Read) = "Brazil" Then TempString2 = "Brazilian"
    If Trim(Read) = "Japan" Then TempString2 = "Japanese"
    If Trim(Read) = "Austria" Then TempString2 = "Austrian"
    If Trim(Read) = "Sweden" Then TempString2 = "Swedish"
    If Trim(Read) = "China" Then TempString2 = "Chinese"
    If Trim(Read) = "America" Then TempString2 = "American"
    If Trim(Read) = "Europe" Then TempString2 = "European"
    If Trim(Read) = "Portugal" Then TempString2 = "Portuguese"
    If Trim(Read) = "Italy" Then TempString2 = "Italian"
    If Trim(Read) = "Belgium" Then TempString2 = "Belgian"
    If Trim(Read) = "Hungary" Then TempString2 = "Hungarian"
    If Trim(Read) = "Germany" Then TempString2 = "German"
    If Trim(Read) = "Great Britain" Then TempString2 = "British"
    If Trim(Read) = "Canada" Then TempString2 = "Canadian"
    If Trim(Read) = "France" Then TempString2 = "French"
    If Trim(Read) = "Spain" Then TempString2 = "Spanish"
    If Trim(Read) = "Monaco" Then TempString2 = "Monaco"
    If Trim(Read) = "San Marino" Then TempString2 = "San Marino"
    If Trim(Read) = "Pacific" Then TempString2 = "Pacific"
    If Trim(Read) = "Mexico" Then TempString2 = "Mexican"
    If Trim(Read) = "Finland" Then TempString2 = "Finnish"
    If Trim(Read) = "Holland" Then TempString2 = "Dutch"
    If Trim(Read) = "Netherlands" Then TempString2 = "Dutch"
    If Trim(Read) = "South Africa" Then TempString2 = "South African"
    If Trim(Read) = "USA" Then TempString2 = "USA"
    If Trim(Read) = "England" Then TempString2 = "British"
    If Trim(Read) = "Argentine" Then TempString2 = "Argentine"
    If Trim(Read) = "Schweiz" Then TempString2 = "Swiss"
    If Trim(Read) = "Mother Earth" Then TempString2 = "Mother Earth"
    If Trim(Read) = "Luxembourg (Germany)" Then TempString2 = "Luxembourg"
    If Trim(Read) = "Luxembourg" Then TempString2 = "Luxembourg"
    If Trim(Read) = "Malysia" Then TempString2 = "Malysian"
    Read = TempString2
End Sub

Public Sub ReadTrackFile(Path As String)
    MDIForm1.txtAdjectiv = ""
    MDIForm1.txtName = ""
    MDIForm1.txtCountry = ""
    MDIForm1.txtLaps = ""
    MDIForm1.txtLength = ""
    MDIForm1.txtPath.Enabled = True
    MDIForm1.txtPath = ""
    MDIForm1.txtPath.Enabled = False
    MDIForm1.txtTire = ""
    Read = ""
    Read2 = ""

    FileNum = FreeFile
    Open Path For Binary As FileNum
    Read = String(14, " ")
    Read2 = UCase("#GP2INFO|Name|")
    Get #FileNum, 1, Read
    If UCase(Read) = UCase(Read2) Then
        Read4 = String(1100, " ")
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
        MDIForm1.lblName = Read2
        
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
        MDIForm1.lblCountry = Read2
        
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
        MDIForm1.lblAuthor = "Author: " + Read2
        
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
        MDIForm1.lblYear = "Year: " + Read2
        
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
        MDIForm1.lblEvent = "Event: " + Read2
        
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
        MDIForm1.lblMisc = "Misc Info: " + Read2
        
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
        MDIForm1.lblLaps = "" + Read2
        
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
        MDIForm1.lblWare = "" + Read2
        
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
        MDIForm1.lblLength = "" + Read2

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
        If Read2 <> "None Entered" Then MDIForm1.lblRLap = Read2
        
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
        If Read2 <> "None Entered" Then MDIForm1.lblQLap = Read2

        MDIForm1.txtPath = Path
        NoSupport = False
        Read = MDIForm1.lblCountry
        GetAdjectiv
        Read4 = ""
    Else
        OldTrackFile
        Close FileNum
        If Count2 <> 1234 Then
            MsgBox "This track file is not supported by " + TH + ".", vbInformation, TH
            NoSupport = True
        Else
            NoSupport = False
            Count2 = 2589
        End If
        Exit Sub
    End If
End Sub
