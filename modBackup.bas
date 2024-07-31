Attribute VB_Name = "modBackup"
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const INVALID_HANDLE_VALUE = -1
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_ALLOWUNDO = &H40
Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Public Sub BackupTrack(ByVal Track As String, ByVal FileName As String)
Dim ItemX
Dim Dirs As String
Dim BackupDef As Boolean
    Count1 = MsgBox("Do you want to backup the oridginal GP2 Jam files?", vbYesNo, TH)
    If Count1 = vbYes Then
        BackupDef = False
    Else
        BackupDef = True
    End If
    MkDir ProgramDir & "\Backup"
    FileNum = FreeFile
    Open Track For Binary As FileNum
    Read = String(2500, " ")
    X = FileLen(Track) - 2504
    Get #FileNum, X, Read
    Close FileNum
    Read2 = String(2, Chr(0))
    For X = 2500 To 1 Step -1
        If Mid(Read, X, 2) = Read2 Then Exit For
    Next
    Read = Mid(Read, X + 4)
    Count1 = 0
    Do Until Len(Read) < 5
        Count1 = Count1 + 1
        Stopp = InStr(1, UCase(Read), UCase(Chr(0)))
        
        If Stopp = 0 Then
            Read2 = Read
        Else
            Read2 = Mid(Read, 1, Stopp - 1)
        End If
        Read3 = oMisc.File_Exists(GP2Dir & "\" & Read2)
        If BackupDef = True Then
            Var.sString1 = JamFound(Read2)
        Else
            Var.sString1 = False
        End If
        If Read3 = True And Var.sString1 = False Then
            For X = Len(Read2) To 0 Step -1
                If Mid(Read2, X, 1) = "\" Then Exit For
            Next
            Read3 = oMisc.File_Exists(ProgramDir & "\Backup\" & Mid(Read2, 1, X))
            If Read3 = False Then
                CreateDir ProgramDir & "\Backup\" & Mid(Read2, 1, X)
            End If
            FileCopy GP2Dir & "\" & Read2, ProgramDir & "\Backup\" & Read2
            frmMain.Jam2Jad.ConvertFile ProgramDir & "\Backup\" & Read2, True
        End If
        If Stopp = 0 Then
            Read = 0
        Else
            Read = Mid(Read, Stopp + 1)
        End If
    Loop
    FileCopy ProgramDir & "\gp2utils\2jam32.exe", ProgramDir & "\Backup\2jam32.exe"
    FileCopy Track, ProgramDir & "\Backup\" & GetFileName(Track)
    FileNum = FreeFile
    
    Open ProgramDir & "\Backup\GP2Info.txt" For Append As FileNum
    Print #FileNum, "--------------GP2 Track Handler - Track Backup File--------------"
    Print #FileNum, ""
    Print #FileNum, "To install this track extract all Jad file's and run the 2jam32.exe. This program will convert all jad files to jam files. Then copy all the files and Directorys to your Gp2 Directory. Move the track file (*.dat) to your track directory and install the track with GP2 Track Handler or WinTrackMan."
    Print #FileNum, ""
    Print #FileNum, "Good Luck!"
    Close FileNum

    X = ShellExecute(frmMain.hwnd, "open", oMisc.GetShortName(ProgramDir & "\gp2utils\pkzip.exe"), " -prex " & oMisc.GetShortName(ProgramDir) & "\Temp.zip " & oMisc.GetShortName(ProgramDir & "\backup\") & "*.*", ProgramDir, 1)
    Read = oMisc.CloseDosPrompt("PKZIP")
    FileCopy ProgramDir & "\temp.zip", FileName
    Kill ProgramDir & "\temp.zip"
    PerformShellAction (ProgramDir & "\Backup")
End Sub

Private Function PerformShellAction(sSource As String) As Long
Dim FOF_FLAGS As Long
Dim SHFileOp As SHFILEOPSTRUCT
    sSource = sSource & Chr$(0) & Chr$(0)
    FOF_FLAGS = BuildBrowseFlags()
    With SHFileOp
        .wFunc = 3
        .pFrom = sSource
        .fFlags = FOF_FLAGS
    End With
    PerformShellAction = SHFileOperation(SHFileOp)
End Function

Private Function BuildBrowseFlags() As Long
Dim flag As Long
    flag = 0&
    flag = FOF_ALLOWUNDO
    flag = FOF_NOCONFIRMATION
    BuildBrowseFlags = flag
End Function

Private Function JamFound(ByVal JamFile As String) As Boolean
    JamFound = False
    If JamFile = "gamejams\dproad_.jam" Then
        JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\horpac.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\pavpafic.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\pac1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\pac2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\pac_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\pac10.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\pac3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\park_rt.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pacjams\land1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\barjams\horbar.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grass.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\barjams\pavbarc.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\barjams\bar1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\barjams\bar2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\barjams\bar_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hunjams\horhun.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hunjams\pavhung.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hunjams\hungary1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hunjams\hungary2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hunjams\hun_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hunjams\hun10.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hunjams\hun3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hunjams\park_rt.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hocjams\horhoc.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hocjams\t_monsp1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hocjams\t_monsp2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hocjams\t_monsp3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hocjams\pavhock.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hocjams\hoc1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\hocjams\hoc_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\porjams\horpor.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\porjams\pavport.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\porjams\por1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\porjams\por2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\porjams\por_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\porjams\por3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\porjams\park_rt.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\brzjams\horbrz.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\brzjams\pavbraz.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\brzjams\brz1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\brzjams\brz2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\brzjams\brzcrwd.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\brzjams\brz_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\brzjams\brz10.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\scrublnd.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\jezjams\horjez.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\jezjams\pavjerez.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\jezjams\jez1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\jezjams\jez2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\jezjams\jez_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\jezjams\jez10.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\jezjams\park_lt.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\magjams\pavmagny.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\magjams\hormag.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\magjams\mag1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\magjams\mag2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\magjams\mag3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\magjams\mag_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\magjams\mag10.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\magjams\mag11.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcospect.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\hormco.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\parf.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\pavmnaco.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\mset3c_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\flat1_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\mwall1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\mwall2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\herm.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\arch.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\rascasse.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\rascc.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\mco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\mco_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\mco2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mcojams\park_rt.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\hormza.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\pavmonza.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\mza1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\mza2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\mza3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\mza_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\mza_ad2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\t_monsp1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\t_monsp2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\t_monsp3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\t_monles.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\mzajams\park_rt.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\horsan.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\pavsanma.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\san0.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\san1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\san3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\san_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\san_ad2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\san40.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\park_rt.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\t_monles.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\sanjams\t_monsp1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bricsand.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\2tarmac2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\horsil.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\pavsilvr.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\t_monsp1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\t_monsp2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\silver1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\silver1a.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\silver2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\silver4.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\silver5.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\silver6.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\silver7.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\silver9.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\sil0.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\park_rt.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\siljams\sil_tar.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\horspa.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\t_monsp1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\t_monsp2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\t_monsp3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\pavspa.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\spa1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\spa2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\spa4.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\spa6.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\spajams\spa_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adljams\horadl.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adljams\pavadel.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adljams\adl1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adljams\adl2A.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adljams\adl3.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adljams\adl_ad1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dproad_.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\verg280.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\monaco22.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\ftrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\dtrees.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clumps2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\bushes.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\clouda.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\grasverg.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts1.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\adverts2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\pavement.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\landscap.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\suzjams\horsuz.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\suzjams\pavsuz.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\suzjams\suz_tex.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\suzjams\suz2.jam" Then JamFound = True
    ElseIf JamFile = "gamejams\suzjams\suz_ad1.jam" Then JamFound = True
    End If
End Function
