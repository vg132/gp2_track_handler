Attribute VB_Name = "modBackup"
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Private Declare Function ConvertFile Lib "ThLib.dll" (ByVal FileName As String) As Boolean
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long

Private Const INVALID_HANDLE_VALUE = -1
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_ALLOWUNDO = &H40
Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3

Private JamData As String

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Public Sub BackupTrack(ByVal Track As String, ByVal sZipName As String, ByVal BackupDef As Boolean)
Dim ItemX
Dim Dirs As String
Dim vArray As Variant
Dim iJams As Integer
Dim NewValue As Integer
    
    If oFile.FileExists(ProgramDir & "\Backup") = True Then PerformShellAction (ProgramDir & "\Backup")
    CreateDir ProgramDir & "\Backup"
    vArray = oFile.GetJamFiles(Track)
    If BackupDef = False Then LoadJams
    For iJams = 0 To UBound(vArray, 2)
        If vArray(0, iJams) = "" Then Exit For
        Read = oFile.FileExists(Gp2Dir & "\" & vArray(0, iJams))
        
        'Progress bar Show
        
        With frmProgress
            NewValue = .prbBackup.Value + Fix(100 / UBound(vArray, 2))
            If NewValue > 100 Then NewValue = 100
            .prbBackup.Value = NewValue
            .lblText = "Copying jam files from " & Gp2Dir
            .lblFileName = vArray(0, iJams)
            .Refresh
        End With
        
        'End Progress bar
        
        If BackupDef = False Then
            Read2 = JamFound(LCase(vArray(0, iJams)))
        Else
            Read2 = False
        End If
        If (Read2 <> True) And (Read = True) Then
            Read = oFile.FileExists(ProgramDir & "\Backup\" & oFile.GetFilePart(Read2, GetFileName))
            If Read = False Then
                CreateDir ProgramDir & "\Backup\" & oFile.GetFilePart(vArray(0, iJams), GetFilePath)
            End If
            FileCopy Gp2Dir & "\" & vArray(0, iJams), ProgramDir & "\Backup\" & vArray(0, iJams)
            ConvertFile ProgramDir & "\Backup\" & vArray(0, iJams)
            Kill (ProgramDir & "\Backup\" & vArray(0, iJams))
        End If
    Next
    
    FileCopy ProgramDir & "\gp2utils\2jam32.exe", ProgramDir & "\Backup\2jam32.exe"
    FileCopy Track, ProgramDir & "\Backup\" & oFile.GetFilePart(Track, GetFileName)
    FileNum = FreeFile
    
    Open ProgramDir & "\Backup\Gp2Info.txt" For Append As FileNum
    Print #FileNum, "--------------Gp2 Track Handler - Track Backup File--------------"
    Print #FileNum, ""
    Print #FileNum, "To install this track extract all Jad file's and run the 2jam32.exe. This program will convert all jad files to jam files. Then copy all the files and Directorys to your Gp2 Directory. Move the track file (*.dat) to your track directory and install the track with Gp2 Track Handler or WinTrackMan."
    Print #FileNum, ""
    Print #FileNum, "Good Luck!"
    Close FileNum
    X = ShellExecute(frmMain.hWnd, "open", oFile.GetShortName(ProgramDir & "\gp2utils\pkzip.exe"), " -prex " & oFile.GetShortName(ProgramDir) & "\Temp.zip " & oFile.GetShortName(ProgramDir & "\backup\") & "*.*", ProgramDir, 1)
    
    X = 110
    Do Until X = 0
        Sleep (250)
        X = FindWindow(CLng(0), "pkzip")
    Loop
    Sleep (50)

    FileCopy ProgramDir & "\temp.zip", sZipName
    Kill ProgramDir & "\temp.zip"
    PerformShellAction (ProgramDir & "\Backup")
    JamData = ""
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case "53"
        Close FileNum
        BackupDef = False
        Resume Next
    Case Else
        On Error Resume Next
        MsgBox "The buckup of this track faild."
        Kill ProgramDir & "\temp.zip"
        PerformShellAction (ProgramDir & "\Backup")
        BackupDef = True
    End Select
    JamData = ""
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
    flag = FOF_NOCONFIRMATION
    BuildBrowseFlags = flag
End Function

Private Function JamFound(ByVal JamFile As String) As Boolean
    JamFound = False
    JamFile = Left(JamFile, Len(JamFile) - 4)
    JamFile = Mid(JamFile, 10)
    JamFile = """" & JamFile & """"
    X = InStr(1, JamData, JamFile)
    If X <> 0 Then JamFound = True
End Function

Private Sub LoadJams()
    JamData = Space(FileLen(ProgramDir & "\Jams.lda"))
    FileNum = FreeFile
    Open ProgramDir & "\Jams.lda" For Binary As FileNum
    Get #FileNum, 1, JamData
    Close FileNum
    JamData = LCase(JamData)
End Sub

Public Sub Uninstall(ByVal Track As String)
Dim vArray As Variant
Dim i As Integer
Dim JamDir As String
    LoadJams
    vArray = oFile.GetJamFiles(Track)
    For i = 0 To UBound(vArray, 2)
        With frmProgress
            NewValue = .prbBackup.Value + Fix(100 / UBound(vArray, 2))
            If NewValue > 100 Then NewValue = 100
            .prbBackup.Value = NewValue
            .lblText = "Working directory " & Gp2Dir
            .lblFileName = "Checking: " & vArray(0, iJams)
            .Refresh
        End With

        If JamFound(LCase(vArray(0, i))) = False Then
            frmProgress.lblFileName = "Deleting: " & vArray(0, iJams)
            frmProgress.Refresh
            If oFile.FileExists(Gp2Dir & "\" & vArray(0, i)) = True Then
                Kill (Gp2Dir & "\" & vArray(0, i))
                RemoveDirectory (Gp2Dir & "\" & oFile.GetFilePart(vArray(0, i), GetFilePath))
            End If
        End If
    Next
    Kill (Track)
    frmProgress.lblFileName = "Deleting: " & oFile.GetFilePart(Track, GetFileName)
End Sub
