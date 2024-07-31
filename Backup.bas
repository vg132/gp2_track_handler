Attribute VB_Name = "modBackup"
Public Sub BackUp(ByVal Track As String, ByVal PathFileName As String, ByVal TName As String, ByVal ZipName As String)
Dim Dir2 As String
Dim FileFound As String
    
    On Error Resume Next
    Kill ProgramDir & "\Backup.bat"
    MkDir ProgramDir & "\Backup"
    MkDir ProgramDir & "\Backup\gamejams"
    FileNum2 = FreeFile
    FileCopy ProgramDir & "\gp2utils\jam2jad.exe", ProgramDir & "\Backup\jam2jad.exe"
    Open ProgramDir & "\Backup\Info.txt" For Append As FileNum2
    Print #FileNum2, "When you want to install this track run Jam2Jad.exe in every dir with Jad files and then copy the jam files and directorys to the gamejams directrory in gp2."
    Print #FileNum2, ""
    Print #FileNum2, "This backup zip was produced with GP2 Track Handler " & Date
    Close FileNum
    FileNum2 = FreeFile
    Open ProgramDir & "\Backup.bat" For Append As FileNum
    Print #FileNum2, "@echo off"
    FileNum = FreeFile
    Open Track For Binary As FileNum
    Read = String(2000, " ")
    X = FileLen(Track) - 2000
    Get #FileNum, X, Read
    Close FileNum
    Start = InStr(1, UCase(Read), UCase("gamejams\"))
    Read = Mid(Read, Start, Len(Read) - Start)
    Do Until Len(Read) < 5
        Stopp = InStr(1, UCase(Read), UCase(".jam"))
        If Stopp = 0 Then
            Exit Do
        End If
        Stopp = Stopp + 3
        Read2 = Mid(Read, 1, Stopp)
        Read3 = oMisc.File_Exists(GP2Dir & "\" & Read2)
        If Read3 = False Then
            MsgBox "This jam file was not found and will not be included: " & Read2
        Else
            For X = Len(Read2) To 1 Step -1
                If Mid(Read2, X, 1) = "\" Then Exit For
            Next
            If X > 1 Then
                Dir2 = Mid(Read2, 1, X)
            End If
            FileFound = oMisc.File_Exists(ProgramDir & "\Backup\" & Dir2)
            If FileFound = False Then
                MkDir ProgramDir & "\Backup\" & Dir2
                FileCopy ProgramDir & "\GP2Utils\Jam2Jad.exe", ProgramDir & "\Backup\" & Dir2 & "Jam2Jad.exe"
                Print #FileNum2, "cd\"
                Print #FileNum2, "cd " & ProgramDir & "\Backup\" & Dir2
                Print #FileNum2, "jam2jad"
                Print #FileNum2, "del jam2jad.exe"
            End If
            FileCopy GP2Dir & "\" & Read2, ProgramDir & "\Backup\" & Read2
        End If
        Read = Mid(Read, Stopp + 2)
    Loop
    FileCopy Track, ProgramDir & "\Backup\" & TName
    Print #FileNum2, "cd\"
    Print #FileNum2, ProgramDir & "\gp2utils\pkzip -prex " & ProgramDir & "1.zip" & " " & ProgramDir & "\backup\*.* > nul"
    Print #FileNum2, "deltree/y backup > nul"
    Print #FileNum2, "cls"
    Print #FileNum2, "echo You can close this window now!"
    Close FileNum2
    RetVal = ShellExecute(frmMain.hwnd, "open", ProgramDir & "backup.bat", vbNullString, vbNullString, 1)
End Sub
