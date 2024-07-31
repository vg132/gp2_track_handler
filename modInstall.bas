Attribute VB_Name = "modInstall"

Public Sub GetJamData(ByVal Track As String, ByVal OverWrite As Boolean)
Dim FileName As String
Dim FileFound As Boolean

    FileNum = FreeFile
    Open Track For Binary As FileNum
    Read = String(3000, " ")
    X = FileLen(Track) - 3000
    Get #FileNum, X, Read
    Close FileNum
    Start = InStr(1, UCase(Read), UCase("gamejams\"))
    Read = Mid(Read, Start, Len(Read) - Start)
    Do Until Len(Read) < 5
        Stopp = InStr(1, UCase(Read), UCase(".jam"))
        Stopp = Stopp + 3
        Read2 = Mid(Read, 1, Stopp)
        
        For X = 1 To Len(Read2)
            X = InStr(X, Read2, "\")
            If X <> 0 Then
                Y = X
            Else
                FileName = Mid(Read2, Y + 1, Len(Read2) - Y)
                FileFound = oMisc.File_Exists(GP2Dir & "\" & Read2)
                If (FileFound = True) And (OverWrite = False) Then
                    Responce = MsgBox("Overwrite " & vbLf & GP2Dir & "\" & Read2 & "?", vbYesNo, TH)
                    If Responce = vbYes Then
                        FileCopy ProgramDir & "\@Track@\" & FileName, GP2Dir & "\" & Read2
                    Else
                        Exit For
                    End If
                Else
                    FileCopy ProgramDir & "\@Track@\" & FileName, GP2Dir & "\" & Read2
                End If
            End If
        Next
        Read = Mid(Read, Stopp + 2)
    Loop
End Sub
