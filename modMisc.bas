Attribute VB_Name = "modMisc"
Private GP2Country As String

Public Sub NewTree()
    frmMain.TreeView1.Nodes.Clear
    Dim nodX As Node    ' Create variable.

    Set nodX = frmMain.TreeView1.Nodes.Add(, , "r", "GP2 Track's", 1, 2)

    X = 11
    Do Until X > 26
        Set nodX = frmMain.TreeView1.Nodes.Add("r", tvwChild, "t" + Trim(Str(X)), "Track " + Trim(Str(X - 10)), 1, 2)
        X = X + 1
    Loop
    nodX.EnsureVisible ' Show all nodes.
End Sub

Public Sub TextSelected()
Dim i As Integer
Dim oMyTextBox As Object
Set oMyTextBox = Screen.ActiveControl
    If TypeName(oMyTextBox) = "TextBox" Then
        i = Len(oMyTextBox.Text)
        oMyTextBox.SelStart = 0
        oMyTextBox.SelLength = i
    End If
End Sub

Public Sub GetGP2Version()
Dim V
'-- Läs GP2.exe och titta vad det är för version (språk)
    Read2 = "GP2 Version:"
    V = "Version 1.0b"
    FileNum = FreeFile
    Open GP2Dir + "\gp2.exe" For Binary As FileNum
    Read = String(23, " ")
    Get #FileNum, 5671742, Read
    If Read = "US English Version 1.0b" Then
        Close FileNum
        GP2V = US
        frmMain.stbMain.Panels(2) = Read2 & " American " + V
        Exit Sub
    End If
    Get #FileNum, 5671743, Read
    If Read = "UK English Version 1.0b" Then
        Close FileNum
        GP2V = UK
        frmMain.stbMain.Panels(2) = Read2 & " UK English " + V
        Exit Sub
    End If
    Get #FileNum, 5673614, Read
    If Read = "Nederlandse versie 1.0b" Then
        Close FileNum
        GP2V = NL
        frmMain.stbMain.Panels(2) = Read2 & " Dutch " + V
        Exit Sub
    End If
    Read = String(5, " ")
    Get #FileNum, 5675458, Read
    If Read = "Versi" Then
        Close FileNum
        GP2V = Sp
        frmMain.stbMain.Panels(2) = Read2 & " Spanish " + V
        Exit Sub
    End If
    Read = String(7, " ")
    Get #FileNum, 5674990, Read
    If Read = "Version" Then
        Close FileNum
        GP2V = FR
        frmMain.stbMain.Panels(2) = Read2 & " French " + V
        Exit Sub
    End If
    Read = String(8, " ")
    Get #FileNum, 5674331, Read
    If Read = "Versione" Then
        Close FileNum
        GP2V = IT
        frmMain.stbMain.Panels(2) = Read2 & " Italian " + V
        Exit Sub
    End If
    Read = String(21, " ")
    Get #FileNum, 5674544, Read
    If Read = "Deutsche Ausgabe 1.0b" Then
        Close FileNum
        GP2V = TY
        frmMain.stbMain.Panels(2) = Read2 & " German " + V
        Exit Sub
    End If
    Close FileNum
    MsgBox LoadResString(107), vbInformation, TH
    End
End Sub

Public Sub LoadGP2Aid()
    Read2 = oMisc.ReadINI("Misc", "Aids", TempFile)
    Read = Mid(Read2, 1, 7)
    If Read <> "" Then
        For X = 0 To 6
            Read3 = Mid(Read, X + 1, 1)
            If Read3 = "1" Then
                frmMain.R(X).Picture = LoadResPicture(101 + X, 0)
                frmMain.R(X).Tag = "On"
            Else
                frmMain.R(X).Picture = LoadResPicture(108 + X, 0)
                frmMain.R(X).Tag = "Off"
            End If
        Next

        Read = Mid(Read2, 8, 7)
        For X = 0 To 6
            Read3 = Mid(Read, X + 1, 1)
            If Read3 = "1" Then
                frmMain.A(X).Picture = LoadResPicture(101 + X, 0)
                frmMain.A(X).Tag = "On"
            Else
                frmMain.A(X).Picture = LoadResPicture(108 + X, 0)
                frmMain.A(X).Tag = "Off"
            End If
        Next
        Read = Mid(Read2, 15, 7)
        For X = 0 To 6
            Read3 = Mid(Read, X + 1, 1)
            If Read3 = "1" Then
                frmMain.S(X).Picture = LoadResPicture(101 + X, 0)
                frmMain.S(X).Tag = "On"
            Else
                frmMain.S(X).Picture = LoadResPicture(108 + X, 0)
                frmMain.S(X).Tag = "Off"
            End If
        Next
        Read = Mid(Read2, 22, 7)
        For X = 0 To 6
            Read3 = Mid(Read, X + 1, 1)
            If Read3 = "1" Then
                frmMain.P(X).Picture = LoadResPicture(101 + X, 0)
                frmMain.P(X).Tag = "On"
            Else
                frmMain.P(X).Picture = LoadResPicture(108 + X, 0)
                frmMain.P(X).Tag = "Off"
            End If
        Next
        Read = Mid(Read2, 29, 7)
        For X = 0 To 6
            Read3 = Mid(Read, X + 1, 1)
            If Read3 = "1" Then
                frmMain.AC(X).Picture = LoadResPicture(101 + X, 0)
                frmMain.AC(X).Tag = "On"
            Else
                frmMain.AC(X).Picture = LoadResPicture(108 + X, 0)
                frmMain.AC(X).Tag = "Off"
            End If
        Next
    Else
        frmMain.DriveHelpDefault
    End If
End Sub

Public Sub RegFileName()
    oReg.SaveValue HKEY_CLASSES_ROOT, REG_SZ, "Track Handler", "", "GP2 Track Handler File"
    oReg.SaveValue HKEY_CLASSES_ROOT, REG_SZ, ".ths", "", "Track Handler"
    Read = """" & LCase(ProgramDir) & "\" & LCase(App.EXEName) & ".exe" & """" & " %1"
    oReg.SaveValue HKEY_CLASSES_ROOT, REG_SZ, "Track Handler\Shell\Open\Command", "", Read
    Read = LCase(ProgramDir) & "\" & LCase(App.EXEName) & ".exe,1"
    oReg.SaveValue HKEY_CLASSES_ROOT, REG_SZ, "Track Handler\DefaultIcon", "", Read
End Sub

Public Function GetFileName(FilePath As String) As String
'*************************************
'Function Name: GetFileName
'Use: Strip File Name from Path
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-08-26
'*************************************
On Error GoTo ErrHandler
    For X = Len(FilePath) To 1 Step -1
        If Mid(FilePath, X, 1) = "\" Then Exit For
    Next
    GetFileName = Mid(FilePath, X + 1)
Exit Function
ErrHandler:

End Function

Public Sub CreateDir(Path As String)
'*************************************
'Function Name: CreateFolders
'Use: Create Dirs
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-08-26
'*************************************
On Error GoTo ErrHandler
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    For X = 4 To Len(Path)
        X = InStr(X, Path, "\")
        If X <> 0 Then
            MkDir Mid(Path, 1, X)
        Else
            Exit For
        End If
    Next
Exit Sub
ErrHandler:
    Select Case Err.Number
    Case 75
        Resume Next
    Case Else
        MsgBox "Error: " & Err.Number
    End Select
End Sub

Public Sub WriteCheckSum(ByVal sFile As String)
    sFile = oMisc.GetShortName(sFile)
    RetVal = ShellExecute(frmMain.hwnd, "open", ProgramDir & "\gp2utils\thcheck.exe", sFile, vbNullString, 1)
    oMisc.CloseDosPrompt "thcheck"
End Sub

Public Function GetExt(File As String) As String
    GetExt = LCase(Mid(File, Len(File) - 3))
End Function
