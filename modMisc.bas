Attribute VB_Name = "modMisc"
Private Declare Function CheckSum Lib "ThLib.dll" (ByVal FileName As String) As Boolean

Private Gp2Country As String

Public Sub NewTree()
    frmMain.TreeView1.Nodes.Clear
    Dim nodX As Node    ' Create variable.

    Set nodX = frmMain.TreeView1.Nodes.Add(, , "r", "Gp2 Track's", 1, 2)

    X = 11
    Do Until X > 26
        Set nodX = frmMain.TreeView1.Nodes.Add("r", tvwChild, "t" + Trim(Str(X)), "Track " + Trim(Str(X - 10)), 1, 2)
        X = X + 1
    Loop
    nodX.EnsureVisible ' Show all nodes.
End Sub

Public Sub TextSelected()
Dim oText As Object
    Set oText = Screen.ActiveControl
    If TypeName(oText) = "TextBox" Then
        oText.SelStart = 0
        oText.SelLength = Len(oText.Text)
    End If
End Sub

Public Sub GetGp2Version()
Dim V As String
Dim Size As Long
    V = ""
    Size = FileLen(Gp2Dir + "\gp2.exe")
    If Size = 5707881 Then
        'Spanish Version
        Gp2V = Sp
        V = "Spanish"
    ElseIf Size = 5707113 Then
        'German Version
        Gp2V = TY
        V = "German"
    ElseIf Size = 5705385 Then
        'Dutch Version
        Gp2V = NL
        V = "Dutch"
    ElseIf Size = 5707369 Then
        'French Version
        Gp2V = FR
        V = "French"
    ElseIf Size = 5706553 Then
        'Italian Version
        Gp2V = IT
        V = "Italian"
    ElseIf Size = 5702937 Then
        FileNum = FreeFile
        Open Gp2Dir & "\gp2.exe" For Binary As FileNum
        Read = String(23, " ")
        Get #FileNum, 5671742, Read
        Close FileNum
        If Read = "US English Version 1.0b" Then
            Gp2V = US
            V = "American"
        Else
            Gp2V = UK
            V = "UK English"
        End If
    Else
        MsgBox "Track Handler was not able to check your Gp2 Version." & vbLf & _
        "You may have to reinstall Gp2 to make this program work.", vbInformation, TH
    End If
    If V <> "" Then
        frmMain.StatusBar1.Panels(2).Text = "Gp2 Version: " & V & " Version 1.0b"
    End If
End Sub

Public Sub LoadGp2Aid(ByVal sAids As String)
    Read2 = ReadINI("Misc", "Aids", TempFile)
    If Read2 = "" Then Read2 = "11111110111111011111101001110100001"
    Read = Mid(sAids, 1, 7)
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
        Read = Mid(sAids, 8, 7)
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
        Read = Mid(sAids, 15, 7)
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
        Read = Mid(sAids, 22, 7)
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
        Read = Mid(sAids, 29, 7)
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
    End If
End Sub

Public Sub RegFileName()
    oReg.SaveValue HKEY_CLASSES_ROOT, REG_SZ, "Track Handler", "", "Gp2 Track Handler File"
    oReg.SaveValue HKEY_CLASSES_ROOT, REG_SZ, ".ths", "", "Track Handler"
    Read = """" & LCase(ProgramDir) & "\" & LCase(App.EXEName) & ".exe" & """" & " %1"
    oReg.SaveValue HKEY_CLASSES_ROOT, REG_SZ, "Track Handler\Shell\Open\Command", "", Read
    Read = LCase(ProgramDir) & "\" & LCase(App.EXEName) & ".exe,1"
    oReg.SaveValue HKEY_CLASSES_ROOT, REG_SZ, "Track Handler\DefaultIcon", "", Read
End Sub

Public Sub CreateDir(Path As String)
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
        MsgBox "Error Nr: " & Str(Err.Number) & vbLf & _
            "Error Desctiption: " & Err.Description & vbLf & _
            "Error Source: CreateDir()", vbCritical, TH & " - Error"
    End Select
End Sub

Public Sub WriteCheckSum(ByVal sFile As String)
Dim RetVal As Boolean
    RetVal = CheckSum(sFile)
    If RetVal = False Then
        MsgBox "Error when writing checksum to file " & sFile & ".", vbCritical, TH
    End If
End Sub

Public Sub INetLink(URL As String, hWnd As Long)
Dim RetVal
    RetVal = ShellExecute(hWnd, "open", URL, vbNullString, vbNullString, 1)
End Sub
