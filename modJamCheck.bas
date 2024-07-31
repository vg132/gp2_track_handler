Attribute VB_Name = "modJamCheck"

Public Sub CheckJam(ByVal Track As String)
Dim ItemX
Dim Fail As Long
    Fail = 0
    FileNum = FreeFile
    frmJamCheck.Show , frmMain
    Open Track For Binary As FileNum
    Read = String(2000, " ")
    X = FileLen(Track) - 2000
    Get #FileNum, X, Read
    Close FileNum
    Start = InStr(1, UCase(Read), UCase("gamejams\"))
    Read = Mid(Read, Start, Len(Read) - Start)
    Set ItemX = frmJamCheck.lstJam.ListItems.Add(1, "Top", "Jam Check", 1, 1)
    Set ItemX = frmJamCheck.lstJam.ListItems.Add(2, "Space", "")
    Count1 = 0
    Do Until Len(Read) < 5
        Count1 = Count1 + 1
        Stopp = InStr(1, UCase(Read), UCase(".jam"))
        If Stopp = 0 Then
            Exit Do
        End If
        Stopp = Stopp + 3
        Read2 = Mid(Read, 1, Stopp)
        Read3 = oMisc.File_Exists(GP2Dir & "\" & Read2)
        If Read3 = False Then
            Set ItemX = frmJamCheck.lstJam.ListItems.Add(, "jam" & Count1, Read2, 2, 2)
            ItemX.SubItems(1) = "Not Found!!!"
            Fail = Fail + 1
        Else
            Set ItemX = frmJamCheck.lstJam.ListItems.Add(, "jam" & Count1, Read2, 1, 1)
            ItemX.SubItems(1) = "Found"
        End If
        Read = Mid(Read, Stopp + 2)
    Loop
    If Fail = 0 Then
        frmJamCheck.lstJam.ListItems.Item(1).SmallIcon = 1
        frmJamCheck.lstJam.ListItems.Item(1).SubItems(1) = "All files found!"
    Else
        frmJamCheck.lstJam.ListItems.Item(1).SmallIcon = 2
        frmJamCheck.lstJam.ListItems.Item(1).SubItems(1) = "Faild!!"
    End If
End Sub
