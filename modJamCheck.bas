Attribute VB_Name = "modJamCheck"

Public Sub CheckJam(ByVal Track As String)
Dim ItemX

    Var.lLong1 = 0
    FileNum = FreeFile
    frmJamCheck.Show , frmMain
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

    Set ItemX = frmJamCheck.lstJam.ListItems.Add(1, "Top", "Jam Check", 1, 1)
    Set ItemX = frmJamCheck.lstJam.ListItems.Add(2, "Space", "")
    Count1 = 0
    Do Until Len(Read) < 5
        Count1 = Count1 + 1
        Stopp = InStr(1, UCase(Read), UCase(Chr(0)))
        If Stopp = 0 Then
            Exit Do
        End If
        Read2 = Mid(Read, 1, Stopp - 1)
        Read3 = oMisc.File_Exists(GP2Dir & "\" & Read2)
        If Read3 = False Then
            Set ItemX = frmJamCheck.lstJam.ListItems.Add(, "jam" & Count1, Read2, 2, 2)
            ItemX.SubItems(1) = "Not Found!!!"
            Var.lLong1 = Var.lLong1 + 1
        Else
            Set ItemX = frmJamCheck.lstJam.ListItems.Add(, "jam" & Count1, Read2, 1, 1)
            ItemX.SubItems(1) = "Found"
        End If
        Read = Mid(Read, Stopp + 1)
    Loop
    If Var.lLong1 = 0 Then
        frmJamCheck.lstJam.ListItems.Item(1).SmallIcon = 1
        frmJamCheck.lstJam.ListItems.Item(1).SubItems(1) = "All files found!"
    Else
        frmJamCheck.lstJam.ListItems.Item(1).SmallIcon = 2
        frmJamCheck.lstJam.ListItems.Item(1).SubItems(1) = "Faild!!"
    End If
End Sub
