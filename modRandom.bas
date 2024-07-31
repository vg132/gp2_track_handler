Attribute VB_Name = "modRandom"
Public Sub RandomTracks(UpperBound)
Dim Support As Boolean
Dim vArray As Variant
Dim Counter As Integer

    Randomize
    ReDim vArray(0 To UpperBound - 1)
    For X = 0 To UpperBound - 1
        vArray(X) = X + 1
    Next
    For X = UpperBound - 1 To 16 Step -1
        Count1 = Int((UBound(vArray)) * Rnd)
        Count2 = UBound(vArray)
        vArray(Count1) = vArray(Count2)
        ReDim Preserve vArray(LBound(vArray) To Count2 - 1)
    Next
    
    'NewFile
    For Counter = 0 To 15
        Count1 = Int((UBound(vArray)) * Rnd)
        frmMain.lstFile.ListItems(vArray(Count1)).Selected = True
        If Counter < 15 Then
            Count2 = UBound(vArray)
            vArray(Count1) = vArray(Count2)
            ReDim Preserve vArray(LBound(vArray) To Count2 - 1)
        End If
        Support = ReadGP2Info(frmMain.lstFile.SelectedItem.Key)
        If Support = True Then
            Read2 = "Track " + Trim(Str(Counter + 1))
            Read = GetAdjectiv(frmMain.lblCountry)
            oMisc.WriteINI Read2, "Adjective", Read, TempFile
            oMisc.WriteINI Read2, "Country", frmMain.lblCountry, TempFile
            oMisc.WriteINI Read2, "Laps", frmMain.lblLaps, TempFile
            oMisc.WriteINI Read2, "Length", frmMain.lblLen, TempFile
            oMisc.WriteINI Read2, "Name", frmMain.lblTrackName, TempFile
            oMisc.WriteINI Read2, "TPath", frmMain.lstFile.SelectedItem.Key, TempFile
            oMisc.WriteINI Read2, "Ware", frmMain.lblWare, TempFile
            oMisc.WriteINI Read2, "QTime", frmMain.lblQual, TempFile
            oMisc.WriteINI Read2, "RTime", frmMain.lblRace, TempFile
            oMisc.WriteINI Read2, "QDriver", "", TempFile
            oMisc.WriteINI Read2, "RDriver", "", TempFile
            oMisc.WriteINI Read2, "QTeam", "", TempFile
            oMisc.WriteINI Read2, "RTeam", "", TempFile
            oMisc.WriteINI Read2, "QDate", "", TempFile
            oMisc.WriteINI Read2, "RDate", "", TempFile
        End If
    Next
    LoadFile
End Sub
