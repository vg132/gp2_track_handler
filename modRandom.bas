Attribute VB_Name = "modRandom"

Public Sub RandomTracks(UpperBound)
Dim Support As Boolean
Dim vArray As Variant
Dim Counter As Integer
    Randomize
    NewFile
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
        Support = ReadGp2Info(frmMain.lstFile.ListItems.Item(vArray(Count1)).Key)
        DoEvents
        If Support = True Then
            Read2 = "Track " & Counter + 1
            WriteINI Read2, "Adjective", GetAdjectiv(TrackInfo.Country), TempFile
            WriteINI Read2, "Country", TrackInfo.Country, TempFile
            WriteINI Read2, "Laps", TrackInfo.Laps, TempFile
            WriteINI Read2, "Length", TrackInfo.LengthMeters, TempFile
            WriteINI Read2, "Name", TrackInfo.Name, TempFile
            WriteINI Read2, "TPath", frmMain.lstFile.ListItems.Item(vArray(Count1)).Key, TempFile
            WriteINI Read2, "Ware", TrackInfo.Tyre, TempFile
            WriteINI Read2, "QTime", TrackInfo.LapRecordQualify, TempFile
            WriteINI Read2, "RTime", TrackInfo.LapRecord, TempFile
        End If
        If Counter < 15 Then
            Count2 = UBound(vArray)
            vArray(Count1) = vArray(Count2)
            ReDim Preserve vArray(LBound(vArray) To Count2 - 1)
        End If
        Tracks(Counter) = True
    Next
    LoadFile
End Sub
