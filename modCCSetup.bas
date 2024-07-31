Attribute VB_Name = "modCCSetup"
Dim FilePos As Long
Dim X As Long

Public Function GetTrackSetup(ByVal TPath As String) As String
Dim Fuel As Integer
    FileNum = FreeFile
    Open TPath For Binary As FileNum
    Read = String(4000, " ")
    X = FileLen(TPath) - 4000
    Get #FileNum, X, Read
    FilePos = X
    X = InStr(1, LCase(Read), "gamejams\")
    If X > 0 Then
        Read = Mid(Read, 1, X)
    Else
        Exit Function
    End If
    X = InStr(1, LCase(Read), "pdh")
    If X > 0 Then
        Read = Mid(Read, X + 3, 25)
        FilePos = FilePos + X + 2
    Else
        Read2 = String(98, Chr(0))
        X = InStr(1, Read, Read2)
        If X > 0 Then
            Read = Mid(Read, X)
            FilePos = FilePos + X
        ElseIf X = 0 Then
            Exit Function
        End If
        For X = 1 To Len(Read)
            Read2 = Mid(Read, X, 1)
            If Read2 <> Chr(0) Then
                Count1 = X - 1
                X = -1
                Exit For
            End If
        Next
        If X <> -1 Then Exit Function
        Read = Mid(Read, Count1, 25)
        FilePos = FilePos + Count1 - 2
    End If
    Get #FileNum, FilePos + 27, Fuel
    Close FileNum
    GetTrackSetup = Read & Fuel
End Function

Public Sub SaveTrackSetup(ByVal TPath As String, ByVal Data As String, ByVal Fuel As Integer)
Dim Temp As Byte
    FileNum = FreeFile
    Open TPath For Binary As FileNum
    Put #FileNum, FilePos, Data
    Put #FileNum, FilePos + 27, Fuel
    Close FileNum
End Sub
