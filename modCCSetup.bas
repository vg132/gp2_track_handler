Attribute VB_Name = "modCCSetup"
Dim PitPos As Long
Dim CCSetup As Long

Public Sub GetCCSetup(ByVal TPath As String)
    FileNum = FreeFile
    Open TPath For Binary As FileNum
    Get #FileNum, 4117, tVar.lLong
    CCSetup = tVar.lLong + 4128
    Read = String(30, " ")
    Get #FileNum, CCSetup, Read
    Close FileNum
    FileNum = FreeFile
    Open ProgramDir & "\File\CCSetup.tmp" For Binary As FileNum
    Put #FileNum, 1, Read
    Close FileNum
End Sub

Public Sub SaveCCCarSetup(ByVal TPath As String)
    FileNum = FreeFile
    Open ProgramDir & "\File\CCSetup.tmp" For Binary As FileNum
    Read = String(30, " ")
    Get #FileNum, 1, Read
    Close FileNum

    FileNum = FreeFile
    Open TPath For Binary As FileNum
    Put #FileNum, CCSetup, Read
    Close FileNum
End Sub

Public Sub GetPitStop(ByVal TPath As String)
'*************************************
'Function Name: GetPitStop
'Use: Get the PitStop Stratergy
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-05
'*************************************
    FileNum = FreeFile
    Open TPath For Binary As FileNum
    Read = String(3000, Chr(0))
    Get #FileNum, FileLen(TPath) - 3015, Read
    Close FileNum
    Read2 = String(2, Chr(0))
    For X = 3000 To 1 Step -1
        If Mid(Read, X, 2) = Read2 Then
            Read = Mid(Read, X - 324, 52)
            FileNum = FreeFile
            Open ProgramDir & "\File\PitStop.tmp" For Binary As FileNum
            Put #FileNum, 1, Read
            Close FileNum
            PitPos = FileLen(TPath) - (3341 - X) + 1
            Exit For
        End If
    Next
End Sub

Public Sub SaveCCPitStop(ByVal TPath As String)
'*************************************
'Function Name: SavePitStop
'Use: Save PitStop Stratergy
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-05
'*************************************
    FileNum = FreeFile
    Open ProgramDir & "\File\PitStop.tmp" For Binary As FileNum
    Read = String(52, " ")
    Get #FileNum, 1, Read
    Close FileNum

    FileNum = FreeFile
    Open TPath For Binary As FileNum
    Put #FileNum, PitPos, Read
    Close FileNum
End Sub
