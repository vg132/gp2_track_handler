Attribute VB_Name = "GP2Info"
Public Function GetGP2Info(ByVal Path As String)
Dim Data As String
Dim Start As Integer
Dim Stopp As Integer

    FileNum = FreeFile
    Open Path For Binary As FileNum
    Read = String(4000, " ")
    Get #FileNum, 1, Read
    Close FileNum

    Else
        MsgBox "This file is not supported", vbInformation, TH
    End If
End Function
