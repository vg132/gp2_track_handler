Attribute VB_Name = "modGP2Setup"

Public Sub ExportRaceSetup()
    Read = ReadINI("Track " & Exp.TrackNr + 1, "TPath", TempFile)
    FileNum = FreeFile
    Open Read For Binary As FileNum
    Get #FileNum, 3997, tVar.iInt
    If tVar.iInt = 12345 Then
        Get #FileNum, 3999, tVar.iInt
        If tVar.iInt = 1 Then
            Read = String(48, Chr(0))
            Get #FileNum, 4049, Read
            Read2 = Mid(Read, 1, 8)
            Put #Exp.F1FileNum, 11251 + (Exp.TrackNr * 96), Read2
            Read2 = Mid(Read, 10, 1)
            Put #Exp.F1FileNum, 11251 + (Exp.TrackNr * 96) + 9, Read2
            Read2 = Mid(Read, 17)
            Put #Exp.F1FileNum, 11251 + (Exp.TrackNr * 96) + 16, Read2
        Else
            ResetSetup Race
        End If
    Else
        ResetSetup Race
    End If
    Close FileNum
End Sub

Public Sub ExportQualSetup()
    Read = ReadINI("Track " & Exp.TrackNr + 1, "TPath", TempFile)
    FileNum = FreeFile
    Open Read For Binary As FileNum
    Get #FileNum, 3997, tVar.iInt
    If tVar.iInt = 12345 Then
        Get #FileNum, 3999, tVar.iInt
        If tVar.iInt = 1 Then
            Read = String(48, Chr(0))
            Get #FileNum, 4001, Read
            Read2 = Mid(Read, 1, 8)
            Put #Exp.F1FileNum, 11299 + (Exp.TrackNr * 96), Read2
            Read2 = Mid(Read, 10, 1)
            Put #Exp.F1FileNum, 11299 + (Exp.TrackNr * 96) + 9, Read2
            Read2 = Mid(Read, 17)
            Put #Exp.F1FileNum, 11299 + (Exp.TrackNr * 96) + 16, Read2
        Else
            ResetSetup Qual
        End If
    Else
        ResetSetup Qual
    End If
    Close FileNum
End Sub

Public Sub AddSetup(ByVal TrackFile As String, ByVal SetupFile As String, QOrR As QR)
    FileNum = FreeFile
    Open SetupFile For Binary As FileNum
    Read = String(48, Chr(0))
    Get #FileNum, 33, Read
    Close FileNum
    FileNum = FreeFile
    Open TrackFile For Binary As FileNum
    If QOrR = Qual Then
        Put #FileNum, 4001, Read
    Else
        Put #FileNum, 4049, Read
    End If
    tVar.iInt = 1
    Put #FileNum, 3999, tVar.iInt
    tVar.iInt = 12345
    Put #FileNum, 3997, tVar.iInt
    Close FileNum
End Sub

Public Sub SaveSetupFile(ByVal Path As String)
Dim bByte As Byte
    FileNum = FreeFile
    Open Path For Binary As FileNum
    With frmSetup
        'Wing
        bByte = .txtFWing
        Put #FileNum, 33, bByte
        bByte = .txtRWing.Text
        Put #FileNum, 34, bByte

        'Gear
        bByte = .txtGear(0).Text
        Put #FileNum, 35, bByte
        bByte = .txtGear(1).Text
        Put #FileNum, 36, bByte
        bByte = .txtGear(2).Text
        Put #FileNum, 37, bByte
        bByte = .txtGear(3).Text
        Put #FileNum, 38, bByte
        bByte = .txtGear(4).Text
        Put #FileNum, 39, bByte
        bByte = .txtGear(5).Text
        Put #FileNum, 40, bByte

        'Brake
        bByte = .hscBrake.Value
        Put #FileNum, 42, bByte

        'Packers
        bByte = .txtPacR(0).Text
        Put #FileNum, 49, bByte
        bByte = .txtPacR(1).Text
        Put #FileNum, 50, bByte
        bByte = .txtPacF(0).Text
        Put #FileNum, 51, bByte
        bByte = .txtPacF(1).Text
        Put #FileNum, 52, bByte
    
        'Fast Dumper
        bByte = .txtFastBumpR(0).Text
        Put #FileNum, 53, bByte
        bByte = .txtFastBumpR(1).Text
        Put #FileNum, 54, bByte
        bByte = .txtFastBumpF(0).Text
        Put #FileNum, 55, bByte
        bByte = .txtFastBumpF(1).Text
        Put #FileNum, 56, bByte

        'Slow Dumper
        bByte = .txtSlowBumpR(0).Text
        Put #FileNum, 61, bByte
        bByte = .txtSlowBumpR(1).Text
        Put #FileNum, 62, bByte
        bByte = .txtSlowBumpF(0).Text
        Put #FileNum, 63, bByte
        bByte = .txtSlowBumpF(1).Text
        Put #FileNum, 64, bByte
        
        'Fast Rebound
        bByte = .txtFastReboundR(0).Text
        Put #FileNum, 57, bByte
        bByte = .txtFastReboundR(1).Text
        Put #FileNum, 58, bByte
        bByte = .txtFastReboundF(0).Text
        Put #FileNum, 59, bByte
        bByte = .txtFastReboundF(1).Text
        Put #FileNum, 60, bByte

        'Slow Rebound
        bByte = .txtSlowReboundR(0).Text
        Put #FileNum, 65, bByte
        bByte = .txtSlowReboundR(1).Text
        Put #FileNum, 66, bByte
        bByte = .txtSlowReboundF(0).Text
        Put #FileNum, 67, bByte
        bByte = .txtSlowReboundF(1).Text
        Put #FileNum, 68, bByte
        
        'Spring
        bByte = .cboSpringR(0).Text / 10
        Put #FileNum, 69, bByte
        bByte = .cboSpringR(1).Text / 10
        Put #FileNum, 70, bByte
        bByte = .cboSpringF(0).Text / 10
        Put #FileNum, 71, bByte
        bByte = .cboSpringF(1).Text / 10
        Put #FileNum, 72, bByte
    
        'Ride Height
        bByte = .hscHeightR(0)
        Put #FileNum, 73, bByte
        bByte = .hscHeightR(1)
        Put #FileNum, 74, bByte
        bByte = .hscHeightF(0)
        Put #FileNum, 75, bByte
        bByte = .hscHeightF(1)
        Put #FileNum, 76, bByte

        'Anti Roll Bar
        tVar.iInt = .cboRollR.ListIndex
        Put #FileNum, 77, tVar.iInt
        tVar.iInt = .cboRollF.ListIndex
        Put #FileNum, 79, tVar.iInt
    End With
    Close FileNum
End Sub

Public Sub OpenSetup(ByVal Path As String)
Dim FileNum As Integer
Dim bByte As Byte
    With frmSetup
        FileNum = FreeFile
        Open Path For Binary As FileNum
        'Wing
        Get #FileNum, 33, bByte
        .txtFWing = bByte
        Get #FileNum, 34, bByte
        .txtRWing.Text = bByte

        'Gear
        Get #FileNum, 35, bByte
        .txtGear(0).Text = bByte
        Get #FileNum, 36, bByte
        .txtGear(1).Text = bByte
        Get #FileNum, 37, bByte
        .txtGear(2).Text = bByte
        Get #FileNum, 38, bByte
        .txtGear(3).Text = bByte
        Get #FileNum, 39, bByte
        .txtGear(4).Text = bByte
        Get #FileNum, 40, bByte
        .txtGear(5).Text = bByte

        'Brake
        Get #FileNum, 42, bByte
        .hscBrake.Value = bByte

        'Packers
        Get #FileNum, 49, bByte
        .txtPacR(0).Text = bByte
        Get #FileNum, 50, bByte
        .txtPacR(1).Text = bByte
        Get #FileNum, 51, bByte
        .txtPacF(0).Text = bByte
        Get #FileNum, 52, bByte
        .txtPacF(1).Text = bByte
    
        'Fast Dumper
        Get #FileNum, 53, bByte
        .txtFastBumpR(0).Text = bByte
        Get #FileNum, 54, bByte
        .txtFastBumpR(1).Text = bByte
        Get #FileNum, 55, bByte
        .txtFastBumpF(0).Text = bByte
        Get #FileNum, 56, bByte
        .txtFastBumpF(1).Text = bByte

        'Slow Dumper
        Get #FileNum, 61, bByte
        .txtSlowBumpR(0).Text = bByte
        Get #FileNum, 62, bByte
        .txtSlowBumpR(1).Text = bByte
        Get #FileNum, 63, bByte
        .txtSlowBumpF(0).Text = bByte
        Get #FileNum, 64, bByte
        .txtSlowBumpF(1).Text = bByte
        
        'Fast Rebound
        Get #FileNum, 57, bByte
        .txtFastReboundR(0).Text = bByte
        Get #FileNum, 58, bByte
        .txtFastReboundR(1).Text = bByte
        Get #FileNum, 59, bByte
        .txtFastReboundF(0).Text = bByte
        Get #FileNum, 60, bByte
        .txtFastReboundF(1).Text = bByte

        'Slow Rebound
        Get #FileNum, 65, bByte
        .txtSlowReboundR(0).Text = bByte
        Get #FileNum, 66, bByte
        .txtSlowReboundR(1).Text = bByte
        Get #FileNum, 67, bByte
        .txtSlowReboundF(0).Text = bByte
        Get #FileNum, 68, bByte
        .txtSlowReboundF(1).Text = bByte
        
        'Spring
        Get #FileNum, 69, bByte
        .cboSpringR(0).Text = bByte * 10
        Get #FileNum, 70, bByte
        .cboSpringR(1).Text = bByte * 10
        Get #FileNum, 71, bByte
        .cboSpringF(0).Text = bByte * 10
        Get #FileNum, 72, bByte
        .cboSpringF(1).Text = bByte * 10
    
        'Ride Height
        Get #FileNum, 73, bByte
        .hscHeightR(0) = bByte
        Get #FileNum, 74, bByte
        .hscHeightR(1) = bByte
        Get #FileNum, 75, bByte
        .hscHeightF(0) = bByte
        Get #FileNum, 76, bByte
        .hscHeightF(1) = bByte

        'Anti Roll Bar
        Get #FileNum, 77, bByte
        .cboRollR.ListIndex = bByte
        Get #FileNum, 79, bByte
        .cboRollF.ListIndex = bByte
    End With
    Close FileNum
End Sub

Public Sub DeteteSetup(File As String)
'*************************************
'Function Name: DeteteSetup
'Use: Delete Setup from a track
'Remarks: Add 1000 chr(0) bytes to the file
'History:
'Programmer: Viktor Gars
'Date: 1999-08-28
'*************************************
    FileNum = FreeFile
    Open File For Binary As FileNum
    Read = String(100, Chr(0))
    Put #FileNum, 3997, Read
    Close FileNum
End Sub

Public Sub ResetSetup(ByVal sQT As QR)
Dim vData As Variant
Dim X As Integer
Dim sData As String
    sData = ""
    vData = LoadResData(101, "Setup")
    For X = 0 To 8
        sData = sData & Chr(vData(X))
    Next
    If sQT = Race Then
        Put #Exp.F1FileNum, 11251 + (Exp.TrackNr * 96), sData
    Else
        Put #Exp.F1FileNum, 11299 + (Exp.TrackNr * 96), sData
    End If
    sData = Chr(vData(9))
    If sQT = Race Then
        Put #Exp.F1FileNum, 11251 + (Exp.TrackNr * 96) + 9, sData
    Else
        Put #Exp.F1FileNum, 11299 + (Exp.TrackNr * 96) + 9, sData
    End If
    sData = ""
    For X = 16 To UBound(vData)
        sData = sData & Chr(vData(X))
    Next
    If sQT = Race Then
        Put #Exp.F1FileNum, 11251 + (Exp.TrackNr * 96) + 16, sData
    Else
        Put #Exp.F1FileNum, 11299 + (Exp.TrackNr * 96) + 16, sData
    End If
End Sub

Public Function CheckSetup(ByVal FileName As String) As Boolean
    FileNum = FreeFile
    Open FileName For Binary As FileNum
    Get #FileNum, 3997, tVar.iInt
    If tVar.iInt = 12345 Then
        Get #FileNum, 3999, tVar.iInt
        If tVar.iInt = 1 Then
            CheckSetup = True
        Else
            CheckSetup = False
        End If
    Else
        CheckSetup = False
    End If
    Close FileNum
End Function
