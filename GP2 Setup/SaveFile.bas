Attribute VB_Name = "modSave"

Public Sub WriteCheckSum(ByVal Path As String)
Dim TempChk As Long
Dim Chk As Integer
    TempChk = CheckSum(Path)
    FileNum = FreeFile
    Open Path For Binary As FileNum
    If TempChk > 32767 Then
        Chk = TempChk - 65536
    Else
        Chk = TempChk
    End If
    Put #FileNum, 83, Chk
    Close FileNum
End Sub

Public Sub WriteFile(ByVal Path As String)
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
        bByte = .txt1.Text
        Put #FileNum, 35, bByte
        bByte = .txt2.Text
        Put #FileNum, 36, bByte
        bByte = .txt3.Text
        Put #FileNum, 37, bByte
        bByte = .txt4.Text
        Put #FileNum, 38, bByte
        bByte = .txt5.Text
        Put #FileNum, 39, bByte
        bByte = .txt6.Text
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
        bByte = .cboRollR.ListIndex + 1
        Put #FileNum, 77, bByte
        bByte = .cboRollF.ListIndex
        Put #FileNum, 79, bByte

    End With
    Close FileNum
    WriteCheckSum Path
End Sub
