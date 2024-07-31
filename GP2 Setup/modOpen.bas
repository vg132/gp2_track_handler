Attribute VB_Name = "modOpen"

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
        .txt1.Text = bByte
        Get #FileNum, 36, bByte
        .txt2.Text = bByte
        Get #FileNum, 37, bByte
        .txt3.Text = bByte
        Get #FileNum, 38, bByte
        .txt4.Text = bByte
        Get #FileNum, 39, bByte
        .txt5.Text = bByte
        Get #FileNum, 40, bByte
        .txt6.Text = bByte

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
        .cboRollR.ListIndex = bByte - 1
        Get #FileNum, 79, bByte
        .cboRollF.ListIndex = bByte
    End With
    Close FileNum
End Sub
