Attribute VB_Name = "Random"

Public Sub Rand(UpperBound)
Dim A(15)
Dim B
Dim MainLoop
Dim MaxNumber
Dim ChosenNumber
Dim Counter
    Randomize
    A(0) = Int((UpperBound) * Rnd)
    A(1) = A(0)
    Do Until (A(1) <> A(0))
        A(1) = Int((UpperBound) * Rnd)
    Loop
    A(2) = A(1)
    Do Until (A(2) <> A(0)) And (A(2) <> A(1))
        A(2) = Int((UpperBound) * Rnd)
    Loop
    A(3) = A(2)
    Do Until (A(3) <> A(0)) And (A(3) <> A(1)) And (A(3) <> A(2))
        A(3) = Int((UpperBound) * Rnd)
    Loop
    A(4) = A(3)
    Do Until (A(4) <> A(0)) And (A(4) <> A(1)) And (A(4) <> A(2)) And (A(4) <> A(3))
        A(4) = Int((UpperBound) * Rnd)
    Loop
    A(5) = A(4)
    Do Until (A(5) <> A(0)) And (A(5) <> A(1)) And (A(5) <> A(2)) And (A(5) <> A(3)) And (A(5) <> A(4))
        A(5) = Int((UpperBound) * Rnd)
    Loop
    A(6) = A(5)
    Do Until (A(6) <> A(0)) And (A(6) <> A(1)) And (A(6) <> A(2)) And (A(6) <> A(3)) And (A(6) <> A(4)) And (A(6) <> A(5))
        A(6) = Int((UpperBound) * Rnd)
    Loop
    A(7) = A(6)
    Do Until (A(7) <> A(0)) And (A(7) <> A(1)) And (A(7) <> A(2)) And (A(7) <> A(3)) And (A(7) <> A(4)) And (A(7) <> A(5)) And (A(7) <> A(6))
        A(7) = Int((UpperBound) * Rnd)
    Loop
    A(8) = A(7)
    Do Until (A(8) <> A(0)) And (A(8) <> A(1)) And (A(8) <> A(2)) And (A(8) <> A(3)) And (A(8) <> A(4)) And (A(8) <> A(5)) And (A(8) <> A(6)) And (A(8) <> A(7))
        A(8) = Int((UpperBound) * Rnd)
    Loop
    A(9) = A(8)
    Do Until (A(9) <> A(0)) And (A(9) <> A(1)) And (A(9) <> A(2)) And (A(9) <> A(3)) And (A(9) <> A(4)) And (A(9) <> A(5)) And (A(9) <> A(6)) And (A(9) <> A(7)) And (A(9) <> A(8))
        A(9) = Int((UpperBound) * Rnd)
    Loop
    A(10) = A(9)
    Do Until (A(10) <> A(0)) And (A(10) <> A(1)) And (A(10) <> A(2)) And (A(10) <> A(3)) And (A(10) <> A(4)) And (A(10) <> A(5)) And (A(10) <> A(6)) And (A(10) <> A(7)) And (A(10) <> A(8)) And (A(10) <> A(9))
        A(10) = Int((UpperBound) * Rnd)
    Loop
    A(11) = A(10)
    Do Until (A(11) <> A(0)) And (A(11) <> A(1)) And (A(11) <> A(2)) And (A(11) <> A(3)) And (A(11) <> A(4)) And (A(11) <> A(5)) And (A(11) <> A(6)) And (A(11) <> A(7)) And (A(11) <> A(8)) And (A(11) <> A(9)) And (A(11) <> A(10))
        A(11) = Int((UpperBound) * Rnd)
    Loop
    A(12) = A(11)
    Do Until (A(12) <> A(0)) And (A(12) <> A(1)) And (A(12) <> A(2)) And (A(12) <> A(3)) And (A(12) <> A(4)) And (A(12) <> A(5)) And (A(12) <> A(6)) And (A(12) <> A(7)) And (A(12) <> A(8)) And (A(12) <> A(9)) And (A(12) <> A(10)) And (A(12) <> A(11))
        A(12) = Int((UpperBound) * Rnd)
    Loop
    A(13) = A(12)
    Do Until (A(13) <> A(0)) And (A(13) <> A(1)) And (A(13) <> A(2)) And (A(13) <> A(3)) And (A(13) <> A(4)) And (A(13) <> A(5)) And (A(13) <> A(6)) And (A(13) <> A(7)) And (A(13) <> A(8)) And (A(13) <> A(9)) And (A(13) <> A(10)) And (A(13) <> A(11)) And (A(13) <> A(12))
        A(13) = Int((UpperBound) * Rnd)
    Loop
    A(14) = A(13)
    Do Until (A(14) <> A(0)) And (A(14) <> A(1)) And (A(14) <> A(2)) And (A(14) <> A(3)) And (A(14) <> A(4)) And (A(14) <> A(5)) And (A(14) <> A(6)) And (A(14) <> A(7)) And (A(14) <> A(8)) And (A(14) <> A(9)) And (A(14) <> A(10)) And (A(14) <> A(11)) And (A(14) <> A(12)) And (A(14) <> A(13))
        A(14) = Int((UpperBound) * Rnd)
    Loop
    A(15) = A(14)
    Do Until (A(15) <> A(0)) And (A(15) <> A(1)) And (A(15) <> A(2)) And (A(15) <> A(3)) And (A(15) <> A(4)) And (A(15) <> A(5)) And (A(15) <> A(6)) And (A(15) <> A(7)) And (A(15) <> A(8)) And (A(15) <> A(9)) And (A(15) <> A(10)) And (A(15) <> A(11)) And (A(15) <> A(12)) And (A(15) <> A(13)) And (A(15) <> A(14))
        A(15) = Int((UpperBound) * Rnd)
    Loop
    
    Counter = 0
    Do Until Counter > 15
        MDIForm1.File1.ListIndex = A(Counter)

        Read4 = ProgramDir + "\WorkCopy.lda"
        Read2 = "Track " + Trim(Str(Counter + 1))
        Read = MDIForm1.lblCountry
        GetAdjectiv
        Read3 = oMisc.WriteINI(Read2, "Adjective", Read, Read4)
        Read3 = oMisc.WriteINI(Read2, "Country", MDIForm1.lblCountry, Read4)
        Read3 = oMisc.WriteINI(Read2, "Laps", MDIForm1.lblLaps, Read4)
        Read3 = oMisc.WriteINI(Read2, "Length", MDIForm1.lblLength, Read4)
        Read3 = oMisc.WriteINI(Read2, "Name", MDIForm1.lblName, Read4)
        If Len(MDIForm1.File1.Path) <> 3 Then
            Read = MDIForm1.File1.Path + "\" + MDIForm1.File1.FileName
        Else
            Read = MDIForm1.File1.Path + MDIForm1.File1.FileName
        End If
        Read3 = oMisc.WriteINI(Read2, "TPath", Read, Read4)
        Read3 = oMisc.WriteINI(Read2, "Ware", MDIForm1.lblWare, Read4)
        Read3 = oMisc.WriteINI(Read2, "QTime", MDIForm1.lblQLap, Read4)
        Read3 = oMisc.WriteINI(Read2, "RTime", MDIForm1.lblRLap, Read4)
        Counter = Counter + 1
    Loop
    MakeNewTree
End Sub
