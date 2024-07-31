Attribute VB_Name = "Module1"

Public Function Crypt(Text, Pass)
    On Error Resume Next
    For T = 1 To Len(Pass)
        sana = Asc(Mid(Pass, T, 1))
        X1 = X1 + sana
    Next
    X1 = Int((X1 * 0.1) / 6)
    Pass = X1
    G = 0
    For TT = 1 To Len(Text)
        sana = Asc(Mid(Text, TT, 1))
        G = G + 1
        If G = 6 Then G = 0
        X1 = 0
        If G = 0 Then X1 = sana - (Pass - 2)
        If G = 1 Then X1 = sana + (Pass - 5)
        If G = 2 Then X1 = sana - (Pass - 4)
        If G = 3 Then X1 = sana + (Pass - 2)
        If G = 4 Then X1 = sana - (Pass - 3)
        If G = 5 Then X1 = sana + (Pass - 5)
        X1 = X1 + G
        Crypted = Crypted & Chr(X1)
    Next
    Form1.Text1 = Crypted
End Function

Public Function DeCrypt(Text, Pass)
    On Error Resume Next
    For T = 1 To Len(Pass)
        sana = Asc(Mid(Pass, T, 1))
        X1 = X1 + sana
    Next
    X1 = Int((X1 * 0.1) / 6)
    Pass = X1
    G = 0
    For TT = 1 To Len(Text)
        sana = Asc(Mid(Text, TT, 1))
        G = G + 1
        If G = 6 Then G = 0
        X1 = 0
        If G = 0 Then X1 = sana + (Pass - 2)
        If G = 1 Then X1 = sana - (Pass - 5)
        If G = 2 Then X1 = sana + (Pass - 4)
        If G = 3 Then X1 = sana - (Pass - 2)
        If G = 4 Then X1 = sana + (Pass - 3)
        If G = 5 Then X1 = sana - (Pass - 5)
        X1 = X1 - G
        DeCrypted = DeCrypted & Chr(X1)
    Next
    Form1.Text2 = DeCrypted
End Function
