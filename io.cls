VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Enum GP2Ver
    US = 1
    UK = 2
    NL = 3
    IT = 4
    Sp = 5
    FR = 6
    TY = 7
End Enum

Public Enum RecentFile
    OpenRecent = 1
    SaveNew = 2
    Check = 3
End Enum


Public Function GetShortName(ByVal sLongFileName As String) As String
Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    GetShortName = Left(sShortPathName, lRetVal)
End Function

Public Function ReadINI(Section, KeyName, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim R
    R = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Public Function File_Exists(ByVal PathName As String) As Boolean
       File_Exists = IIf(Dir$(PathName) = "", False, True)
End Function

Public Function SaveFile(ByVal FilePath As String, ProgramDir As String, ByVal FileType As Integer) As Boolean
Dim OldFile As String
Dim FileNum As Integer
Dim NewFile As String
    If FileType = 1 Then
        FileNum = FreeFile
        Open ProgramDir & "\mall.lda" For Binary As FileNum
        OldFile = String(FileLen(ProgramDir & "\mall.lda"), " ")
        Get #FileNum, 1, OldFile
        Close FileNum
        FileNum = FreeFile
        Open ProgramDir & "\WorkCopy.lda" For Binary As FileNum
        NewFile = String(FileLen(ProgramDir & "\WorkCopy.lda"), 1)
        Get #FileNum, 1, NewFile
    End If
    If NewFile <> OldFile Then
        SaveFile = True
    Else
        SaveFile = False
    End If
End Function

Public Function RecentFile(ByVal OpenNr As Integer, Path As String, Name As String, ProgramName As String, RecentType As RecentFile) As Variant
Dim Name1 As String
Dim Name2 As String
Dim Name3 As String
Dim Path1 As String
Dim Path2 As String
Dim Path3 As String
On Error Resume Next
    Name1 = GetSetting(ProgramName, "RecentFile", "Name1")
    Name2 = GetSetting(ProgramName, "RecentFile", "Name2")
    Name3 = GetSetting(ProgramName, "RecentFile", "Name3")

    Path1 = GetSetting(ProgramName, "RecentFile", "Path1")
    Path2 = GetSetting(ProgramName, "RecentFile", "Path2")
    Path3 = GetSetting(ProgramName, "RecentFile", "Path3")

    If RecentType = OpenRecent Then
        If OpenNr = 2 Then
            SaveSetting ProgramName, "RecentFile", "Name1", Name2
            SaveSetting ProgramName, "RecentFile", "Name2", Name1
            SaveSetting ProgramName, "RecentFile", "Path1", Path2
            SaveSetting ProgramName, "RecentFile", "Path2", Path1
        ElseIf OpenNr = 3 Then
            SaveSetting ProgramName, "RecentFile", "Name1", Name3
            SaveSetting ProgramName, "RecentFile", "Name2", Name1
            SaveSetting ProgramName, "RecentFile", "Name3", Name2

            SaveSetting ProgramName, "RecentFile", "Path1", Path3
            SaveSetting ProgramName, "RecentFile", "Path2", Path1
            SaveSetting ProgramName, "RecentFile", "Path3", Path2
        End If
    ElseIf RecentType = SaveNew Then
        If Path = Path1 Then
            Exit Function
        ElseIf Path = Path2 Then
            SaveSetting ProgramName, "RecentFile", "Path1", Path2
            SaveSetting ProgramName, "RecentFile", "Path2", Path1
            SaveSetting ProgramName, "RecentFile", "Name1", Name1
            SaveSetting ProgramName, "RecentFile", "Name2", Name2
        Else
            SaveSetting ProgramName, "RecentFile", "Path1", Path
            SaveSetting ProgramName, "RecentFile", "Path2", Path1
            SaveSetting ProgramName, "RecentFile", "Path3", Path2
            SaveSetting ProgramName, "RecentFile", "Name1", Name
            SaveSetting ProgramName, "RecentFile", "Name2", Name1
            SaveSetting ProgramName, "RecentFile", "Name3", Name2
        End If
    ElseIf RecentType = Check Then
        If Path1 = "" Then
            RecentFile = 0
        ElseIf Path2 = "" Then
            RecentFile = 1
        ElseIf Path3 = "" Then
            RecentFile = 2
        Else
            RecentFile = 3
        End If
    End If
End Function

Public Function DeCrypt(ByVal Text, Pass) As String
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
End Function
