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
Attribute VB_Exposed = False
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
    GetAll = 4
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

Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName)
    Dim R
    R = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Public Function File_Exists(ByVal PathName As String) As Boolean
       File_Exists = IIf(Dir$(PathName) = "", False, True)
End Function

Public Function RecentFile(ByVal RecentType As RecentFile, Optional Path As String, Optional Name As String, Optional OpenNr As String) As Variant
Dim Name1 As String
Dim Name2 As String
Dim Name3 As String
Dim Path1 As String
Dim Path2 As String
Dim Path3 As String
On Error Resume Next
    Name1 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Name1")
    Name2 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Name2")
    Name3 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Name3")
    
    Path1 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Path1")
    Path2 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Path2")
    Path3 = oReg.GetValue(HKEY_CURRENT_USER, "Software\GP2 Track Handler\Files", "Path3")

    If RecentType = OpenRecent Then
        If OpenNr = 2 Then
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name1", Name2
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name2", Name1
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path1", Path2
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path2", Path1
        ElseIf OpenNr = 3 Then
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name1", Name3
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name2", Name1
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name3", Name2

            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path1", Path3
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path2", Path1
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path3", Path2
        End If
    ElseIf RecentType = SaveNew Then
        If Path = Path1 Then
            Exit Function
        ElseIf Path = Path2 Then
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name1", Name2
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name2", Name1
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path1", Path2
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path2", Path1
        Else
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name1", Name
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name2", Name1
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Name3", Name2
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path1", Path
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path2", Path1
            oReg.SaveValue HKEY_CURRENT_USER, REG_SZ, "Software\GP2 Track Handler\Files", "Path3", Path2
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
