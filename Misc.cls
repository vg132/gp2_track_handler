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
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)

Private Const WM_CLOSE = &H10

Private lpIDList As Long
Private sBuffer As String
Private szTitle As String
Private tBrowseInfo As BrowseInfo

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260&

Public Enum GP2Ver
    US = 1
    UK = 2
    NL = 3
    IT = 4
    Sp = 5
    FR = 6
    TY = 7
End Enum

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Public Enum RecentFile
    OpenRecent = 1
    SaveNew = 2
    Check = 3
    GetAll = 4
End Enum

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Function ShowOpen(ByVal Filter As String, ByVal hwnd As Long, ByVal InitDir As String, Optional Title As String, Optional DefName As String) As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = hwnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = ConvertFilter(Filter)
    OpenFile.nFilterIndex = 1
    If DefName = "" Then
        OpenFile.lpstrFile = Space$(1024) & Chr(0)
    Else
        OpenFile.lpstrFile = DefName & Space$(1024) & Chr(0)
    End If
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile)
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile & Chr(0)
    OpenFile.lpstrInitialDir = InitDir & Chr(0)
    OpenFile.flags = 0
    If Title <> "" Then OpenFile.lpstrTitle = Title
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        ShowOpen = ""
    Else
        ShowOpen = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, vbNullChar) - 1)
    End If
End Function

Public Function ShowSave(ByVal Filter As String, ByVal DefExt As String, ByVal hwnd As Long, ByVal InitDir As String, Optional Title As String, Optional DefName As String) As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = hwnd
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = ConvertFilter(Filter)
    OpenFile.nFilterIndex = 1
    If DefName = "" Then
        OpenFile.lpstrFile = Space$(1024) & Chr(0)
    Else
        OpenFile.lpstrFile = DefName & Space$(1024) & Chr(0)
    End If
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile)
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile & Chr(0)
    OpenFile.lpstrInitialDir = InitDir & Chr(0)
    OpenFile.flags = 0
    OpenFile.lpstrDefExt = DefExt
    If Title <> "" Then OpenFile.lpstrTitle = Title
    lReturn = GetSaveFileName(OpenFile)
    If lReturn = 0 Then
        ShowSave = ""
    Else
        ShowSave = Left(OpenFile.lpstrFile, InStr(OpenFile.lpstrFile, vbNullChar) - 1)
    End If
End Function

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
    DoEvents
End Function

Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName)
    Dim R
    R = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
    DoEvents
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

Private Function ConvertFilter(ByVal Filter As String) As String
Dim X As Long
Dim Read As String
    Read = ""
    For X = 1 To Len(Filter)
        If Mid(Filter, X, 1) = "|" Then
            Read = Read & Chr(0)
        Else
            Read = Read & Mid(Filter, X, 1)
        End If
    Next
    ConvertFilter = Read & Chr(0)
End Function

Public Function BrowseFolders(ByVal Title As String, ByVal hwnd As Long) As String
    szTitle = Title
    With tBrowseInfo
        .hwndOwner = hwnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        BrowseFolders = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    End If
End Function

Public Sub INetLink(URL As String, hwnd As Long)
Dim RetVal
    RetVal = ShellExecute(hwnd, "open", URL, vbNullString, vbNullString, 1)
End Sub

Public Function CloseDosPrompt(ByVal Title As String) As Integer
'*************************************
'Function Name: CloseDosPrompt
'Use: Close a Dos Prompt Window When the program is ready
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-09-04
'*************************************
Dim RetVal As Long
Dim Handle As Long
Dim X As Integer
On Error GoTo ErrHandler
    Sleep (500)
    Handle = GetForegroundWindow()
    X = 1
    Do Until X > 20
        RetVal = FindWindow(vbNullString, Title)
        If RetVal = 0 Then
            RetVal = SendMessage(Handle, WM_CLOSE, &O0, &O0)
            Exit Do
        End If
        Sleep (500)
        X = X + 1
    Loop
Exit Function
ErrHandler:
    MsgBox "Error"
End Function
