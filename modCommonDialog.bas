Attribute VB_Name = "modCommonDialog"
Option Explicit
Dim TextLen As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

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

Public Function comOpen() As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = frmMain.hwnd
    OpenFile.hInstance = App.hInstance
    sFilter = "GP2 Track Hndler Files (*.ths)" & Chr(0) & "*.ths" & Chr(0) & "All Files (*.*)" & Chr(0) & "*.*" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, " ")
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = ProgramDir
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        comOpen = ""
    Else
        comOpen = Trim(OpenFile.lpstrFile)
        Read = Trim(OpenFile.lpstrFileTitle)
        FileInfo.Name = Trim(Mid(Read, 1, Len(Read) - 1))
        Read = Trim(OpenFile.lpstrFile)
        FileInfo.Path = Trim(Mid(Read, 1, Len(Read) - 1))
    End If
End Function

Public Function comSave() As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = frmMain.hwnd
    OpenFile.hInstance = App.hInstance
    sFilter = "GP2 Track Hndler Files (*.ths)" & Chr(0) & "*.ths" & Chr(0) & "All files (*.*)" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, " ")
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = ProgramDir
    OpenFile.flags = 0
    OpenFile.lpstrDefExt = ".ths"
    lReturn = GetSaveFileName(OpenFile)
    If lReturn = 0 Then
        comSave = ""
    Else
        Read = Trim(OpenFile.lpstrFile)
        Read = Mid(Read, 1, Len(Read) - 1)
        FileInfo.Path = Read
        Read = Trim(OpenFile.lpstrFileTitle)
        Read = Mid(Read, 1, Len(Read) - 1)
        FileInfo.Name = Read
        comSave = FileInfo.Path
    End If
End Function

Public Function comExe() As String
Dim OpenFile As OPENFILENAME
Dim lReturn As Long
Dim sFilter As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = frmMain.hwnd
    OpenFile.hInstance = App.hInstance
    sFilter = "GP2Edit Dos Patch EXE Files (*.exe)" & Chr(0) & "*.exe" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, " ")
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = ProgramDir
    OpenFile.flags = 0
    OpenFile.lpstrDefExt = ".exe"
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        comExe = ""
    Else
        Read = Trim(OpenFile.lpstrFile)
        Read = Mid(Read, 1, Len(Read) - 1)
        comExe = Read
    End If
End Function
