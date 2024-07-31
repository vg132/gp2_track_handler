Attribute VB_Name = "modAPI"
Option Explicit

Public Adj(15) As String
Public TrackName(15) As String
Public Country(15) As String
Public TempFile As String
Public ProgramDir As String
Public GP2Dir As String
Public Read As String 'Public temp string
Public Read2 As String
Public Read3 As String
Public Read4 As String
Public CountExport As String
Public GP2V As String
Public GP2NameFile As String
Public dbFile As String

Public F1SaveFileNum As Integer
Public FileNum As Integer
Public FileNum2 As Integer
Public GP2FileNum As Integer
Public TreeNr As Integer
Public Responce As Integer
Public Log As Integer

Public X As Long 'Public temp long
Public Count1 As Long
Public Count2 As Long

Public TempDouble As Double

Public Const TH = "GP2 Track Handler"
Public oMisc As New Misc
Public oData As New GP2Info
Public oReg As New oReg
Public oDB As New clsDB
Public FileInfo As File
Public TheDate As Date
Public tImp As ImportType
Public tExp As ExportType

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&
Private Const HKEY_CLASSES_ROOT = &H80000000
Public Const MAX_PATH = 260&
Private Const REG_SZ = 1

Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&

Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)

Public Const LVS_EX_FULLROWSELECT As Long = &H20
Public Const LVS_EX_GRIDLINES As Long = &H1
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVIS_STATEIMAGEMASK As Long = &HF000

Public Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   State        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Public Type File
    Path As String * 257
    Name As String * 257
    Saved As Boolean
    Changes As Boolean
    Import As Boolean
End Type

Public Type ImportType
    iInt As Integer
    bByte As Byte
    lLong As Long
    Year As String * 4
End Type

Public Type ExportType
    iInt As Integer
    bByte As Byte
    lLong As Long
    Year As String * 4
End Type

Public Enum ImpExpTime
    Qual = 0
    Race = 1
End Enum

Public Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
