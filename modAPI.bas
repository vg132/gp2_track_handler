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
Public Exp As ExpVar
Public Tracks(15) As Boolean

Public Var As VariableType

Public FileNum As Integer
Public FileNum2 As Integer
Public TreeNr As Integer

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

Public Const MAX_PATH = 260&

Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&

Public Const LVM_FIRST As Long = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)

Public Const LVS_EX_FULLROWSELECT As Long = &H20

Private Const WM_USER = &H400
Private Const TB_SETSTYLE = WM_USER + 56
Private Const TB_GETSTYLE = WM_USER + 57
Private Const TBSTYLE_FLAT = &H800
Private Const TBSTYLE_LIST = &H1000

Public Const OF_READ = &H0

'Info om den öppnade filen
Public Type File
    Path As String
    Name As String
    Saved As Boolean
    Changes As Boolean
    Import As Boolean
End Type

Public Enum RecEnum
    F1gstate = 0
    RecFile = 1
End Enum

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

'Variablar
Public Type VariableType
    iInt1 As Integer
    iInt2 As Integer
    bByte1 As Byte
    bByte2 As Byte
    lLong1 As Long
    lLong2 As Long
    sString1 As String
    sString2 As String
    dDouble1 As Double
    dDouble2 As Double
End Type

Public Type ExpVar
    TrackNr As Integer
    GP2FileNum As Integer
    F1FileNum As Integer
End Type

Public Enum ImpExpTime
    Qual = 0
    Race = 1
End Enum

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type SYSTEMTIME
  wYear          As Integer
  wMonth         As Integer
  wDayOfWeek     As Integer
  wDay           As Integer
  wHour          As Integer
  wMinute        As Integer
  wSecond        As Integer
  wMilliseconds  As Long
End Type

Type OFSTRUCT
   cBytes      As Byte
   fFixedDisk  As Byte
   nErrCode    As Integer
   Reserved1   As Integer
   Reserved2   As Integer
   szPathName(MAX_PATH) As Byte
End Type

Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Sub FlatToolbar(tlb As Toolbar)
Dim lngStyle As Long
Dim lngResult As Long
Dim lngHWND As Long
    lngHWND = FindWindowEx(tlb.hwnd, 0&, "ToolbarWindow32", vbNullString)
    lngStyle = SendMessageLong(lngHWND, TB_GETSTYLE, 0&, 0&)
    lngStyle = lngStyle Or TBSTYLE_FLAT
    lngResult = SendMessageLong(lngHWND, TB_SETSTYLE, 0, lngStyle)
    tlb.Refresh
End Sub

Public Function GetFileDateString(CT As FILETIME) As String
'*************************************
'Function Name: GetFileDateString
'Use: Convert to normal date
'Remarks:
'History:
'Programmer: Viktor Gars
'Date: 1999-08-17
'*************************************
Dim ST As SYSTEMTIME
Dim R As Long
Dim ds As Single
  If FileTimeToSystemTime(CT, ST) Then
    ds = DateSerial(ST.wYear, ST.wMonth, ST.wDay)
    'GetFileDateString = Format$(ds, "DDDD MMMM D, YYYY")
    GetFileDateString = Format$(ds, "YYYY-MM-DD")
  End If
End Function
