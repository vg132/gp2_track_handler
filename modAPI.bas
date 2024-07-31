Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Adj(15) As String
Public TrackName(15) As String
Public Country(15) As String
Public TempFile As String
Public ProgramDir As String
Public Gp2Dir As String
Public Read As String 'Public temp string
Public Read2 As String
Public Read3 As String
Public Read4 As String
Public CountExport As String
Public Gp2V As String
Public Gp2NameFile As String
Public dbFile As String
Public Exp As ExpVar
Public Tracks(15) As Boolean
Public FileType As Byte
'0=dat
'1=small pic
'2=big pic
'3=dir
'4=other

Public FileNum As Integer
Public FileNum2 As Integer
Public TreeNr As Integer

Public X As Long 'Public temp long
Public Count1 As Long
Public Count2 As Long

Public Const TH = "Gp2 Track Handler"
Public oFile As New clsFile
Public oData As New GP2Info
Public oReg As New clsReg
Public oDB As New clsLapTime
Public FileInfo As File
Public TheDate As Date
Public tVar As VarType
Public ProgStart As Boolean

Public TrackInfo As TypeTrackInfo

Public Const MAX_PATH = 260&

Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&

Private Const WM_USER = &H400
Private Const TB_SETSTYLE = WM_USER + 56
Private Const TB_GETSTYLE = WM_USER + 57
Private Const TBSTYLE_FLAT = &H800

Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Private Const LVS_EX_FULLROWSELECT = &H20

'Info om den öppnade filen
Public Type File
    Path As String
    Name As String
    Saved As Boolean
    Import As Boolean
End Type

Public Enum RecEnum
    F1gstate = 0
    RecFile = 1
End Enum

Public Type VarType
    iInt As Integer
    bByte As Byte
    lLong As Long
    dDouble As Double
End Type

'Variablar
Public Type ExpVar
    TrackNr As Integer
    Gp2FileNum As Integer
    F1FileNum As Integer
End Type

Public Enum QR
    Qual = 0
    Race = 1
End Enum

Public Type TypeTrackInfo
    Name As String
    Country As String
    Author As String
    Year As String
    Event As String
    Desc As String
    Laps As String
    Slot As String
    Tyre As String
    LengthMeters As String
    LapRecord As String
    LapRecordQualify As String
End Type

Public Sub FlatToolbar(tlb As Toolbar)
Dim lngStyle As Long
Dim lngResult As Long
Dim lngHWND As Long
    lngHWND = FindWindowEx(tlb.hWnd, 0&, "ToolbarWindow32", vbNullString)
    
    lngStyle = SendMessageLong(lngHWND, TB_GETSTYLE, 0&, 0&)
    '----------
    If lngStyle And TBSTYLE_FLAT Then
        lngStyle = lngStyle Xor TBSTYLE_FLAT
    Else
        lngStyle = lngStyle Or TBSTYLE_FLAT
    End If
    '----------
    'lngStyle = lngStyle Or TBSTYLE_FLAT
    lngResult = SendMessageLong(lngHWND, TB_SETSTYLE, 0, lngStyle)
    tlb.Refresh
End Sub

Public Sub FullRowSelect(ByVal lstvw As ListView)
Dim Rs As Long
Dim R As Long
    Rs = SendMessageLong(lstvw.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    Rs = Rs Or LVS_EX_FULLROWSELECT
    R = SendMessageLong(lstvw.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, Rs)
End Sub
