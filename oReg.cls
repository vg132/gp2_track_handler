VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum SELECT_HKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Enum REG_DATA
    REG_SZ = 1
    REG_DWORD = 4
End Enum

Const ERROR_SUCCESS = 0&

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub DeleteValue(ByVal hKey As SELECT_HKEY, ByVal strPath As String, ByVal strValue As String)
Dim hCurKey As Long
Dim lRegResult As Long
    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, strValue)
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Sub DeleteKey(ByVal hKey As SELECT_HKEY, ByVal strPath As String)
Dim lRegResult As Long
    lRegResult = RegDeleteKey(hKey, strPath)
End Sub

Public Sub CreateKey(ByVal hKey As SELECT_HKEY, strPath As String)
Dim hCurKey As Long
Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function SaveValue(hKey As SELECT_HKEY, lType As REG_DATA, strPath As String, strValue As String, Optional strData As String, Optional lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    Select Case lType
        Case REG_SZ
            lRegResult = RegSetValueEx(hCurKey, strValue, 0, lType, ByVal strData, Len(strData))
        Case REG_DWORD
            lRegResult = RegSetValueEx(hCurKey, strValue, 0&, lType, lData, 4)
    End Select
    lRegResult = RegCloseKey(hCurKey)
End Function

Public Function GetValue(hKey As SELECT_HKEY, strPath As String, strValue As String) As Variant
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long
Dim lBuffer As Long
Dim byBuffer() As Byte

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lValueType = REG_DWORD Then
        lDataBufferSize = 4
        lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
        GetValue = lBuffer
    ElseIf lValueType = REG_SZ Then
        strBuffer = String(lDataBufferSize, " ")
        lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
        intZeroPos = InStr(strBuffer, Chr$(0))
        If intZeroPos > 0 Then
            GetValue = Left$(strBuffer, intZeroPos - 1)
        Else
            GetValue = strBuffer
        End If
    End If
    lRegResult = RegCloseKey(hCurKey)
End Function
