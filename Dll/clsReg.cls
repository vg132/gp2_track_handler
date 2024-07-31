VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Enum SELECT_HKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
End Enum

Enum REG_DATA
    REG_SZ = 1
    REG_DWORD = 4                      ' 32-bit number
End Enum

Const ERROR_SUCCESS = 0&

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function SaveValue(hKey As SELECT_HKEY, lType As REG_DATA, Path As String, ByVal ValueName As String, Optional ByVal strData As String, Optional ByVal lData As Long)
Dim hCurKey As Long
Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, Path, hCurKey)
    Select Case lType
        Case REG_SZ
            lRegResult = RegSetValueEx(hCurKey, ValueName, 0, lType, ByVal strData, Len(strData))
        Case REG_DWORD
            lRegResult = RegSetValueEx(hCurKey, ValueName, 0&, lType, lData, 4)
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

Public Function GetAllValues(hKey As SELECT_HKEY, strPath As String) As Variant
Dim lRegResult As Long
Dim hCurKey As Long
Dim lValueNameSize As Long
Dim strValueName As String
Dim lCounter As Long
Dim byDataBuffer(4000) As Byte
Dim lDataBufferSize As Long
Dim lValueType As Long
Dim strNames() As String
Dim lTypes() As Long
Dim intZeroPos As Integer

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    Do
        ' Initialise bufffers
        lValueNameSize = 255
        strValueName = String$(lValueNameSize, " ")
        lDataBufferSize = 4000
        lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
        If lRegResult = ERROR_SUCCESS Then
            ' Save the type
            ReDim Preserve strNames(lCounter) As String
            ReDim Preserve lTypes(lCounter) As Long
            lTypes(UBound(lTypes)) = lValueType
            'Tidy up string and save it
            intZeroPos = InStr(strValueName, Chr$(0))
            If intZeroPos > 0 Then
                strNames(UBound(strNames)) = Left$(strValueName, intZeroPos - 1)
            Else
                strNames(UBound(strNames)) = strValueName
            End If
            lCounter = lCounter + 1
        Else
            Exit Do
        End If
    Loop
    'Move data into array
    If lCounter > 0 Then
        ReDim Finisheddata(UBound(strNames), 0 To 1) As Variant
        For lCounter = 0 To UBound(strNames)
            Finisheddata(lCounter, 0) = strNames(lCounter)
            Finisheddata(lCounter, 1) = GetValue(hKey, strPath, strNames(lCounter))
        Next
        GetAllValues = Finisheddata
    Else
        GetAllValues = vbNull
    End If
End Function
