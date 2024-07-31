Attribute VB_Name = "Module1"
Option Explicit
Declare Function RegUpDown Lib "UpDown.ocx" Alias "DllRegisterServer" () As Long
Declare Function UnRegUpDown Lib "UpDown.ocx" Alias "DllUnregisterServer" () As Long
Declare Function RegJad2Jam Lib "Jad2Jam.ocx" Alias "DllRegisterServer" () As Long
Declare Function UnRegJad2Jam Lib "Jad2Jam.ocx" Alias "DllUnregisterServer" () As Long

Declare Function RegTabCtl32 Lib "TabCtl32.ocx" Alias "DllRegisterServer" () As Long
Declare Function RegComCtl32 Lib "ComCtl32.ocx" Alias "DllRegisterServer" () As Long

Const ERROR_SUCCESS = &H0

Public Function File_Exists(ByVal PathName As String) As Boolean
       File_Exists = IIf(Dir$(PathName) = "", False, True)
End Function
