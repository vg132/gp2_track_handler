VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim FileNum As Integer
Dim Data As String
Dim X As Long

Public Function GetRecord(ByVal File As String, ByVal RecNr As Long) As String
    FileNum = FreeFile
    Open File For Input As FileNum
    For X = 0 To RecNr
        Line Input #FileNum, Data
    Next
    GetRecord = Data
    Close FileNum
End Function

Public Sub SaveNew(ByVal File As String, ByVal Data As String)
    FileNum = FreeFile
    Open File For Append As FileNum
    Print #FileNum, Data
    Close FileNum
End Sub

Public Function RecCount(ByVal File As String) As Long
    FileNum = FreeFile
    Open File For Binary As FileNum
    Close FileNum
    FileNum = FreeFile
    Open File For Input As FileNum
    X = 0
    Do Until EOF(FileNum)
        Line Input #FileNum, Data
        X = X + 1
    Loop
    RecCount = X
    Close FileNum
End Function

Public Function DeleteRecord(ByVal File As String, ByVal TempFile As String, ByVal RecNr As Long)
Dim Rec As Long
Dim FileNum2 As Integer
    FileNum = FreeFile
    Open File For Input As FileNum
    FileNum2 = FreeFile
    Open TempFile For Append As FileNum2
    X = 0
    Do Until EOF(FileNum)
        Line Input #FileNum, Data
        If X <> RecNr Then Print #FileNum2, Data
        X = X + 1
    Loop
    Close FileNum
    Close FileNum2
    Kill (File)
    FileCopy TempFile, File
    Kill (TempFile)
End Function
