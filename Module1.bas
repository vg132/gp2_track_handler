Attribute VB_Name = "Misc"
'-- Almäna variablar för programmet
Option Explicit

Public oMisc As New TrackHandler.Misc
Public oData As New TrackHandler.Data
Public GP2V As GP2Ver
Public FileInfo As Save
Public TrackEdit As TEdit

Public Enum SaveType2
    FileNew = 1
    FileOpen = 2
    FileImport = 3
End Enum

Public Type TEdit
    Mid1 As String * 500
    Mid2 As String * 100
End Type

Public Type Save
    FilePath As String * 255
    FileName As String * 255
    FileType As SaveType2
End Type

Dim V As String
Public GP2TH As String
Public TH As String

Public TheDate As Date
Public TempDouble As Single

Public SelectVersion As Boolean
Public OpenPage As Long
Public X17 As Integer

Public Adj(15) As String
Public TrackName(15) As String
Public Country(15) As String

Global GP2NameFile As String
Global DefaultTrackPath As String
Global LastClick As String
Global TrackPath As String
Global GP2Country As String
Global TargetFile As String
Global SourceFile As String
Global Read As String   'Läs från en fil till denna
Global Read2 As String  'Samma
Global Read3 As String
Global Read4 As String
Global AddNode As String
Global NodeName As String   'Vad ska grenen som läggs till heta
Global Gp2Dir As String 'Var finns gp2
Global ProgramDir As String

Global NoSupport As Boolean
Global InDrag As Boolean 'Variabel som används när filer dras

Global TimeBase As TimeDataBase
Global WareCount As Double

Global LastRecord2 As Long
Global CurrentRecord3 As Long
Global RecordLen As Long
Global CurrentRecord As Long    'Vad som ska visas
Global LastRecord As Long
Global FileSize As Long
Global Count1 As Long
Global Count2 As Long
Global Count3 As Long
Global CountNr As Long
Global X As Long    'Vart i filen ska programmet läsa

Global DataBaseFileNum As Integer
Global F1SaveFileNum As Integer
Global CurrentRecord2 As Integer
Global GP2FileNum As Integer
Global Responce As Integer  'Vad svarar användaren på frågor
Global FileNum As Integer   'Fil nummer
Global FileNum2 As Integer  'Fil nummer
Global PicX As Integer  'Hur vid är bilden
Global PicY As Integer  'Hur hög är bilden
Global Count4 As Integer
Global CountExport As Integer

'--Information om var olika saker finns i gp2.exe

'Time DataBase
Type TimeDataBase
    TName As String * 22
    QTime As String * 8
    RTime As String * 8
    QTeam As String * 12
    RTeam As String * 12
    QDriver As String * 22
    RDriver As String * 22
    QDate As String * 10
    RDate As String * 10
End Type

'--Fil Struktur 1.0/1.1
Type SaveFileInfo
    Path As String * 200
    Country As String * 20
    Country2 As String * 20
    Track As String * 30
    Laps As String * 3
    Pic As String * 200
    Pic2 As String * 200
    Length As String * 4
    EXE As String * 16
    CarSet As String * 200
End Type

'--Fil Struktur 1.2
Type SaveFileInfo2
    Track As String * 22
    Country As String * 22
    Country2 As String * 22
    Laps As String * 3
    Ware As String * 5
    Pic As String * 100
    Pic2 As String * 100
    Path As String * 100
    CarSet As String * 100
    Length As String * 4
    Points As String * 52
    EXE As String * 48
End Type

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const ES_NUMBER = &H2000&

Public Sub SelectText()

Dim i As Integer
Dim oMyTextBox As Object

Set oMyTextBox = Screen.ActiveControl
    If TypeName(oMyTextBox) = "TextBox" Then
        i = Len(oMyTextBox.Text)
        oMyTextBox.SelStart = 0
        oMyTextBox.SelLength = i
    End If

End Sub

Public Sub Dec2Bin(MyNum As Integer)
Dim LoopCounter As Integer
    Do
        If (MyNum And 2 ^ LoopCounter) = 2 ^ LoopCounter Then
            Read2 = Read2 & "1"
        Else
            Read2 = Read2 & "0"
        End If
        LoopCounter = LoopCounter + 1
    Loop Until 2 ^ LoopCounter > MyNum
End Sub

Public Sub GetGP2Version()
'-- Läs GP2.exe och titta vad det är för version (språk)
    V = "Version 1.0b"
    FileNum = FreeFile
    Open Gp2Dir + "\gp2.exe" For Binary As FileNum
    Read = String(23, " ")
    Get #FileNum, 5671742, Read
    If Read = "US English Version 1.0b" Then
        Close FileNum
        GP2V = US
        GP2Country = "American " + V
        Exit Sub
    End If
    Get #FileNum, 5671743, Read
    If Read = "UK English Version 1.0b" Then
        Close FileNum
        GP2V = UK
        GP2Country = "UK English " + V
        Exit Sub
    End If
    Get #FileNum, 5673614, Read
    If Read = "Nederlandse versie 1.0b" Then
        Close FileNum
        GP2V = NL
        GP2Country = "Dutch " + V
        Exit Sub
    End If
    Read = String(5, " ")
    Get #FileNum, 5675458, Read
    If Read = "Versi" Then
        Close FileNum
        GP2V = Sp
        GP2Country = "Spanish " + V
        Exit Sub
    End If
    Read = String(7, " ")
    Get #FileNum, 5674990, Read
    If Read = "Version" Then
        Close FileNum
        GP2V = FR
        GP2Country = "French " + V
        Exit Sub
    End If
    Read = String(8, " ")
    Get #FileNum, 5674331, Read
    If Read = "Versione" Then
        Close FileNum
        GP2V = IT
        GP2Country = "Italian " + V
        Exit Sub
    End If
    Read = String(21, " ")
    Get #FileNum, 5674544, Read
    If Read = "Deutsche Ausgabe 1.0b" Then
        Close FileNum
        GP2V = TY
        GP2Country = "German " + V
        Exit Sub
    End If
    Close FileNum
    If SelectVersion = False Then
        Responce = MsgBox("Sorry but this program was not able to detect your GP2 Version! Do you want to select your version manualy?", vbExclamation + vbCritical + vbYesNo, TH)
        If Responce = vbYes Then frmSelect.Show , MDIForm1
        If Responce = vbNo Then End
    End If
End Sub

Public Sub NewTree()
    MDIForm1.TreeView1.Nodes.Remove (1)
    Dim nodX As Node    ' Create variable.
    
    Set nodX = MDIForm1.TreeView1.Nodes.Add(, , "q", GP2TH, 1, 2)
    Set nodX = MDIForm1.TreeView1.Nodes.Add("q", tvwChild, "r", "GP2 Track's", 1, 2)
    Set nodX = MDIForm1.TreeView1.Nodes.Add("q", tvwChild, "e", "GP2 Settings", 1, 2)
    
    X = 11
    Do Until X > 26
        Set nodX = MDIForm1.TreeView1.Nodes.Add("r", tvwChild, "t" + Trim(Str(X)), "Track " + Trim(Str(X - 10)), 1, 2)
        X = X + 1
    Loop
    nodX.EnsureVisible ' Show all nodes.
End Sub

Public Sub MakeNewTree()
    MDIForm1.TreeView1.Nodes.Remove (1)
    Dim nodX As Node    ' Create variable.
    
    Set nodX = MDIForm1.TreeView1.Nodes.Add(, , "q", GP2TH, 1, 2)
    Set nodX = MDIForm1.TreeView1.Nodes.Add("q", tvwChild, "r", "GP2 Track's", 1, 2)
    Set nodX = MDIForm1.TreeView1.Nodes.Add("q", tvwChild, "e", "GP2 Settings", 1, 2)
    nodX.EnsureVisible
    X = 1
    Read3 = ProgramDir + "\WorkCopy.lda"
    Do Until X > 16
        Read = oMisc.ReadINI("Track " + Trim(Str(X)), "Name", Read3)
        If Read <> "" Then
            Set nodX = MDIForm1.TreeView1.Nodes.Add("r", tvwChild, "t" + Trim(Str(X + 10)), Trim(Str(X)) + ". " + Read, 1, 2)
        Else
            Set nodX = MDIForm1.TreeView1.Nodes.Add("r", tvwChild, "t" + Trim(Str(X + 10)), "Track " + Trim(Str(X)), 1, 2)
        End If
        X = X + 1
    Loop
    nodX.EnsureVisible
    X = 1
    Do Until X > 16
        Read = oMisc.ReadINI("Track " + Trim(Str(X)), "TPath", Read3)
        If Read <> "" Then
            Set nodX = MDIForm1.TreeView1.Nodes.Add("t" + Trim(Str(X + 10)), tvwChild, "t" + Trim(Str(X + 100)), "Track File: " + Read, 4, 4)
        End If
        X = X + 1
    Loop
    X = 1
    Do Until X > 16
        Read = oMisc.ReadINI("Track " + Trim(Str(X)), "BPic", Read3)
        If Read <> "" Then
            Set nodX = MDIForm1.TreeView1.Nodes.Add("t" + Trim(Str(X + 10)), tvwChild, "t" + Trim(Str(X + 200)), "Full Pic: " + Read, 3, 3)
        End If
        Read = oMisc.ReadINI("Track " + Trim(Str(X)), "SPic", Read3)
        If Read <> "" Then
            Set nodX = MDIForm1.TreeView1.Nodes.Add("t" + Trim(Str(X + 10)), tvwChild, "t" + Trim(Str(X + 300)), "Framed Pic: " + Read, 3, 3)
        End If
        X = X + 1
    Loop
End Sub

Public Sub Resize_Form()
    On Error Resume Next
    MDIForm1.TreeView1.Height = MDIForm1.Height - 1600
    MDIForm1.TreeView1.Top = 100
    MDIForm1.TreeView1.Left = 100

    MDIForm1.picData.Width = MDIForm1.Width - MDIForm1.picTree.Width - 100
    
    MDIForm1.StatusBar1.Panels(1).Width = ((MDIForm1.Width - 1800) / 3)
    MDIForm1.StatusBar1.Panels(2).Width = ((MDIForm1.Width - 1800) / 3)
    MDIForm1.StatusBar1.Panels(3).Width = ((MDIForm1.Width - 1800) / 3)
End Sub

Public Sub GP2AidsSet()
    X = 0
    Do Until X = 7
        MDIForm1.R(X) = MDIForm1.On1(X)
        MDIForm1.R(X).Tag = "Off"
        X = X + 1
    Loop
    MDIForm1.A(0).Picture = MDIForm1.Off(0).Picture
    MDIForm1.A(0).Tag = "On"
    X = 1
    Do Until X = 7
        MDIForm1.A(X) = MDIForm1.On1(X)
        MDIForm1.A(X).Tag = "Off"
        X = X + 1
    Loop
    MDIForm1.S(0).Picture = MDIForm1.Off(0).Picture
    MDIForm1.S(0).Tag = "On"
    X = 1
    Do Until X = 7
        MDIForm1.S(X) = MDIForm1.On1(X)
        MDIForm1.S(X).Tag = "Off"
        X = X + 1
    Loop
    MDIForm1.P(0).Picture = MDIForm1.Off(0).Picture
    MDIForm1.P(0).Tag = "On"
    MDIForm1.P(1).Picture = MDIForm1.On1(1).Picture
    MDIForm1.P(1).Tag = "Off"
    MDIForm1.P(2).Picture = MDIForm1.Off(2).Picture
    MDIForm1.P(2).Tag = "On"
    MDIForm1.P(3).Picture = MDIForm1.Off(3).Picture
    MDIForm1.P(3).Tag = "On"
    MDIForm1.P(4).Picture = MDIForm1.On1(4).Picture
    MDIForm1.P(4).Tag = "Off"
    MDIForm1.P(5).Picture = MDIForm1.On1(5).Picture
    MDIForm1.P(5).Tag = "Off"
    MDIForm1.P(6).Picture = MDIForm1.On1(6).Picture
    MDIForm1.P(6).Tag = "Off"

    MDIForm1.AC(0).Picture = MDIForm1.Off(0).Picture
    MDIForm1.AC(0).Tag = "On"
    MDIForm1.AC(1).Picture = MDIForm1.On1(1).Picture
    MDIForm1.AC(1).Tag = "Off"
    MDIForm1.AC(2).Picture = MDIForm1.Off(2).Picture
    MDIForm1.AC(2).Tag = "On"
    MDIForm1.AC(3).Picture = MDIForm1.Off(3).Picture
    MDIForm1.AC(3).Tag = "On"
    MDIForm1.AC(4).Picture = MDIForm1.Off(4).Picture
    MDIForm1.AC(4).Tag = "On"
    MDIForm1.AC(5).Picture = MDIForm1.Off(5).Picture
    MDIForm1.AC(5).Tag = "On"
    MDIForm1.AC(6).Picture = MDIForm1.On1(6).Picture
    MDIForm1.AC(6).Tag = "Off"
End Sub

Public Sub DeleteTrack()
    On Error GoTo ErrorTrap
    Read = Mid(MDIForm1.TreeView1.SelectedItem.Key, 2, 4)

    Rensa

    If Len(MDIForm1.TreeView1.SelectedItem.Key) > 3 Then
        If (Read > 100) And (Read < 200) Then
            X = Read - 100
            SaveLastClick
            MDIForm1.TreeView1.SelectedItem.Parent.Text = "Track " + Trim(Str(X))
            MDIForm1.TreeView1.Nodes.Remove (MDIForm1.TreeView1.SelectedItem.Index)
            Exit Sub
        ElseIf (Read > 200) And (Read < 300) Then
            MDIForm1.txtFullPath = ""
            Set MDIForm1.imgFull = Nothing
            MDIForm1.TreeView1.Nodes.Remove (MDIForm1.TreeView1.SelectedItem.Index)
            MDIForm1.lblFull.Visible = False
            SaveLastClick
            Exit Sub
        ElseIf (Read > 300) And (Read < 400) Then
            MDIForm1.txtFramedPath = ""
            Set MDIForm1.imgFramed = Nothing
            MDIForm1.TreeView1.Nodes.Remove (MDIForm1.TreeView1.SelectedItem.Index)
            MDIForm1.lblFramed.Visible = False
            SaveLastClick
            Exit Sub
        End If
    Else
        Read = Mid(MDIForm1.TreeView1.SelectedItem.Key, 2, 4)
        X = Read - 10
        MDIForm1.txtFullPath = ""
        Set MDIForm1.imgFull = Nothing
        MDIForm1.txtFramedPath = ""
        Set MDIForm1.imgFramed = Nothing
        SaveLastClick
        Count1 = MDIForm1.TreeView1.SelectedItem.Children
        MDIForm1.TreeView1.SelectedItem.Text = "Track " + Trim(Str(X))
        Do Until Count1 = 0
            MDIForm1.TreeView1.Nodes.Remove (MDIForm1.TreeView1.SelectedItem.Child.Index)
            Count1 = Count1 - 1
        Loop
    End If
    Exit Sub

ErrorTrap:
    Select Case Err.Number
        Case "13"
            MsgBox "You can't delete this item.", vbInformation, TH
            Exit Sub
        Case Else
            MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
    End Select
End Sub

Public Sub Info()
    Read2 = oMisc.ReadINI("Misc", "EXEPath", ProgramDir + "\WorkCopy.lda")
    Read2 = Read2 & " -F"
    Dim RetVal
    RetVal = Shell(Read2, vbNormalFocus)
End Sub

Public Sub Rensa()
    With MDIForm1
        .txtAdjectiv = ""
        .txtPath = ""
        .txtName = ""
        .txtQDate = ""
        .txtQDriver = ""
        .txtQTeam = ""
        .txtQTime = ""
        .txtRDate = ""
        .txtRDriver = ""
        .txtRTeam = ""
        .txtRTime = ""
        .VScroll1.Value = 3
        .txtLaps = ""
        .txtLength = ""
        .txtCountry = ""
        .txtTire = ""
    End With
End Sub

Public Sub ShowRecent()
Dim RetVal
    RetVal = oMisc.RecentFile(1, "", "", GP2TH, Check)
    If RetVal = 0 Then Exit Sub
    If RetVal > 0 Then
        MDIForm1.mnuOpen1.Visible = True
        MDIForm1.mnuOpen1.Caption = GetSetting(GP2TH, "RecentFile", "Name1")
        MDIForm1.mnuOpen1.Tag = GetSetting(GP2TH, "RecentFile", "Path1")
        MDIForm1.mnuSep10.Visible = True
    End If
    If RetVal > 1 Then
        MDIForm1.mnuOpen2.Visible = True
        MDIForm1.mnuOpen2.Caption = GetSetting(GP2TH, "RecentFile", "Name2")
        MDIForm1.mnuOpen2.Tag = GetSetting(GP2TH, "RecentFile", "Path2")
    End If
    If RetVal > 2 Then
        MDIForm1.mnuOpen3.Visible = True
        MDIForm1.mnuOpen3.Caption = GetSetting(GP2TH, "RecentFile", "Name3")
        MDIForm1.mnuOpen3.Tag = GetSetting(GP2TH, "RecentFile", "Path3")
    End If
End Sub
