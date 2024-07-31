VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.TreeView TreeView1 
      Height          =   5900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   10398
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0352
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":06A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":07B6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim nodX As Node    ' Create variable.
    Set nodX = TreeView1.Nodes.Add(, , "q", "Hej")
    NewTree
End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
    On Error GoTo errorstop
    Dim nodX As Node
    CountNr = Len(frmMain.File1.FileName)
    Read = UCase(Mid(frmMain.File1.FileName, CountNr - 2, 3))
    If (Read = "DAT") And (Len(TreeView1.DropHighlight.Key) = 3) Then
        Count3 = Mid(TreeView1.DropHighlight.Key, 2, 2) - 10
        GetDataFromLabel
        TreeView1.DropHighlight.Text = Trim(Str(Count3)) + ". " + frmMain.lblName
        Read = ReadINI("Track " + Trim(Str(Count3)), "TPath", ProgramDir + "\WorkCopy.lda")
        Count2 = 90
        Read2 = "DAT"
        Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 100)), "Track File: " + Read, 4, 4)
        TreeView1.DropHighlight = Nothing
        Exit Sub
    End If
    If ((Read = "BMP") Or (Read = "GIF")) And (Len(TreeView1.DropHighlight.Key) = 3) Then
        If (PicX = 640) Then
            Count3 = Mid(TreeView1.DropHighlight.Key, 2, 2) - 10
            Read = "Track " + Trim(Str(Count3))
            If (Mid(frmMain.Dir1.Path, 3, 1) = "\") And (Len(frmMain.Dir1.Path) = 3) Then
                Read = WriteINI(Read, "BPic", frmMain.Dir1.Path + frmMain.File1.FileName, ProgramDir + "\WorkCopy.lda")
            Else
                Read = WriteINI(Read, "BPic", frmMain.Dir1.Path + "\" + frmMain.File1.FileName, ProgramDir + "\WorkCopy.lda")
            End If
            Read = ReadINI("Track " + Trim(Str(Count3)), "BPic", ProgramDir + "\WorkCopy.lda")
            Count2 = 190
            Read2 = "BMP"
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 200)), "Full Pic: " + Read, 3, 3)
            TreeView1.DropHighlight = Nothing
            Exit Sub
        Else
            Count3 = Mid(TreeView1.DropHighlight.Key, 2, 2) - 10
            Read = "Track " + Trim(Str(Count3))
            If (Mid(frmMain.Dir1.Path, 3, 1) = "\") And (Len(frmMain.Dir1.Path) = 3) Then
                Read = WriteINI(Read, "SPic", frmMain.Dir1.Path + frmMain.File1.FileName, ProgramDir + "\WorkCopy.lda")
            Else
                Read = WriteINI(Read, "SPic", frmMain.Dir1.Path + "\" + frmMain.File1.FileName, ProgramDir + "\WorkCopy.lda")
            End If
            Read = ReadINI("Track " + Trim(Str(Count3)), "SPic", ProgramDir + "\WorkCopy.lda")
            Count2 = 290
            Read2 = "BMP"
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 300)), "Framed Pic: " + Read, 3, 3)
            TreeView1.DropHighlight = Nothing
            Exit Sub
        End If
    End If
    MsgBox "You can't place this object on this place!", vbInformation, TH
    TreeView1.DropHighlight = Nothing
    Exit Sub
errorstop:
    Select Case Err.Number
    Case "35602"
        CountNr = Mid(TreeView1.DropHighlight.Child.Key, 2, 3) - Count2
        If CountNr = Mid(TreeView1.DropHighlight.Key, 2, 2) Then
            X = TreeView1.DropHighlight.Child.Index
        Else
            CountNr = Mid(TreeView1.DropHighlight.Child.Next.Key, 2, 3) - Count2
            If CountNr = Mid(TreeView1.DropHighlight.Key, 2, 2) Then
                X = TreeView1.DropHighlight.Child.Next.Index
            Else
                CountNr = Mid(TreeView1.DropHighlight.Child.Next.Next.Key, 2, 3) - Count2
                If CountNr = Mid(TreeView1.DropHighlight.Key, 2, 2) Then
                    X = TreeView1.DropHighlight.Child.Next.Next.Index
                End If
            End If
        End If
        TreeView1.Nodes.Remove (X)
        If Read2 = "DAT" Then
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 100)), "Track File: " + Read, 4, 4)
        End If
        If Count2 = 290 Then
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 300)), "Framed Pic: " + Read, 3, 3)
        End If
        If Count2 = 190 Then
            Set nodX = TreeView1.Nodes.Add(TreeView1.DropHighlight.Key, tvwChild, "t" + Trim(Str(Count3 + 200)), "Full Pic: " + Read, 3, 3)
        End If
        TreeView1.DropHighlight = Nothing
    Case "91"
        Exit Sub
    Case Else
        MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
        Set TreeView1.DropHighlight = Nothing
        Exit Sub
    End Select
End Sub


Private Sub TreeView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If InDrag = True Then
        ' Set DropHighlight to the mouse's coordinates.
        Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
    End If
End Sub

Private Sub TreeView1_GotFocus()
    MDIForm1.Toolbar1.Buttons.Item(5).Enabled = True
End Sub

Private Sub TreeView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        DeleteTrack
    End If
End Sub

Private Sub TreeView1_LostFocus()
    MDIForm1.Toolbar1.Buttons.Item(5).Enabled = False
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
    If (TreeView1.SelectedItem.Text = "GP2 Track's") Or (Form1.TreeView1.SelectedItem.Key = "q") Then
        SaveLastClick
        frmMain.frameFile.Visible = True
        frmMain.frameTrackInfo.Visible = True
        frmMain.frameInfo.Visible = False
        frmMain.frameData.Visible = False
        frmMain.frmaeTime.Visible = False
        frmMain.frameFile.Top = frmMain.frameInfo.Top
        frmMain.frameFile.Left = frmMain.frameInfo.Left
        frmMain.imgFramed.Picture = Nothing
        frmMain.imgFull.Picture = Nothing
        frmMain.frameGlobal.Visible = False
        frmMain.framePlayer.Visible = False
        frmMain.lblNote.Visible = False
        LastClick = "GP2 Track's"
        CurrentRecord = 0
        Exit Sub
    End If
    frmMain.imgPre.Picture = Nothing
    If Mid(TreeView1.SelectedItem.Key, 1, 1) = "t" Then
        SaveLastClick
        frmMain.frameFile.Visible = False
        frmMain.frameTrackInfo.Visible = False
        frmMain.frameInfo.Visible = True
        frmMain.frameData.Visible = True
        frmMain.frmaeTime.Visible = True
        frmMain.frameGlobal.Visible = False
        frmMain.framePlayer.Visible = False
        frmMain.lblNote.Visible = False

        Count4 = Mid(TreeView1.SelectedItem.Key, 2, Len(TreeView1.SelectedItem.Key) - 1)
        If (Count4 > 10) And (Count4 < 100) Then Count4 = Count4 - 10
        If (Count4 > 100) And (Count4 < 200) Then Count4 = Count4 - 100
        If (Count4 > 200) And (Count4 < 300) Then Count4 = Count4 - 200
        If Count4 > 300 Then Count4 = Count4 - 300
        
        If Count4 = CurrentRecord Then Exit Sub
        'If CurrentRecord <> 0 Then SaveFileData
        GetFileData Count4
        CurrentRecord = Count4
        LastClick = "Track"
        Exit Sub
    End If
    If TreeView1.SelectedItem.Key = "e" Then
        SaveLastClick
        frmMain.frameFile.Visible = False
        frmMain.frameTrackInfo.Visible = False
        frmMain.frameInfo.Visible = False
        frmMain.frameData.Visible = False
        frmMain.frmaeTime.Visible = False
        frmMain.frameGlobal.Visible = True
        frmMain.framePlayer.Visible = True
        GetPlayerData
        LastClick = "Player"
        Read = File_Exists(Gp2Dir + "\f1gstate.sav")
        If Read = False Then
            frmMain.HScroll2.Enabled = False
            frmMain.lblNote.Visible = True
        Else
            frmMain.HScroll2.Enabled = True
            frmMain.lblNote.Visible = False
        End If
    End If
    Exit Sub
errortrap:
    Select Case Err.Number
        Case Else
            MsgBox "Error # " + Str(Err.Number) + " " + Err.Description
            Set TreeView1.DropHighlight = Nothing
            Exit Sub
    End Select
    Exit Sub
End Sub
