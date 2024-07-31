VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TabStrip1_Click()
    If TabStrip1.Tabs(1).Selected = True Then
        frameTrackInfo.Visible = True
        frameEditTrackInfo.Visible = False
    Else
        frameEditTrackInfo.Visible = True
        frameTrackInfo.Visible = False
    End If
End Sub
