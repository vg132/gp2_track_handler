VERSION 5.00
Begin VB.UserControl UpDown 
   ClientHeight    =   285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   ScaleHeight     =   285
   ScaleWidth      =   645
   ToolboxBitmap   =   "UpDown.ctx":0000
   Begin VB.VScrollBar vscUpDown 
      Height          =   270
      Left            =   495
      Max             =   0
      Min             =   100
      TabIndex        =   1
      Top             =   0
      Width           =   150
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "UpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000&

Event Change()

Private Sub txtValue_Change()
    RaiseEvent Change
    If txtValue.Text = "" Then Exit Sub
    If txtValue.Text >= vscUpDown.Max And txtValue.Text <= vscUpDown.Min Then
        vscUpDown.Value = txtValue.Text
    ElseIf txtValue.Text > vscUpDown.Min Then
        vscUpDown.Value = vscUpDown.Min
        txtValue.Text = vscUpDown.Value
    End If
End Sub

Private Sub txtValue_GotFocus()
    txtValue.SelStart = 0
    txtValue.SelLength = Len(txtValue.Text)
End Sub

Private Sub UserControl_Initialize()
Dim X As Long
    X = GetWindowLong(txtValue.hWnd, GWL_STYLE)
    X = X Or ES_NUMBER
    Call SetWindowLong(txtValue.hWnd, GWL_STYLE, X)
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < 200 Then UserControl.Width = 200
    txtValue.Width = UserControl.Width - 150
    txtValue.Height = UserControl.Height
    vscUpDown.Left = UserControl.Width - 150
    vscUpDown.Height = UserControl.Height
End Sub

Private Sub vscUpDown_Change()
    RaiseEvent Change
    txtValue.Text = vscUpDown.Value
End Sub

Private Sub vscUpDown_GotFocus()
    vscUpDown.Value = txtValue.Text
End Sub

Private Sub vscUpDown_Scroll()
    txtValue.Text = vscUpDown.Value
End Sub

Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of an object."
    Value = vscUpDown.Value
End Property

Public Property Let Value(ByVal iValue As Integer)
    vscUpDown.Value = iValue
    PropertyChanged "Value"
    txtValue.Text = vscUpDown.Value
End Property

Public Property Get Max() As Integer
Attribute Max.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
    Max = vscUpDown.Min
End Property

Public Property Let Max(ByVal New_Max As Integer)
    vscUpDown.Min() = New_Max
    PropertyChanged "Max"
End Property

Public Property Get Min() As Integer
Attribute Min.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
    Min = vscUpDown.Max
End Property

Public Property Let Min(ByVal New_Min As Integer)
    vscUpDown.Max() = New_Min
    PropertyChanged "Min"
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = txtValue.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtValue.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    vscUpDown.Value = PropBag.ReadProperty("Value", 0)
    vscUpDown.Min = PropBag.ReadProperty("Max", 100)
    vscUpDown.Max = PropBag.ReadProperty("Min", 0)
    txtValue.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    txtValue.Text = PropBag.ReadProperty("Text", "")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", vscUpDown.Value, 0)
    Call PropBag.WriteProperty("Max", vscUpDown.Min, 100)
    Call PropBag.WriteProperty("Min", vscUpDown.Max, 0)
    Call PropBag.WriteProperty("ToolTipText", txtValue.ToolTipText, "")
    Call PropBag.WriteProperty("Text", txtValue.Text, "")
End Sub

Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtValue.Text
End Property
