VERSION 5.00
Begin VB.UserControl LabelMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00EED2C1&
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   750
   LockControls    =   -1  'True
   ScaleHeight     =   270
   ScaleWidth      =   750
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&System"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   510
   End
End
Attribute VB_Name = "LabelMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Event Declarations:
Public Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Private WithEvents MyTimer As clsTimer
Attribute MyTimer.VB_VarHelpID = -1
Private OldBackColor As Long

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

    BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"

End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."

    BorderStyle = UserControl.BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)

    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"

End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."

    Caption = Label1.Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    Label1.Caption() = New_Caption
    Width = Label1.Width + 210
    PropertyChanged "Caption"

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"

End Property

Private Sub Label1_Click()

    RaiseEvent Click
    MyBackGround True

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    MyBackGround True

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)
    MyBackGround True

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub MyBackGround(ByVal IsOver As Boolean)

    If IsOver Then
        If MyTimer Is Nothing Then
            Set MyTimer = New clsTimer
            MyTimer.StartTimer 10
            OldBackColor = BackColor
            BackColor = &HEED2C1
            BorderStyle = 1
          ElseIf BackColor <> &HEED2C1 Then
            OldBackColor = BackColor
            BackColor = &HEED2C1
            BorderStyle = 1
        End If
      Else
        MyTimer.StopTimer
        Set MyTimer = Nothing
        BackColor = OldBackColor
        BorderStyle = 0
    End If

End Sub

Private Sub MyTimer_OnTime(ByVal Int_Ticks As Long, ByVal DwTime As Long)

    If MyTimer.isMouseOver(UserControl.Hwnd) = False Then
        MyBackGround False
    End If

End Sub

Private Sub UserControl_Click()

    RaiseEvent Click
    MyBackGround True

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    MyBackGround True

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)
    MyBackGround True

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.Caption = PropBag.ReadProperty("Caption", "&System")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Width = Label1.Width + 210

    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)

End Sub

Private Sub UserControl_Resize()

    Height = 270
    Width = Label1.Width + 210

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "&System")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)

End Sub

