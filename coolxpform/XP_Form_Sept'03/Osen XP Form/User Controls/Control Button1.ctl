VERSION 5.00
Begin VB.UserControl ControlButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   LockControls    =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   2640
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      Picture         =   "Control Button1.ctx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Width           =   315
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3810
      Top             =   2340
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   11
      Left            =   720
      Picture         =   "Control Button1.ctx":0582
      Top             =   3240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   10
      Left            =   1080
      Picture         =   "Control Button1.ctx":0B04
      Top             =   3240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   9
      Left            =   1440
      Picture         =   "Control Button1.ctx":1086
      ToolTipText     =   "close"
      Top             =   3240
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   14
      Left            =   720
      Picture         =   "Control Button1.ctx":1608
      Top             =   3975
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   13
      Left            =   1080
      Picture         =   "Control Button1.ctx":1B8A
      Top             =   3975
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   12
      Left            =   1440
      Picture         =   "Control Button1.ctx":210C
      ToolTipText     =   "close"
      Top             =   3975
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   11
      Left            =   720
      Picture         =   "Control Button1.ctx":268E
      Top             =   4350
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   10
      Left            =   1080
      Picture         =   "Control Button1.ctx":2C10
      Top             =   4350
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   9
      Left            =   1440
      Picture         =   "Control Button1.ctx":3192
      ToolTipText     =   "close"
      Top             =   4350
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   11
      Left            =   720
      Picture         =   "Control Button1.ctx":3714
      Top             =   3600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   10
      Left            =   1080
      Picture         =   "Control Button1.ctx":3C96
      Top             =   3600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   9
      Left            =   1440
      Picture         =   "Control Button1.ctx":4218
      ToolTipText     =   "close"
      Top             =   3600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   11
      Left            =   1830
      Picture         =   "Control Button1.ctx":479A
      ToolTipText     =   "close"
      Top             =   3990
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   8
      Left            =   1800
      Picture         =   "Control Button1.ctx":4D1C
      Top             =   3255
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   10
      Left            =   2190
      Picture         =   "Control Button1.ctx":529E
      Top             =   3990
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   8
      Left            =   1800
      Picture         =   "Control Button1.ctx":5820
      Top             =   4350
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   8
      Left            =   1800
      Picture         =   "Control Button1.ctx":5DA2
      Top             =   3600
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   7
      Left            =   1800
      Picture         =   "Control Button1.ctx":6324
      Top             =   2010
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   7
      Left            =   1800
      Picture         =   "Control Button1.ctx":68A6
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   8
      Left            =   1800
      Picture         =   "Control Button1.ctx":6E28
      Top             =   2385
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   7
      Left            =   1800
      Picture         =   "Control Button1.ctx":73AA
      Top             =   1665
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   9
      Left            =   2160
      Picture         =   "Control Button1.ctx":792C
      ToolTipText     =   "close"
      Top             =   2385
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   6
      Left            =   1440
      Picture         =   "Control Button1.ctx":7EAE
      ToolTipText     =   "close"
      Top             =   2010
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   5
      Left            =   1080
      Picture         =   "Control Button1.ctx":8430
      Top             =   2010
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   4
      Left            =   720
      Picture         =   "Control Button1.ctx":89B2
      Top             =   2010
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   6
      Left            =   1440
      Picture         =   "Control Button1.ctx":8F34
      ToolTipText     =   "close"
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   5
      Left            =   1080
      Picture         =   "Control Button1.ctx":94B6
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   4
      Left            =   720
      Picture         =   "Control Button1.ctx":9A38
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   7
      Left            =   1440
      Picture         =   "Control Button1.ctx":9FBA
      ToolTipText     =   "close"
      Top             =   2385
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   6
      Left            =   1080
      Picture         =   "Control Button1.ctx":A53C
      Top             =   2385
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   5
      Left            =   720
      Picture         =   "Control Button1.ctx":AABE
      Top             =   2385
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   6
      Left            =   1440
      Picture         =   "Control Button1.ctx":B040
      ToolTipText     =   "close"
      Top             =   1650
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   5
      Left            =   1080
      Picture         =   "Control Button1.ctx":B5C2
      Top             =   1650
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   4
      Left            =   720
      Picture         =   "Control Button1.ctx":BB44
      Top             =   1650
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   0
      Left            =   750
      Picture         =   "Control Button1.ctx":C0C6
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   1
      Left            =   1110
      Picture         =   "Control Button1.ctx":C648
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   2
      Left            =   1470
      Picture         =   "Control Button1.ctx":CBCA
      ToolTipText     =   "close"
      Top             =   90
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   0
      Left            =   750
      Picture         =   "Control Button1.ctx":D14C
      Top             =   825
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   1
      Left            =   1110
      Picture         =   "Control Button1.ctx":D6CE
      Top             =   825
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   2
      Left            =   1470
      Picture         =   "Control Button1.ctx":DC50
      ToolTipText     =   "close"
      Top             =   825
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   0
      Left            =   750
      Picture         =   "Control Button1.ctx":E1D2
      Top             =   1200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   1
      Left            =   1110
      Picture         =   "Control Button1.ctx":E754
      Top             =   1200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   2
      Left            =   1470
      Picture         =   "Control Button1.ctx":ECD6
      ToolTipText     =   "close"
      Top             =   1200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   0
      Left            =   750
      Picture         =   "Control Button1.ctx":F258
      Top             =   450
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   1
      Left            =   1110
      Picture         =   "Control Button1.ctx":F7DA
      Top             =   450
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   2
      Left            =   1470
      Picture         =   "Control Button1.ctx":FD5C
      ToolTipText     =   "close"
      Top             =   450
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   4
      Left            =   2190
      Picture         =   "Control Button1.ctx":102DE
      ToolTipText     =   "close"
      Top             =   825
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbClose 
      Height          =   315
      Index           =   3
      Left            =   1830
      Picture         =   "Control Button1.ctx":10860
      Top             =   105
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMax 
      Height          =   315
      Index           =   3
      Left            =   1830
      Picture         =   "Control Button1.ctx":10DE2
      Top             =   825
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbRestore 
      Height          =   315
      Index           =   3
      Left            =   1830
      Picture         =   "Control Button1.ctx":11364
      Top             =   1200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image CbMin 
      Height          =   315
      Index           =   3
      Left            =   1830
      Picture         =   "Control Button1.ctx":118E6
      Top             =   450
      Visible         =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "ControlButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ButtonType
    MinButton = 0
    MaxButton = 1
    CloseButton = 2
    RestoreButton = 3
End Enum
'mouse over effects
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Enum XPVisualTheme
    Blue = 0
    [Olive Green] = 1
    Silver = 2
End Enum

Private Type POINTAPI
  X As Long
  Y As Long
End Type
'Default Property Values:
Const m_def_IsActivate = True
Const m_def_Theme = 0
Const m_def_ButtonStyle = 0
'Property Variables:
Dim m_IsActivate As Boolean
Dim m_Theme As XPVisualTheme
Dim m_ButtonStyle As ButtonType
'Event Declarations:
Public Event Click() 'MappingInfo=PicMain,PicMain,-1,Click
Private MyButton As Integer

Private Sub PicMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        MyButton = 1
        ChangeIfOver False
        Timer1.Enabled = True
    End If
End Sub

Private Sub PicMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MyButton = Button
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim MyOver As Boolean
    MyOver = isMouseOver(PicMain.Hwnd)
    If (MyButton = 1 And MyOver) Then
        ChangeIfOver False
    Else
        ChangeIfOver MyOver
    End If
    If MyOver = False Then
        RefreshControl IsActivate
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Initialize()
    RefreshControl
End Sub

Private Sub UserControl_Resize()
    Width = 315
    Height = 315
End Sub

Private Function isMouseOver(ByVal Hwnd As Long) As Boolean
    Dim pt As POINTAPI
    GetCursorPos pt
    isMouseOver = (WindowFromPoint(pt.X, pt.Y) = Hwnd)
End Function
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = PicMain.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    PicMain.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Private Sub PicMain_Click()
    RefreshControl
    RaiseEvent Click
End Sub
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = PicMain.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    PicMain.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    PicMain.Refresh
End Sub
Public Property Get Theme() As XPVisualTheme
Attribute Theme.VB_Description = "WIndows Xp Theme"
    Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As XPVisualTheme)
    m_Theme = New_Theme
    PropertyChanged "Theme"
    RefreshControl
End Property
Public Property Get ButtonStyle() As ButtonType
    ButtonStyle = m_ButtonStyle
End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As ButtonType)
    m_ButtonStyle = New_ButtonStyle
    PropertyChanged "ButtonStyle"
    RefreshControl
End Property
Public Property Get Hwnd() As Long
Attribute Hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    Hwnd = PicMain.Hwnd
End Property
Private Sub UserControl_InitProperties()
    m_Theme = m_def_Theme
    m_ButtonStyle = m_def_ButtonStyle
    m_IsActivate = m_def_IsActivate
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    PicMain.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    PicMain.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
    m_ButtonStyle = PropBag.ReadProperty("ButtonStyle", m_def_ButtonStyle)
    PicMain.Enabled = PropBag.ReadProperty("Enabled", True)
    m_IsActivate = PropBag.ReadProperty("IsActivate", m_def_IsActivate)
    RefreshControl
    If Theme = Blue Then
        If Enabled = True Then
            RefreshControl
        Else
            PicMain.Picture = CbMax(4).Picture
        End If
    Else
        If Enabled = True Then
            RefreshControl
        Else
            PicMain.Picture = CbMax(9).Picture
        End If
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoRedraw", PicMain.AutoRedraw, False)
    Call PropBag.WriteProperty("ToolTipText", PicMain.ToolTipText, "")
    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
    Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, m_def_ButtonStyle)
    Call PropBag.WriteProperty("Enabled", PicMain.Enabled, True)
    Call PropBag.WriteProperty("IsActivate", m_IsActivate, m_def_IsActivate)
End Sub

Public Sub RefreshControl(Optional ByVal IpRes As Integer = 1)
    MyButton = 0
    
    If PicMain.Enabled = False Then
        If m_Theme = 0 Then
            PicMain.Picture = CbMax(4).Picture
        ElseIf m_Theme = 1 Then
            PicMain.Picture = CbMax(9).Picture
        Else
            PicMain.Picture = CbMax(10).Picture
        End If
        Exit Sub
    End If
    
    Select Case m_ButtonStyle
        Case MinButton:
            If m_Theme = 0 Then
                If IpRes = 0 Then
                    PicMain.Picture = CbMin(3).Picture
                Else
                    PicMain.Picture = CbMin(0).Picture
                End If
            ElseIf m_Theme = 1 Then
                If IpRes = 0 Then
                    PicMain.Picture = CbMin(7).Picture
                Else
                    PicMain.Picture = CbMin(4).Picture
                End If
            Else
                If IpRes = 0 Then
                    PicMain.Picture = CbMin(8).Picture
                Else
                    PicMain.Picture = CbMin(11).Picture
                End If
            End If
        Case MaxButton:
            If m_Theme = 0 Then
                If IpRes = 0 Then
                    PicMain.Picture = CbMax(3).Picture
                Else
                    PicMain.Picture = CbMax(0).Picture
                End If
            ElseIf m_Theme = 1 Then
                If IpRes = 0 Then
                    PicMain.Picture = CbMax(8).Picture
                Else
                    PicMain.Picture = CbMax(5).Picture
                End If
            Else
                If IpRes = 0 Then
                    PicMain.Picture = CbMax(11).Picture
                Else
                    PicMain.Picture = CbMax(14).Picture
                End If
            End If
       Case CloseButton:
            If m_Theme = 0 Then
                If IpRes = 0 Then
                    PicMain.Picture = CbClose(3).Picture
                Else
                    PicMain.Picture = CbClose(0).Picture
                End If
            ElseIf m_Theme = 1 Then
                If IpRes = 0 Then
                    PicMain.Picture = CbClose(7).Picture
                Else
                    PicMain.Picture = CbClose(4).Picture
                End If
            Else
                If IpRes = 0 Then
                    PicMain.Picture = CbClose(8).Picture
                Else
                    PicMain.Picture = CbClose(11).Picture
                End If
            End If
        Case Else:
            If m_Theme = 0 Then
                If IpRes = 0 Then
                    PicMain.Picture = CbRestore(3).Picture
                Else
                    PicMain.Picture = CbRestore(0).Picture
                End If
            ElseIf m_Theme = 1 Then
                If IpRes = 0 Then
                    PicMain.Picture = CbRestore(7).Picture
                Else
                    PicMain.Picture = CbRestore(4).Picture
                End If
            Else
                If IpRes = 0 Then
                    PicMain.Picture = CbRestore(8).Picture
                Else
                    PicMain.Picture = CbRestore(11).Picture
                End If
            End If
        End Select
End Sub
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = PicMain.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    PicMain.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    RefreshControl
End Property

Private Sub ChangeIfOver(ByVal IpRes As Boolean)
On Error GoTo Err_RF

If MyButton = 1 Then IpRes = False

    If m_Theme = 0 Then
        If m_ButtonStyle = MinButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbMin(2).Picture
            Else
                PicMain.Picture = CbMin(1).Picture
            End If
        ElseIf m_ButtonStyle = MaxButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbMax(2).Picture
            Else
                PicMain.Picture = CbMax(1).Picture
            End If
        ElseIf m_ButtonStyle = CloseButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbClose(2).Picture
            Else
                PicMain.Picture = CbClose(1).Picture
            End If
        Else
            If IpRes = 0 Then
                PicMain.Picture = CbRestore(2).Picture
            Else
                PicMain.Picture = CbRestore(1).Picture
            End If
        End If
    ElseIf m_Theme = 1 Then
        If m_ButtonStyle = MinButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbMin(6).Picture
            Else
                PicMain.Picture = CbMin(5).Picture
            End If
        ElseIf m_ButtonStyle = MaxButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbMax(7).Picture
            Else
                PicMain.Picture = CbMax(6).Picture
            End If
        ElseIf m_ButtonStyle = CloseButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbClose(6).Picture
            Else
                PicMain.Picture = CbClose(5).Picture
            End If
        Else
            If IpRes = 0 Then
                PicMain.Picture = CbRestore(6).Picture
            Else
                PicMain.Picture = CbRestore(5).Picture
            End If
        End If
    Else
        If m_ButtonStyle = MinButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbMin(9).Picture
            Else
                PicMain.Picture = CbMin(10).Picture
            End If
        ElseIf m_ButtonStyle = MaxButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbMax(12).Picture
            Else
                PicMain.Picture = CbMax(13).Picture
            End If
        ElseIf m_ButtonStyle = CloseButton Then
            If IpRes = 0 Then
                PicMain.Picture = CbClose(9).Picture
            Else
                PicMain.Picture = CbClose(10).Picture
            End If
        Else
            If IpRes = 0 Then
                PicMain.Picture = CbRestore(9).Picture
            Else
                PicMain.Picture = CbRestore(10).Picture
            End If
        End If
    End If
Exit Sub
Err_RF:
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get IsActivate() As Boolean
    IsActivate = m_IsActivate
End Property

Public Property Let IsActivate(ByVal New_IsActivate As Boolean)
    m_IsActivate = New_IsActivate
    PropertyChanged "IsActivate"
End Property

