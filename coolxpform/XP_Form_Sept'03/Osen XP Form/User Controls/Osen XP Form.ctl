VERSION 5.00
Begin VB.UserControl OsenXPForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8E9EC&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   174
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   325
   ToolboxBitmap   =   "Osen XP Form.ctx":0000
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   1170
      ScaleHeight     =   1635
      ScaleWidth      =   2955
      TabIndex        =   9
      Top             =   3270
      Visible         =   0   'False
      Width           =   2955
   End
   Begin Osen_XP_Form_Ctl.ThemeX ThemeX1 
      Left            =   5400
      Top             =   750
      _ExtentX        =   1561
      _ExtentY        =   1667
   End
   Begin Osen_XP_Form_Ctl.ControlButton CloseButton 
      Height          =   315
      Left            =   1860
      TabIndex        =   8
      Top             =   660
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      ButtonStyle     =   2
   End
   Begin Osen_XP_Form_Ctl.ControlButton MaximizeButton 
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   660
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      ButtonStyle     =   1
   End
   Begin Osen_XP_Form_Ctl.ControlButton Minimizebutton 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   660
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
   End
   Begin VB.PictureBox pICmenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   420
      ScaleHeight     =   360
      ScaleWidth      =   3825
      TabIndex        =   3
      Top             =   -1125
      Visible         =   0   'False
      Width           =   3825
      Begin Osen_XP_Form_Ctl.LabelMenu LbMenu 
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   -375
         Visible         =   0   'False
         Width           =   660
         _ExtentX        =   1270
         _ExtentY        =   476
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   -720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   -30
         X2              =   2670
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   2655
         Y1              =   345
         Y2              =   345
      End
   End
   Begin VB.Image picicon 
      Height          =   240
      Left            =   1920
      Picture         =   "Osen XP Form.ctx":0312
      Top             =   -1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label LbAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   195
      Left            =   2100
      TabIndex        =   2
      Top             =   -1020
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Image TitleIcon 
      Height          =   240
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Image BottomLeft 
      Height          =   60
      Left            =   0
      Picture         =   "Osen XP Form.ctx":045C
      Top             =   2520
      Width           =   60
   End
   Begin VB.Image BottomRight 
      Height          =   60
      Left            =   4800
      Picture         =   "Osen XP Form.ctx":0795
      Top             =   2520
      Width           =   60
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   60
      Picture         =   "Osen XP Form.ctx":0ACF
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   4755
   End
   Begin VB.Image Right 
      Height          =   2085
      Left            =   4800
      MousePointer    =   9  'Size W E
      Picture         =   "Osen XP Form.ctx":0DFD
      Stretch         =   -1  'True
      Top             =   450
      Width           =   60
   End
   Begin VB.Image Left 
      Height          =   2085
      Left            =   0
      Picture         =   "Osen XP Form.ctx":112B
      Stretch         =   -1  'True
      Top             =   450
      Width           =   60
   End
   Begin VB.Label Caption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SEN MASTER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   375
      TabIndex        =   0
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label Caption2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SEN MASTER"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   405
      TabIndex        =   1
      Top             =   150
      Width           =   1125
   End
   Begin VB.Image Title 
      Height          =   450
      Left            =   150
      Picture         =   "Osen XP Form.ctx":1459
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image TitleRight 
      Height          =   450
      Left            =   4710
      Picture         =   "Osen XP Form.ctx":1E73
      Top             =   0
      Width           =   150
   End
   Begin VB.Image TitleLeft 
      Height          =   450
      Left            =   0
      Picture         =   "Osen XP Form.ctx":2273
      Top             =   0
      Width           =   150
   End
End
Attribute VB_Name = "OsenXPForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'/**************** Declare API Function **************************************************************************
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum XPTheme
     Blue = 0
     [Olive Green] = 1
     Silver = 2
End Enum

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal Hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvPara As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const WM_NCLBUTTONDOWN         As Long = &HA1
Private Const HTCAPTION                As Long = 2
Private bTransparent                   As Boolean
Private Const MF_BYPOSITION            As Long = &H400&
Private Const MF_BYCOMMAND             As Long = 0
Private Const SC_RESTORE               As Long = &HF120
Private Const SC_MOVE                  As Long = &HF010
Private Const SC_SIZE                  As Long = &HF000
Private Const SC_MINIMIZE              As Long = &HF020
Private Const SC_MAXIMIZE              As Long = &HF030
Private Const SC_CLOSE                 As Long = &HF060
Private Const WM_GETSYSMENU            As Long = &H313
Private Const HWND_TOPMOST             As Long = -1
Private Const HWND_NOTOPMOST           As Long = -2
Private Const SWP_SHOWWINDOW           As Long = &H40

Private MenuCOUNT As Integer

Private Oldcp As POINTAPI ':( Missing Scope
Private Newcp As POINTAPI ':( Missing Scope

Private WithEvents MyForm As Form
Attribute MyForm.VB_VarHelpID = -1

Private Const GWL_STYLE     As Long = (-16)
Private Const WS_SYSMENU    As Long = &H80000
Private m_AutoLoad          As Boolean
Private m_ShowMinimize      As Boolean
Private m_ShowMaximize      As Boolean
Private m_ShowClose         As Boolean
Private m_ShowHelp          As Boolean
Private m_EnableMaximize    As Boolean

Public Event Resize(IsTop As Integer, IsHeight As Integer, IsWidth As Integer) 'MappingInfo=UserControl,UserControl,-1,Resize
Public Event Help()
Public Event CloseForm()

Private Const m_def_AutoLoad        As Boolean = False
Private Const m_def_ShowMinimize    As Boolean = True
Private Const m_def_ShowMaximize    As Boolean = True
Private Const m_def_ShowClose       As Boolean = True
Private Const m_def_ShowHelp        As Boolean = False
Private Const m_def_EnableMaximize  As Boolean = True

Private MyTitleIcon     As Image
Private IsHoverMenu     As Integer
Private IsPressMenu     As Integer
Private MyMainMenu()    As Object
Private IsLoad          As Boolean

Private Const m_def_IpModal As Long = -1
Private m_IpModal As Integer

'Default Property Values:
Const m_def_Theme = 0
Private Const m_def_TitleTop    As Integer = -2
Private Const m_def_IconTop     As Integer = 7
Private Const m_def_IconIndex   As Integer = 0
Private Const m_def_CloseActive As Integer = False

'Property Variables:
Private m_Theme As XPTheme
Private m_TitleTop As Integer
Private m_IconTop As Integer
Private m_IconIndex As Integer
Private m_CloseActive As Boolean
Private m_HaveChild As Boolean

'Event Declarations:
Public Event Click()
Public Event DblClick()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=PicMain,PicMain,-1,MouseMove
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=PicMain,PicMain,-1,MouseDown
Private m_activeform As Integer

Public Sub About()
Attribute About.VB_UserMemId = -552

    M_About_Theme = m_Theme
    Frm_About.Show 1

End Sub

Public Property Get AutoLoad() As Boolean

    AutoLoad = m_AutoLoad

End Property

Public Property Let AutoLoad(ByVal New_AutoLoad As Boolean)

    m_AutoLoad = New_AutoLoad
    PropertyChanged "AutoLoad"

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

    PicMain.BackColor() = New_BackColor
    PropertyChanged "BackColor"

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = PicMain.BackColor

End Property

Private Sub Bottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) Then
        GetCursorPos Oldcp
    End If

End Sub

Private Sub Bottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Z
    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) And (MyForm.WindowState = 0) Then
        Bottom.MousePointer = 7
      Else 'NOT (MYFORM.BORDERSTYLE...
        Bottom.MousePointer = 0
    End If

    If MyForm.WindowState = 2 Then
        TaskBarShow
    End If
Z:

End Sub

Private Sub Bottom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
        If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
        If (MyForm.BorderStyle = 2) Then
            If MyForm.WindowState = 0 Then
                GetCursorPos Newcp
                ResizeForm MyForm, Oldcp, Newcp, 3
                SetStyle MyForm
                If MyForm.BorderStyle <> 0 Then MyForm.Height = MyForm.Height + 375 ':( Expand Structure
                UserControl.Height = MyForm.Height
                UserControl.Width = MyForm.Width
                ReTransObj MyForm
            End If
        End If

End Sub ':( On Error Resume still active

Private Sub BottomLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) Then
        GetCursorPos Oldcp
    End If

End Sub

Private Sub BottomLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Z
    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) And (MyForm.WindowState = 0) Then
        BottomLeft.MousePointer = 6
      Else 'NOT (MYFORM.BORDERSTYLE...
        BottomLeft.MousePointer = 0
    End If
Z:

End Sub

Private Sub BottomLeft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
        If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition

        If (MyForm.BorderStyle = 2) Then

            If MyForm.WindowState = 0 Then

                GetCursorPos Newcp
                ResizeForm MyForm, Oldcp, Newcp, 5
                SetStyle MyForm

                If MyForm.BorderStyle <> 0 Then MyForm.Height = MyForm.Height + 375 ':( Expand Structure

                UserControl.Height = MyForm.Height
                UserControl.Width = MyForm.Width
                ReTransObj MyForm

            End If

        End If

End Sub ':( On Error Resume still active

Private Sub BottomRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) Then
        GetCursorPos Oldcp
    End If

End Sub

Private Sub BottomRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Z
    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) And (MyForm.BorderStyle = 2) And (MyForm.WindowState = 0) Then
        BottomRight.MousePointer = 8
      Else 'NOT (MYFORM.BORDERSTYLE...
        BottomRight.MousePointer = 0
    End If
Z:

End Sub

Private Sub BottomRight_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
        If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
        If (MyForm.BorderStyle = 2) Then
            If MyForm.WindowState = 0 Then
                GetCursorPos Newcp
                ResizeForm MyForm, Oldcp, Newcp, 4
                SetStyle MyForm
                If MyForm.BorderStyle <> 0 Then MyForm.Height = MyForm.Height + 375 ':( Expand Structure
                UserControl.Height = MyForm.Height
                UserControl.Width = MyForm.Width
                ReTransObj MyForm
            End If
        End If

End Sub ':( On Error Resume still active

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."

    Caption = Caption1.Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    Caption1.Caption() = New_Caption
    Caption2.Caption = New_Caption
    UserControl.Parent.Caption = New_Caption
    PropertyChanged "Caption"

End Property

Private Sub Caption1_Change()

    Caption2.Caption = Caption1.Caption

End Sub

Private Sub Caption1_DblClick()

    Title_DblClick

End Sub

Private Sub Caption1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Lets user move parent form

    Call ReleaseCapture
    Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub

Private Sub Caption2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Lets user move parent form

    Call ReleaseCapture
    Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub

Public Property Get CloseActive() As Boolean

    CloseActive = m_CloseActive

End Property

Public Property Let CloseActive(ByVal New_CloseActive As Boolean)

    m_CloseActive = New_CloseActive
    PropertyChanged "CloseActive"

End Property

Private Sub CbMin_Click(Index As Integer)

End Sub

Private Sub CloseButton_Click()

    On Error GoTo EF
    If CloseActive Then
        RaiseEvent CloseForm
      Else 'CLOSEACTIVE = FALSE/0
        If Not MyForm Is Nothing Then Unload MyForm ':( Expand Structure
    End If
EF:

End Sub

Public Sub ContainerCheck()

    On Error GoTo hjk
  Dim Control As Object ':( Move line to top of current Sub
    For Each Control In UserControl.Parent
        If Control.Container.Hwnd = UserControl.ContainerHwnd Then
            Control.Left = Control.Left + 75
            Control.Top = Control.Top + 450
        End If
    Next Control
hjk:

End Sub

Public Function DefaultBackgroundColor() As String

    DefaultBackgroundColor = &HD8E9EC   '&HEAF1F1   'Returns a common off-white Windows XP color

End Function

Public Property Let Enabled(ByVal New_Enabled As Boolean)
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

    PicMain.Enabled() = New_Enabled
    PropertyChanged "Enabled"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PicMain,PicMain,-1,Enabled
Public Property Get Enabled() As Boolean

    Enabled = PicMain.Enabled

End Property

Public Property Get EnableMaximize() As Boolean

    EnableMaximize = m_EnableMaximize

End Property

Public Property Let EnableMaximize(ByVal New_EnableMaximize As Boolean)

    m_EnableMaximize = New_EnableMaximize
    MaximizeButton.Enabled = m_EnableMaximize
    PropertyChanged "EnableMaximize"
    
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512

    Set Font = Caption1.Font

End Property

Public Property Set Font(ByVal New_Font As Font)

    Set Caption1.Font = New_Font
    Set Caption2.Font = New_Font
    PropertyChanged "Font"

End Property

Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)

    On Error Resume Next
      Dim wFlags As Long, placement As Long ':( Move line to top of current Sub
        wFlags = &H2 Or &H1 Or &H40 Or &H10
        Select Case bTopMost
          Case True
            placement = -1
          Case False
            placement = -2
        End Select
        SetWindowPos hWindow, placement, 0, 0, 0, 0, wFlags

End Sub ':( On Error Resume still active

Public Function GetCompName() As String

  Dim Commstr As String, nErr As Long

    Commstr = Space$(255)
    nErr = GetComputerName(Commstr, 255)
    GetCompName = Commstr

End Function

Public Property Let HaveChild(ByVal NewHaveChild As Boolean)

    m_HaveChild = NewHaveChild
    PropertyChanged "HaveChild"

End Property

Public Property Get HaveChild() As Boolean

    HaveChild = m_HaveChild

End Property

Public Property Get Hwnd() As Long
Attribute Hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."

    Hwnd = PicMain.Hwnd

End Property

Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Returns/sets a graphic to be displayed in a control."

    Set Icon = TitleIcon.Picture

End Property

Public Property Set Icon(ByVal New_Icon As Picture)

    On Error GoTo Z
    Set TitleIcon.Picture = New_Icon
    Set UserControl.Parent.Icon = TitleIcon.Picture
    If Not New_Icon Is Nothing Then
        ShowIcon = True
    Else
        ShowIcon = False
    End If
    PropertyChanged "Icon"
Z:

End Property

Public Property Let IconTop(ByVal New_IconTop As Integer)

    m_IconTop = New_IconTop
    TitleIcon.Top = New_IconTop
    PropertyChanged "IconTop"

End Property

Public Property Get IconTop() As Integer

    IconTop = m_IconTop
    TitleIcon.Top = IconTop

End Property

Public Property Get IpModal() As Integer

    IpModal = m_IpModal

End Property

Public Property Let IpModal(ByVal New_IpModal As Integer)

    m_IpModal = New_IpModal
    PropertyChanged "IpModal"

End Property

Private Sub LbMenu_Click(Index As Integer)
  Dim Xok As Long, Yok As Long, J As Integer ':( Move line to top of current Sub
  On Error GoTo Z
    
        Xok = LbMenu(Index).Left + pICmenu.Left + 60
        Yok = LbMenu(Index).Top + pICmenu.Top + 450 + LbMenu(Index).Height - 80
        MyForm.PopupMenu MyMainMenu(Index), , Xok, Yok
Z:

End Sub

Private Sub LbMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Xok As Long, Yok As Long, J As Integer ':( Move line to top of current Sub
  On Error GoTo Z
    
    If Button = 1 Then
        Xok = LbMenu(Index).Left + pICmenu.Left + 60
        Yok = LbMenu(Index).Top + pICmenu.Top + 480 + LbMenu(Index).Height - 80
        MyForm.PopupMenu MyMainMenu(Index), , Xok, Yok
    End If
Z:

End Sub

Private Sub Left_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) Then
        GetCursorPos Oldcp
    End If

End Sub

Private Sub Left_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Z
    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) And (MyForm.BorderStyle = 2) And (MyForm.WindowState = 0) Then
        Left.MousePointer = 9
      Else 'NOT (MYFORM.BORDERSTYLE...
        Left.MousePointer = 0
    End If
Z:

End Sub

Private Sub Left_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next

        If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
        If (MyForm.BorderStyle = 2) Then
            If MyForm.WindowState = 0 Then
                GetCursorPos Newcp
                ResizeForm MyForm, Oldcp, Newcp, 0
                SetStyle MyForm
                If MyForm.BorderStyle <> 0 Then MyForm.Height = MyForm.Height + 375 ':( Expand Structure
                UserControl.Height = MyForm.Height
                UserControl.Width = MyForm.Width
                ReTransObj MyForm
            End If
        End If

End Sub ':( On Error Resume still active

Public Sub LoadXP(Optional ByVal OptModal As Integer = 0, Optional ByVal OwnForm As Object)

    On Error GoTo Z

  Dim IpForm As Object ':( Move line to top of current Sub
  Dim XP_Name As Object ':( Move line to top of current Sub
  Dim oCtl As Control ':( Move line to top of current Sub
  Dim i As Integer ':( Move line to top of current Sub

    Set IpForm = UserControl.Parent
    Set MyForm = IpForm
    IsLoad = True

    i = 0
    SetCursorPos 9000, 9000
    '/******* Hidden Object in procccess **********************
    For Each oCtl In MyForm
        If TypeOf oCtl Is OsenXPForm Then
            oCtl.Top = 0
            oCtl.Left = 0
            Set XP_Name = oCtl
          ElseIf TypeOf oCtl Is Menu Then 'NOT TYPEOF...
            If UCase$(oCtl.Tag) = "MAINMENU" Then
                pICmenu.Visible = True
                i = i + 1
                Load LbMenu(i)
                ReDim Preserve MyMainMenu(i)
                Set MyMainMenu(i) = oCtl
                
                With LbMenu(i)
                
                    .Caption = oCtl.Caption
                    .BackColor = pICmenu.BackColor
                    .Visible = True
                    .Enabled = oCtl.Enabled
                    .Top = 45
                    
                    If i = 1 Then
                        .Left = 60
                      Else 'NOT I...
                        .Left = LbMenu(i - 1).Width + LbMenu(i - 1).Left
                    End If
                
                End With 'LBMENU(I)
                oCtl.Visible = False
                MenuCOUNT = i
            End If
        End If
    Next oCtl
    If i <> 0 Then Load LbMenu(100) ':( Expand Structure
    '/Setting size
    IpForm.Width = XP_Name.Width
    IpForm.Height = XP_Name.Height

    If IpForm.BorderStyle <> 0 Then
        IpForm.Height = XP_Name.Height + 375
    End If

    If IpForm.BorderStyle <> 2 Then
        ShowMaximize = False
        ShowMinimize = False
    End If

    If Not MyForm.MaxButton Then EnableMaximize = False ':( Expand Structure

    If IpForm.BorderStyle = 1 Then

        If IpForm.MinButton Then

            ShowMinimize = True
            ShowMaximize = True
            EnableMaximize = False

        End If

    End If

    PicMain.Visible = HaveChild

    '*************** Set FORM Style *************************
    If IpForm.BorderStyle <> 0 Then SetStyle IpForm ':( Expand Structure

    XP_Name.Width = IpForm.Width
    XP_Name.Height = IpForm.Height

    '/**************** Set Transparant ************************
    ReTransObj IpForm
    SetCursorPos (MyForm.Left + MyForm.Width) / 15, (MyForm.Top / 15)
    DoEvents

    If Not OwnForm Is Nothing Then
        IpForm.Visible = False
        DoEvents
        IpForm.Show OptModal, OwnForm
        Exit Sub '>---> Bottom
      Else 'NOT NOT...
        If OptModal = 1 Or IpModal = 1 Then
            IpForm.Visible = False
            DoEvents
            IpForm.Show 1
        End If
    End If

Exit Sub

Z:
    MyForm.Show

End Sub

Public Sub MaxBtnClick()

    MaximizeButton_Click

End Sub

Private Sub MaximizeButton_Click()

    On Error GoTo xc
    If Not MyForm Is Nothing Then
        If MyForm.WindowState = 0 Then
            MyForm.WindowState = 2
          Else 'NOT MYFORM.WINDOWSTATE...
            MyForm.WindowState = 0
        End If
        DoEvents
        UserControl.Width = MyForm.Width
        UserControl.Height = MyForm.Height
        ReTransObj MyForm
        Repos
    End If
xc:
End Sub
Public Property Get MenuBackColor() As OLE_COLOR
Attribute MenuBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

    MenuBackColor = pICmenu.BackColor

End Property

Public Property Let MenuBackColor(ByVal New_MenuBackColor As OLE_COLOR)

    pICmenu.BackColor() = New_MenuBackColor
    PropertyChanged "MenuBackColor"

End Property

Private Sub MinimizeButton_Click()
    If Not MyForm Is Nothing Then
        MyForm.WindowState = 1
    End If
End Sub

Private Sub MYFORM_Activate()

    SetFormActiveStyle True

End Sub

Private Sub MYFORM_Deactivate()

    SetFormActiveStyle False

End Sub

Private Sub MyForm_Resize()
    If MyForm.WindowState = 2 Then
        MaximizeButton.ButtonStyle = RestoreButton
    Else
        MaximizeButton.ButtonStyle = MaxButton
    End If
End Sub

Private Sub PicMain_Click()

    RaiseEvent Click

End Sub

Private Sub PicMain_DblClick()

    RaiseEvent DblClick

End Sub

Private Sub PicMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub PicMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)


End Sub

Private Sub pICmenu_Resize()

    Line1.X1 = 0
    Line2.X1 = 0
    Line1.X2 = pICmenu.Width * 15
    Line2.X2 = pICmenu.Width * 15
    Line1.Refresh
    Line2.Refresh

End Sub

Public Property Set Picture(ByVal New_Picture As Picture)
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."

    Set PicMain.Picture = New_Picture
    PropertyChanged "Picture"

End Property

Public Property Get Picture() As Picture

    Set Picture = PicMain.Picture

End Property

Public Sub Repos()

  'This repositions the different controls on the form when it is resized
  Dim X As Single ':( Move line to top of current Sub
  Dim Y As Single ':( Move line to top of current Sub

    If UserControl.Height < 615 Then UserControl.Height = 615   'Checks that form':( Expand Structure
    If UserControl.Width < 1695 Then UserControl.Width = 1695   'is not too small':( Expand Structure

    X = UserControl.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = UserControl.Height / Screen.TwipsPerPixelY  'form in pixels

    'Titlebar
    With TitleLeft
        .Left = 0
        .Top = 0
        .Height = 30
    End With 'TITLELEFT

    With Title
        .Height = 30
        .Left = TitleLeft.Width
        .Top = 0
        .Width = X - TitleLeft.Width - TitleRight.Width
        pICmenu.Top = .Top + .Height
    End With 'TITLE

    With TitleRight
        .Left = Title.Left + Title.Width
        .Top = 0
        .Height = 30
    End With 'TITLERIGHT

    'Borders
    With BottomLeft
        .Left = 0
        .Height = 4
        .Top = Y - .Height
        .Width = 4
    End With 'BOTTOMLEFT

    With BottomRight
        .Height = 4
        .Width = 4
        .Left = X - .Width
        .Top = Y - .Height
    End With 'BOTTOMRIGHT

    With Left
        .Left = 0
        .Width = 4
        .Top = TitleLeft.Top + TitleLeft.Height
        .Height = BottomLeft.Top - .Top
    End With 'LEFT

    With Right
        .Width = 4
        .Left = X - .Width
        .Top = TitleRight.Top + TitleRight.Height
        .Height = BottomRight.Top - .Top
    End With 'RIGHT

    With Bottom
        .Height = 4
        .Left = BottomLeft.Width
        .Top = Y - Bottom.Height
        .Width = X - BottomLeft.Width - BottomRight.Width
        pICmenu.Width = .Width
        pICmenu.Left = .Left
    End With 'BOTTOM

    'Buttons
    With CloseButton
        .Left = Right.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With 'CLOSEBUTTON

    With MaximizeButton
        .Left = CloseButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With 'MAXIMIZEBUTTON


    With Minimizebutton
        .Left = MaximizeButton.Left - .Width - 2
        .Top = (Title.Height - .Height) / 2
    End With 'MINIMIZEBUTTON

    'Icon
    With TitleIcon
        .Left = Left.Left + Left.Width + 2
        .Top = IconTop '(Title.Height - .Height) / 2
    End With 'TITLEICON

    'Titlebar Caption
    With Caption1
        If TitleIcon.Visible = True Then ':( Remove Pleonasm
            .Left = TitleIcon.Left + TitleIcon.Width + 3
          Else 'NOT TITLEICON.VISIBLE...
            .Left = Left.Left + Left.Width + 2.5
        End If
        .Top = (((Title.Height - 13) / 2) - 9) + Caption1.FontSize + m_TitleTop
        .Width = Minimizebutton.Left - TitleIcon.Left - TitleIcon.Width - 10
        If Minimizebutton.Visible = False Then
            .Width = MaximizeButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        End If
        If Minimizebutton.Visible = False And TitleIcon.Visible = False Then
            .Width = MaximizeButton.Left - Left.Left - Left.Width - 10
        End If
        If Minimizebutton.Visible = False And MaximizeButton.Visible = False Then
            .Width = CloseButton.Left - TitleIcon.Left - TitleIcon.Width - 10
        End If
        If Minimizebutton.Visible = False And MaximizeButton.Visible = False And TitleIcon.Visible = False Then
            .Width = CloseButton.Left - Left.Left - Left.Width - 10
        End If
        .AutoSize = True
    End With 'CAPTION1

    With Caption2
        .Left = Caption1.Left - 1
        .Top = Caption1.Top - 1
        .Width = Caption1.Width
        .Caption = Caption1.Caption
        .Width = Caption1.Width
    End With 'CAPTION2

    'Checks if it should have transparent corners
    If bTransparent = True Then ':( Remove Pleonasm
        ReTrans
    End If
    
    Minimizebutton.RefreshControl
    MaximizeButton.RefreshControl
    CloseButton.RefreshControl
    
    SetFormActiveStyle True
    
    If pICmenu.Visible Then

        RaiseEvent Resize(450 + (pICmenu.Height * 15), UserControl.Height - 510 - (pICmenu.Height * 15), UserControl.Width - 120)
        PicMain.Top = pICmenu.Height + 30

      Else 'PICMENU.VISIBLE = FALSE/0

        RaiseEvent Resize(450, UserControl.Height - 510, UserControl.Width - 120)
        PicMain.Top = 30

    End If

End Sub

Private Sub ResizeForm(frm As Form, Oldcp As POINTAPI, Newcp As POINTAPI, ResizeMode As Integer)

    On Error Resume Next
      Dim DifferenceX ':( As Variant ?':( Move line to top of current Sub
      Dim DifferenceY ':( As Variant ?':( Move line to top of current Sub
        DifferenceX = (Newcp.X - Oldcp.X) * Screen.TwipsPerPixelX
        DifferenceY = (Newcp.Y - Oldcp.Y) * Screen.TwipsPerPixelY
        Select Case ResizeMode
          Case 0
            frm.Move frm.Left + DifferenceX, frm.Top, frm.Width - DifferenceX, frm.Height
          Case 1
            frm.Move frm.Left, frm.Top, frm.Width + DifferenceX, frm.Height
          Case 2
            frm.Move frm.Left, frm.Top + DifferenceY, frm.Width, frm.Height - DifferenceY
          Case 3
            frm.Move frm.Left, frm.Top, frm.Width, frm.Height + DifferenceY
          Case 4
            frm.Move frm.Left, frm.Top, frm.Width + DifferenceX, frm.Height + DifferenceY
          Case 5
            frm.Move frm.Left + DifferenceX, frm.Top, frm.Width - DifferenceX, frm.Height + DifferenceY
          Case 6
            frm.Move frm.Left, frm.Top + DifferenceY, frm.Width + DifferenceX, frm.Height - DifferenceY
          Case 7
            frm.Move frm.Left + DifferenceX, frm.Top + DifferenceY, frm.Width - DifferenceX, frm.Height - DifferenceY
        End Select

End Sub ':( On Error Resume still active

Private Sub ReTrans()

  Dim Add As Long
  Dim Sum As Long

  Dim X As Single
  Dim Y As Single

    If UserControl.Height < 615 Then UserControl.Height = 615   'Checks that form':( Expand Structure
    If UserControl.Width < 1695 Then UserControl.Width = 1695   'is not too small':( Expand Structure

    X = UserControl.Width / Screen.TwipsPerPixelX   'Registers the size of the
    Y = UserControl.Height / Screen.TwipsPerPixelY  'form in pixels

    Sum = CreateRectRgn(5, 0, X - 5, 1)
    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
    SetWindowRgn UserControl.ContainerHwnd, Sum, True   'Sets corners transparent

End Sub

Public Sub ReTransObj(IpObject As Object)

    On Error Resume Next
      Dim Add As Long ':( Move line to top of current Sub
      Dim Sum As Long ':( Move line to top of current Sub
      Dim X As Single ':( Move line to top of current Sub
      Dim Y As Single ':( Move line to top of current Sub
        If IpObject.Height < 615 Then IpObject.Height = 615   'Checks that form':( Expand Structure
        If IpObject.Width < 1695 Then IpObject.Width = 1695   'is not too small':( Expand Structure
        X = IpObject.Width / Screen.TwipsPerPixelX   'Registers the size of the
        Y = IpObject.Height / Screen.TwipsPerPixelY  'form in pixels
        Sum = CreateRectRgn(5, 0, X - 5, 1)
        CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
        CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
        CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
        CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
        CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
        SetWindowRgn IpObject.Hwnd, Sum, True   'Sets corners transparent

End Sub ':( On Error Resume still active

Private Sub Right_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
    If (MyForm.BorderStyle = 2) Then
        GetCursorPos Oldcp
    End If

End Sub

Private Sub Right_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo Z
    If Not MyForm Is Nothing Then
        If (MyForm.BorderStyle = 2) And (MyForm.BorderStyle = 2) And (MyForm.WindowState = 0) Then
            Right.MousePointer = 9
          Else 'NOT (MYFORM.BORDERSTYLE...
            Right.MousePointer = 0
        End If
    End If
Z:

End Sub

Private Sub Right_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
        If MyForm Is Nothing Then Exit Sub ':( Expand Structure or consider reversing Condition
        If (MyForm.BorderStyle = 2) Then
            If MyForm.WindowState = 0 Then
                GetCursorPos Newcp
                ResizeForm MyForm, Oldcp, Newcp, 1
                SetStyle MyForm
                If MyForm.BorderStyle <> 0 Then MyForm.Height = MyForm.Height + 375 ':( Expand Structure
                UserControl.Height = MyForm.Height
                UserControl.Width = MyForm.Width
                ReTransObj MyForm
            End If
        End If

End Sub

Public Sub SetMDIPosition(ByVal iLeft As Long, ByVal iWidth As Long, ByVal IHeight As Long)

    With PicMain
        .Left = iLeft
        .Height = IHeight
        .Width = iWidth
    End With 'PICMAIN

End Sub
Public Sub SetStyle(ByVal IpForm As Object)

    On Error Resume Next
      Dim lCurrentSettings As Long ':( Move line to top of current Sub
      Const WS_MINIMIZEBOX = &H20000 ':( Move line to top of current Sub
      Const WS_MAXIMIZEBOX = &H10000 ':( Move line to top of current Sub
      Const WS_THICKFRAME = &H40000 ':( Move line to top of current Sub
      Const WS_DLGFRAME = &H400000 ':( Move line to top of current Sub
      Const WS_CAPTION = &HC00000 ':( Move line to top of current Sub
        lCurrentSettings = GetWindowLong(IpForm.Hwnd, GWL_STYLE)
        lCurrentSettings = lCurrentSettings And Not WS_THICKFRAME
        lCurrentSettings = lCurrentSettings And Not WS_DLGFRAME
        lCurrentSettings = lCurrentSettings And Not WS_CAPTION
        lCurrentSettings = lCurrentSettings And Not WS_MINIMIZEBOX
        lCurrentSettings = lCurrentSettings And Not WS_MAXIMIZEBOX
        lCurrentSettings = lCurrentSettings Or WS_SYSMENU
        SetWindowLong IpForm.Hwnd, GWL_STYLE, lCurrentSettings
        SetWindowPos IpForm.Hwnd, 0, IpForm.Left / 15, IpForm.Top / 15, (IpForm.Width / 15), (IpForm.Height / 15), &H40
        If IpForm.BorderStyle <> 0 Then
            IpForm.Height = IpForm.Height - 365
        End If

End Sub ':( On Error Resume still active

Public Sub ShowChild(ByVal IpForm As Object)
    
    IpForm.Show 0, MyForm
    SetParent IpForm.Hwnd, PicMain.Hwnd
    m_activeform = m_activeform + 1
    
End Sub
Public Function NChildForm() As Integer
    NChildForm = m_activeform
End Function
Public Sub CloseChildForm()
    If m_activeform > 0 Then
        m_activeform = m_activeform - 1
    End If
End Sub
Public Property Let ShowClose(ByVal New_ShowClose As Boolean)

    m_ShowClose = New_ShowClose
    CloseButton.Visible = m_ShowClose
    PropertyChanged "ShowClose"

End Property

Public Property Get ShowClose() As Boolean

    ShowClose = m_ShowClose

End Property
Public Property Get ShowIcon() As Boolean
Attribute ShowIcon.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

    ShowIcon = TitleIcon.Visible

End Property

Public Property Let ShowIcon(ByVal New_ShowIcon As Boolean)

    TitleIcon.Visible = New_ShowIcon
    Repos
    PropertyChanged "ShowIcon"

End Property

Public Property Get ShowMaximize() As Boolean

    ShowMaximize = m_ShowMaximize

End Property

Public Property Let ShowMaximize(ByVal New_ShowMaximize As Boolean)

    m_ShowMaximize = New_ShowMaximize
    MaximizeButton.Visible = m_ShowMaximize
    PropertyChanged "ShowMaximize"

End Property

Public Property Let ShowMinimize(ByVal New_ShowMinimize As Boolean)

    m_ShowMinimize = New_ShowMinimize
    Minimizebutton.Visible = m_ShowMinimize
    PropertyChanged "ShowMinimize"

End Property

Public Property Get ShowMinimize() As Boolean

    ShowMinimize = m_ShowMinimize

End Property

Public Sub TaskBarShow()

  Dim rtn As Long

    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, &H40)

End Sub

Private Sub Title_DblClick()

    On Error Resume Next
        If Not MyForm Is Nothing Then
            If MyForm.BorderStyle = 2 Then

                If (EnableMaximize And MyForm.MaxButton) Then

                    If MyForm.WindowState = 0 Then
                        MyForm.WindowState = 2
                      Else 'NOT MYFORM.WINDOWSTATE...
                        MyForm.WindowState = 0
                    End If

                    UserControl.Width = MyForm.Width
                    UserControl.Height = MyForm.Height

                    ReTransObj MyForm
                    Repos
                    Minimizebutton.RefreshControl
                    MaximizeButton.RefreshControl
                    CloseButton.RefreshControl

                End If

            End If
        End If

End Sub ':( On Error Resume still active

Private Sub Title_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call ReleaseCapture
    Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub


Private Sub TitleIcon_DblClick()

    On Error Resume Next
        If Not MyForm Is Nothing Then
            If MyForm.BorderStyle = 2 Then
                If MyForm.WindowState = 0 Then
                    MyForm.WindowState = 2
                  Else 'NOT MYFORM.WINDOWSTATE...
                    MyForm.WindowState = 0
                End If
                UserControl.Width = MyForm.Width
                UserControl.Height = MyForm.Height
                ReTransObj MyForm
            End If
        End If

End Sub ':( On Error Resume still active

Private Sub TitleIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call ReleaseCapture
    Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub

Private Sub TitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Lets user move parent form

    Call ReleaseCapture
    Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub

Private Sub TitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Lets user move parent form

    Call ReleaseCapture
    Call SendMessage(UserControl.ContainerHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub
Public Property Get TitleTop() As Integer

    TitleTop = m_TitleTop

End Property

Public Property Let TitleTop(ByVal New_TitleTop As Integer)

    m_TitleTop = New_TitleTop
    PropertyChanged "TitleTop"
    Repos

End Property

Public Sub TransparentEdges()

    bTransparent = True
    Repos

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next
        bTransparent = False  'So we do not set the corners transparent while still in design mode
        IsLoad = False
        Repos   'Reposition

End Sub ':( On Error Resume still active

Private Sub UserControl_InitProperties()

    On Error GoTo Z
    UserControl.Parent.BackColor = DefaultBackgroundColor
    m_ShowMinimize = m_def_ShowMinimize
    m_ShowMaximize = m_def_ShowMaximize
    m_ShowClose = m_def_ShowClose
    m_ShowHelp = m_def_ShowHelp
    m_EnableMaximize = m_def_EnableMaximize
    m_AutoLoad = True
    Caption1.Caption = "Hello " & GetCompName
    UserControl.Parent.Caption = "Osen Kusnadi<osen_kusnadi@yahoo.com>"
    m_IpModal = m_def_IpModal
    m_CloseActive = m_def_CloseActive
    m_IconIndex = m_def_IconIndex
    m_IconTop = m_def_IconTop
    m_TitleTop = m_def_TitleTop
    ShowIcon = False
    Minimizebutton.ButtonStyle = MinButton
    MaximizeButton.ButtonStyle = MaxButton
    CloseButton.ButtonStyle = CloseButton
    
Z:

    m_Theme = m_def_Theme
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Caption1.Caption = PropBag.ReadProperty("Caption", "SEN MASTER")
    Caption2.Caption = Caption1.Caption
    Set Caption1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Caption2.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ShowMinimize = PropBag.ReadProperty("ShowMinimize", m_def_ShowMinimize)
    m_ShowMaximize = PropBag.ReadProperty("ShowMaximize", m_def_ShowMaximize)
    m_ShowClose = PropBag.ReadProperty("ShowClose", m_def_ShowClose)
    m_ShowHelp = PropBag.ReadProperty("ShowHelp", m_def_ShowHelp)
    m_EnableMaximize = PropBag.ReadProperty("EnableMaximize", m_def_EnableMaximize)
    Set TitleIcon.Picture = PropBag.ReadProperty("Icon", Nothing)
    m_AutoLoad = PropBag.ReadProperty("AutoLoad", m_def_AutoLoad)
    TitleIcon.Visible = PropBag.ReadProperty("ShowIcon", True)
    m_IpModal = PropBag.ReadProperty("IpModal", m_def_IpModal)
    pICmenu.BackColor = PropBag.ReadProperty("MenuBackColor", &H80000004)
    m_CloseActive = PropBag.ReadProperty("CloseActive", m_def_CloseActive)
    m_IconTop = PropBag.ReadProperty("IconTop", m_def_IconTop)
    TitleIcon.Top = m_IconTop
    m_TitleTop = PropBag.ReadProperty("TitleTop", m_def_TitleTop)
    m_HaveChild = PropBag.ReadProperty("HaveChild", False)
    PicMain.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    PicMain.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Theme = PropBag.ReadProperty("Theme", m_def_Theme)
    MaximizeButton.Enabled = m_EnableMaximize
    Repos
End Sub

Private Sub UserControl_Resize()

    Repos   'Reposition

End Sub

Private Sub UserControl_Show()

    On Error GoTo Z
    Repos
    If AutoLoad Then LoadXP ':( Expand Structure
Z:

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Caption", Caption1.Caption, "SEN MASTER")
        Call .WriteProperty("Font", Caption1.Font, Ambient.Font)
        Call .WriteProperty("ShowMinimize", m_ShowMinimize, m_def_ShowMinimize)
        Call .WriteProperty("ShowMaximize", m_ShowMaximize, m_def_ShowMaximize)
        Call .WriteProperty("ShowClose", m_ShowClose, m_def_ShowClose)
        Call .WriteProperty("ShowHelp", m_ShowHelp, m_def_ShowHelp)
        Call .WriteProperty("EnableMaximize", m_EnableMaximize, m_def_EnableMaximize)
        Call .WriteProperty("AutoLoad", m_AutoLoad, m_def_AutoLoad)
        Call .WriteProperty("ShowIcon", TitleIcon.Visible, True)
        Call .WriteProperty("Icon", TitleIcon.Picture, Nothing)
        Call .WriteProperty("IpModal", m_IpModal, m_def_IpModal)
        Call .WriteProperty("MenuBackColor", pICmenu.BackColor, &H80000004)
        Call .WriteProperty("CloseActive", m_CloseActive, m_def_CloseActive)
        Call .WriteProperty("IconTop", m_IconTop, m_def_IconTop)
        Call .WriteProperty("TitleTop", m_TitleTop, m_def_TitleTop)
        Call .WriteProperty("HaveChild", m_HaveChild, False)
        TitleIcon.Top = m_IconTop
    End With 'PROPBAG
    Call PropBag.WriteProperty("BackColor", PicMain.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Enabled", PicMain.Enabled, True)

    Call PropBag.WriteProperty("Theme", m_Theme, m_def_Theme)
End Sub
Public Property Get ColorScheme() As XPTheme
    ColorScheme = m_Theme
End Property

Public Property Let ColorScheme(ByVal New_Theme As XPTheme)
    m_Theme = New_Theme
    PropertyChanged "Theme"
    DoEvents
    Repos
    Repos
End Property

Public Sub SetFormActiveStyle(ByVal ActiveForm As Boolean)
    Dim Index As Integer
    
    Index = m_Theme
    ChangeTitleFontcolor ActiveForm
    
    ThemeX1.ChangeTheme Left, TitleLeft, _
    Title, TitleRight, Right, Bottom, BottomLeft, BottomRight, Index, ActiveForm
    
    Minimizebutton.IsActivate = ActiveForm
    MaximizeButton.IsActivate = ActiveForm
    CloseButton.IsActivate = ActiveForm
    Minimizebutton.Theme = m_Theme
    MaximizeButton.Theme = m_Theme
    CloseButton.Theme = m_Theme
    Minimizebutton.RefreshControl ActiveForm
    CloseButton.RefreshControl ActiveForm
    
    If MaximizeButton.Enabled Then
        MaximizeButton.RefreshControl ActiveForm
    End If
    
End Sub

Private Sub ChangeTitleFontcolor(Optional IsActive As Boolean)
Caption2.Visible = IsActive
    If m_Theme <> 2 Then
        If IsActive Then
            Caption1.ForeColor = vbWhite
            Caption2.ForeColor = IIf(m_Theme, &H4000&, &H400000)       '&H8000&
        Else
            Caption1.ForeColor = vbWhite
            Caption2.ForeColor = vbWhite
        End If
    Else
        If IsActive Then
            Caption1.ForeColor = &H0&
            Caption2.ForeColor = &HE0E0E0
        Else
            Caption1.ForeColor = &HC0C0C0
            Caption2.ForeColor = &HE0E0E0
        End If
    End If
End Sub







