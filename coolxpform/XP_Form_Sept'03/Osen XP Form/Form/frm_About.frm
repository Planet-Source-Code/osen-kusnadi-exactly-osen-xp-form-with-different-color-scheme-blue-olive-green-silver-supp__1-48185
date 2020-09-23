VERSION 5.00
Begin VB.Form Frm_About 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "About Osen XP Form ActiveX Control"
   ClientHeight    =   4290
   ClientLeft      =   2505
   ClientTop       =   1245
   ClientWidth     =   5655
   Icon            =   "frm_About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Osen_XP_Form_Ctl.XPButton Command2 
      Height          =   360
      Left            =   4110
      TabIndex        =   14
      Top             =   3420
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&System Info ..."
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Osen_XP_Form_Ctl.XPButton Command1 
      Height          =   360
      Left            =   4110
      TabIndex        =   13
      Top             =   3000
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
      ForeColor       =   -2147483630
   End
   Begin Osen_XP_Form_Ctl.OsenXPForm OsenXPForm1 
      Height          =   4275
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7541
      Caption         =   "About Osen XP Form ActiveX Control"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoLoad        =   -1  'True
      ShowIcon        =   0   'False
      MenuBackColor   =   16777215
      BackColor       =   -2147483643
   End
   Begin VB.Label LbUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Osen Kusnadi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   660
      TabIndex        =   11
      Top             =   2130
      Width           =   975
   End
   Begin VB.Label LbCompanyName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mecoindo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   660
      TabIndex        =   10
      Top             =   2355
      Width           =   675
   End
   Begin VB.Label LbSerial 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number : OSEN-2210-BEST-XPCTL-XPSUITE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   660
      TabIndex        =   9
      Top             =   2595
      Width           =   3570
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB6"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A28&
      Height          =   270
      Index           =   3
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label LbThird 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3rd Party"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VB6"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This product is licenced to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   5
      Top             =   1890
      Width           =   2235
   End
   Begin VB.Label LbTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Osen XP Form ActiveX Control"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C56A31&
      Height          =   360
      Left            =   1110
      TabIndex        =   4
      Top             =   630
      Width           =   4080
   End
   Begin VB.Label LbEdition 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Professional Edition"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   3525
      TabIndex        =   3
      Top             =   960
      Width           =   1890
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00B99D7F&
      X1              =   285
      X2              =   5460
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label LbNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frm_About.frx":06C4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1365
      Left            =   300
      TabIndex        =   0
      Top             =   2940
      Width           =   3630
   End
   Begin VB.Image ImgPic 
      Height          =   600
      Left            =   300
      Picture         =   "frm_About.frx":07E6
      Top             =   540
      Width           =   600
   End
   Begin VB.Label LbVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.0.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1350
      TabIndex        =   2
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Label LbCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003 Osen Kusnadi, All rights reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1350
      MouseIcon       =   "frm_About.frx":1AE8
      TabIndex        =   1
      ToolTipText     =   "mailto: Osen Kusnadi<okusnadi@cikarang.actaris.com>"
      Top             =   1305
      Width           =   3795
   End
End
Attribute VB_Name = "Frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Dim N As Integer

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub command2_Click()
    StartSysInfo
End Sub

Private Sub Form_Activate()
    
    SetCursorPos ((Screen.Width - Me.Width) / 2) / 15 + Command1.Left / 15 + Command1.Width / 30, ((Screen.Height - Me.Height) / 2) / 15 + Command1.Top / 15 + Command1.Height / 30

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    OsenXPForm1.ColorScheme = M_About_Theme
    LbVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LbCopyright.ForeColor = vbBlue Then
        LbCopyright.ForeColor = vbBlack
        LbCopyright.FontUnderline = False
        LbCopyright.MousePointer = 0
    End If
End Sub

Private Sub LbCopyright_Click()
    ShellExecute Hwnd, "open", "mailto:Osen Kusnadi<okusnadi@cikarang.actaris.com>", vbNullString, vbNullString, SW_SHOW
    LbCopyright.ForeColor = vbRed
End Sub

Private Sub LbCopyright_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LbCopyright.ForeColor <> vbBlue Then
        LbCopyright.ForeColor = vbBlue
        LbCopyright.FontUnderline = True
        LbCopyright.MousePointer = 99
    End If
End Sub

