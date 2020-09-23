VERSION 5.00
Object = "*\A..\Osen XP Form\Osen Xp Form.vbp"
Begin VB.Form Form2 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   0  'None
   Caption         =   "Please Vote Me ..."
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   2100
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Osen_XP_Form_Ctl.XPButton XPButton1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   1500
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OK"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Osen_XP_Form_Ctl.OsenXPForm OsenXPForm1 
      Height          =   2085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   3678
      Caption         =   "Please Vote Me ..."
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please ! Please ! Please Give me your Vote , i didn't get much from this Application So i have left further development ...... "
      Height          =   645
      Left            =   1350
      TabIndex        =   1
      Top             =   660
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "Form2.frx":0000
      Top             =   570
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Const SW_NORMAL = 1

Public Sub OpenWebsite(strWebsite As String)
    If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, SW_NORMAL) < 33 Then
        ' Insert Error handling code here
    End If
End Sub

Private Sub Form_Activate()
Dim X As Long, Y As Long
    X = (XPButton1.Left / 15) + (XPButton1.Width / 30)
    X = ((Screen.Width - Me.Width) / 30) + X
    Y = (XPButton1.Top / 15) + (XPButton1.Height / 30)
    Y = ((Screen.Height - Me.Height) / 30) + Y
    SetCursorPos X, Y
End Sub

Private Sub Form_Load()
    OsenXPForm1.ColorScheme = Form1.OsenXPForm1.ColorScheme
    If OsenXPForm1.ColorScheme <> 2 Then
        BackColor = &HD8E9EC
    Else
        BackColor = &HE0E0E0
    End If
End Sub

Private Sub XPButton1_Click()
    Call OpenWebsite("http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=48185&optCodeRatingValue=5")
    Unload Me
End Sub
