VERSION 5.00
Object = "*\A..\Osen XP Form\Osen Xp Form.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Cool Alternative XP Form (Skin ActiveX Controls)"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin Osen_XP_Form_Ctl.XPButton XPButton1 
      Height          =   495
      Left            =   2100
      TabIndex        =   1
      Top             =   4440
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Download Other Controls"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin Osen_XP_Form_Ctl.OsenXPForm OsenXPForm1 
      Height          =   5385
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6105
      _ExtentX        =   10663
      _ExtentY        =   9499
      Caption         =   "Cool Alternative XP Form (Skin ActiveX Controls)"
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
      Icon            =   "Form1.frx":058A
      MenuBackColor   =   16777215
      BackColor       =   -2147483643
   End
   Begin VB.Menu mnuColor 
      Caption         =   "Color Scheme"
      Tag             =   "mainmenu"
      Begin VB.Menu mnuScheme 
         Caption         =   "Blue (Default)"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuScheme 
         Caption         =   "Olive Green"
         Index           =   2
      End
      Begin VB.Menu mnuScheme 
         Caption         =   "Silver"
         Index           =   3
      End
      Begin VB.Menu mnu_separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "&New ..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Tag             =   "mainmenu"
      Begin VB.Menu mnuVote 
         Caption         =   "&Vote Me"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About XP Form"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Purpose : Show About dialog
Private Sub mnuAbout_Click()
    OsenXPForm1.About
End Sub

' Purpose : Create New Form1
Private Sub mnuNew_Click()
    Dim MyForm As New Form1
    MyForm.Show 0
End Sub

' Purpose : Change color scheme
Private Sub mnuScheme_Click(Index As Integer)
    OsenXPForm1.ColorScheme = (Index - 1)
    If Index <> 3 Then
        BackColor = &HD8E9EC
    Else
        BackColor = &HE3DFE0
    End If
    ChangeCheck Index
End Sub

' Purpose : Set Checked Active
Private Sub ChangeCheck(ByVal Index As Integer)
    Dim I As Integer
    For I = 1 To 3
        mnuScheme(I).Checked = False
    Next I
    mnuScheme(Index).Checked = True
End Sub

Private Sub mnuVote_Click()
    Form2.Show 1
End Sub

Private Sub XPButton1_Click()
Form2.OpenWebsite "http://geocities.com/osen_kusnadi/newxpctl.zip"

End Sub
