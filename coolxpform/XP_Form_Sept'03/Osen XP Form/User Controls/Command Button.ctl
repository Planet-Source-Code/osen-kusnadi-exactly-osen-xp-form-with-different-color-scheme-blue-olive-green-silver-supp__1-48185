VERSION 5.00
Begin VB.UserControl XPButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000006&
   ScaleHeight     =   51
   ScaleMode       =   0  'User
   ScaleWidth      =   79
   ToolboxBitmap   =   "Command Button.ctx":0000
   Begin VB.Timer HoverTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "XPButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'mouse over effects
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'draw and set rectangular area of the control
Private Declare Function GetClientRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long

'draw by pixel or by line
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal Hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Const PS_SOLID As Long = 0

'select and delete created objects
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long

'create regions of pixels and remove them to make the control transparent
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Const RGN_DIFF As Long = 4

'set text color and draw it to the control
Private Declare Function GetTextColor Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal Hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_CALCRECT As Long = &H400
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_CENTER As Long = &H1

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

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()

Private rc As RECT
Private W As Long, H As Long
Private rgMain As Long, rgn1 As Long
Private isOver As Boolean
Private flgHover As Integer
Private flgFocus As Boolean
Private LastButton As Integer
Private LastKey As Integer
Private R As Long, l As Long, t As Long, B As Long
Private mEnabled As Boolean
Private mCaption As String
Private mForeHover As OLE_COLOR

Private Sub DrawButton()
Dim pt As POINTAPI, Pen As Long, hPen As Long

  With UserControl
    'left top corner
    hPen = CreatePen(PS_SOLID, 1, RGB(122, 149, 168))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, t + 1, pt
    LineTo .Hdc, l + 2, t
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, RGB(37, 87, 131))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 2, t, pt
    LineTo .Hdc, l, t + 2
    SelectObject .Hdc, Pen
    DeleteObject hPen
    SetPixel .Hdc, l, t + 2, RGB(37, 87, 131)
    SetPixel .Hdc, l + 1, t + 2, RGB(191, 206, 220)
    SetPixel .Hdc, l + 2, t + 1, RGB(192, 207, 221)
    
    'top line
    hPen = CreatePen(PS_SOLID, 1, RGB(0, 60, 116))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 3, t, pt
    LineTo .Hdc, R - 2, t
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    'right top corner
    hPen = CreatePen(PS_SOLID, 1, RGB(37, 87, 131))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, R - 2, t, pt
    LineTo .Hdc, R + 1, t + 3
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, RGB(122, 149, 168))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, R - 1, t, pt
    LineTo .Hdc, R, t + 2
    SetPixel .Hdc, R, t + 1, RGB(122, 149, 168)
    SetPixel .Hdc, R - 2, t + 1, RGB(213, 223, 232)
    SetPixel .Hdc, R - 1, t + 2, RGB(191, 206, 219)
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    'right line
    hPen = CreatePen(PS_SOLID, 1, RGB(0, 60, 116))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, R, t + 3, pt
    LineTo .Hdc, R, B - 3
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    'right bottom corner
    hPen = CreatePen(PS_SOLID, 1, RGB(37, 87, 131))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, R, B - 3, pt
    LineTo .Hdc, R - 3, B
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, RGB(122, 149, 168))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, R, B - 2, pt
    LineTo .Hdc, R - 2, B
    SetPixel .Hdc, R - 2, B - 2, RGB(177, 183, 182)
    SetPixel .Hdc, R - 1, B - 3, RGB(182, 189, 189)
    SelectObject .Hdc, Pen
    DeleteObject hPen
  
    'bottom line
    hPen = CreatePen(PS_SOLID, 1, RGB(0, 60, 116))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 3, B - 1, pt
    LineTo .Hdc, R - 2, B - 1
    SelectObject .Hdc, Pen
    DeleteObject hPen
  
    'left bottom corner
    hPen = CreatePen(PS_SOLID, 1, RGB(37, 87, 131))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, B - 3, pt
    LineTo .Hdc, l + 3, B
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, RGB(122, 149, 168))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, B - 2, pt
    LineTo .Hdc, l + 2, B
    SetPixel .Hdc, l + 1, B - 3, RGB(191, 199, 202)
    SetPixel .Hdc, l + 2, B - 2, RGB(163, 174, 180)
    SelectObject .Hdc, Pen
    DeleteObject hPen
  
    'left line
    hPen = CreatePen(PS_SOLID, 1, RGB(0, 60, 116))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, t + 3, pt
    LineTo .Hdc, l, B - 3
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
  End With
End Sub
Private Sub DrawFocus()
Dim pt As POINTAPI, Pen As Long, hPen As Long
Dim i As Long, ColorR As Long, ColorG As Long, ColorB As Long
  With UserControl
    'top line
    hPen = CreatePen(PS_SOLID, 1, RGB(206, 231, 251))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 2, t + 1, pt
    LineTo .Hdc, R - 1, t + 1
    SelectObject .Hdc, Pen
    DeleteObject hPen
  
    hPen = CreatePen(PS_SOLID, 1, RGB(188, 212, 246))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 1, t + 2, pt
    LineTo .Hdc, R, t + 2
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    'draw gradient
    ColorR = 186
    ColorG = 211
    ColorB = 246
    For i = t + 3 To B - 4 Step 1
      hPen = CreatePen(PS_SOLID, 2, RGB(ColorR, ColorG, ColorB))
      Pen = SelectObject(.Hdc, hPen)
      MoveToEx .Hdc, l + 2, i, pt
      LineTo .Hdc, l + 2, i + 1
      MoveToEx .Hdc, R - 1, i, pt
      LineTo .Hdc, R - 1, i + 1
      SelectObject .Hdc, Pen
      DeleteObject hPen
      If ColorB >= 228 Then
        ColorR = ColorR - 4
        ColorG = ColorG - 3
        ColorB = ColorB - 1
      End If
    Next i
    
    hPen = CreatePen(PS_SOLID, 1, RGB(ColorR, ColorG, ColorB))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 1, B - 3, pt
    LineTo .Hdc, R - 1, B - 3
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    SetPixel .Hdc, l + 2, B - 2, RGB(77, 125, 193)
    hPen = CreatePen(PS_SOLID, 1, RGB(97, 125, 229))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 3, B - 2, pt
    LineTo .Hdc, R - 2, B - 2
    SetPixel .Hdc, R - 2, B - 2, RGB(77, 125, 193)
    
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
  End With
End Sub
Private Sub DrawHighlight()
Dim pt As POINTAPI, Pen As Long, hPen As Long
Dim i As Long, ColorR As Long, ColorG As Long, ColorB As Long
  With UserControl
    'top line
    hPen = CreatePen(PS_SOLID, 1, RGB(255, 240, 207))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 2, t + 1, pt
    LineTo .Hdc, R - 1, t + 1
    SelectObject .Hdc, Pen
    DeleteObject hPen
  
    hPen = CreatePen(PS_SOLID, 1, RGB(253, 216, 137))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 1, t + 2, pt
    LineTo .Hdc, R, t + 2
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    'draw gradient
    ColorR = 254
    ColorG = 223
    ColorB = 154
    For i = t + 2 To B - 3 Step 1
      hPen = CreatePen(PS_SOLID, 1, RGB(ColorR, ColorG, ColorB))
      Pen = SelectObject(.Hdc, hPen)
      MoveToEx .Hdc, l + 1, i, pt
      LineTo .Hdc, l + 1, i + 1
      MoveToEx .Hdc, R - 1, i, pt
      LineTo .Hdc, R - 1, i + 1
      SelectObject .Hdc, Pen
      DeleteObject hPen
      If ColorB >= 49 Then
        ColorR = ColorR - 1
        ColorG = ColorG - 3
        ColorB = ColorB - 7
      End If
    Next i
    ColorR = 252
    ColorG = 210
    ColorB = 121
    For i = t + 3 To B - 3 Step 1
      hPen = CreatePen(PS_SOLID, 1, RGB(ColorR, ColorG, ColorB))
      Pen = SelectObject(.Hdc, hPen)
      MoveToEx .Hdc, l + 2, i, pt
      LineTo .Hdc, l + 2, i + 1
      MoveToEx .Hdc, R - 2, i, pt
      LineTo .Hdc, R - 2, i + 1
      SelectObject .Hdc, Pen
      DeleteObject hPen
      If ColorB >= 57 Then
        ColorR = ColorR - 1
        ColorG = ColorG - 4
        ColorB = ColorB - 8
      End If
    Next i
    
    hPen = CreatePen(PS_SOLID, 1, RGB(ColorR, ColorG, ColorB))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 3, B - 3, pt
    LineTo .Hdc, R, B - 3
    SelectObject .Hdc, Pen
    DeleteObject hPen
        
    hPen = CreatePen(PS_SOLID, 1, RGB(229, 151, 0))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l + 2, B - 2, pt
    LineTo .Hdc, R - 1, B - 2
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
  End With
End Sub

Private Sub DrawButtonFace()
Dim pt As POINTAPI, Pen As Long, hPen As Long
Dim i As Long, ColorR As Long, ColorG As Long, ColorB As Long
  
  With UserControl
  
    .AutoRedraw = True
    .Cls
    .ScaleMode = 3
    
    'draw gradient
    ColorR = 255
    ColorG = 255
    ColorB = 253
    
    For i = t + 3 To B - 3 Step 1
      hPen = CreatePen(PS_SOLID, 1, RGB(ColorR, ColorG, ColorB))
      Pen = SelectObject(.Hdc, hPen)
      MoveToEx .Hdc, l, i, pt
      LineTo .Hdc, R, i
      SelectObject .Hdc, Pen
      DeleteObject hPen
      
      If ColorB >= 230 Then
        ColorR = ColorR - 1
        ColorG = ColorG - 1
        ColorB = ColorB - 1
      End If
    Next i
    
    'bottom shadow
    hPen = CreatePen(PS_SOLID, 1, RGB(214, 208, 197))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, B - 2, pt
    LineTo .Hdc, R, B - 2
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, RGB(226, 223, 214))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, B - 3, pt
    LineTo .Hdc, R, B - 3
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, RGB(236, 235, 230))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, B - 4, pt
    LineTo .Hdc, R, B - 4
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
  End With
  
End Sub
Private Sub DrawButtonDown()
Dim pt As POINTAPI, Pen As Long, hPen As Long
Dim i As Long, ColorR As Long, ColorG As Long, ColorB As Long
  With UserControl
    .AutoRedraw = True
    .Cls
    .ScaleMode = 3
    'draw gradient
    ColorR = 226
    ColorG = 225
    ColorB = 218
    For i = t + 3 To B - 2 Step 4
      hPen = CreatePen(PS_SOLID, 4, RGB(ColorR, ColorG, ColorB))
      Pen = SelectObject(.Hdc, hPen)
      MoveToEx .Hdc, l, i, pt
      LineTo .Hdc, R, i
      SelectObject .Hdc, Pen
      DeleteObject hPen
      If ColorB >= 218 Then
        ColorR = ColorR - 1
        ColorG = ColorG - 1
        ColorB = ColorB - 1
      End If
    Next i
    'top shadow
    hPen = CreatePen(PS_SOLID, 1, RGB(209, 204, 192))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, t + 1, pt
    LineTo .Hdc, R, t + 1
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, RGB(220, 216, 207))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, t + 2, pt
    LineTo .Hdc, R, t + 2
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    'bottom shadow
    hPen = CreatePen(PS_SOLID, 1, RGB(234, 233, 227))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, B - 3, pt
    LineTo .Hdc, R, B - 3
    SelectObject .Hdc, Pen
    DeleteObject hPen
    
    hPen = CreatePen(PS_SOLID, 1, RGB(242, 241, 238))
    Pen = SelectObject(.Hdc, hPen)
    MoveToEx .Hdc, l, B - 2, pt
    LineTo .Hdc, R, B - 2
    SelectObject .Hdc, Pen
    DeleteObject hPen
  End With
End Sub
Private Sub DrawButtonDisabled()
Dim pt As POINTAPI, Pen As Long, hPen As Long
Dim i As Long, ColorR As Long, ColorG As Long, ColorB As Long
Dim hBrush As Long

  With UserControl
    .AutoRedraw = True
    .Cls
    .ScaleMode = 3
    hBrush = CreateSolidBrush(RGB(245, 244, 234))
    FillRect UserControl.Hdc, rc, hBrush
    DeleteObject hBrush
  
    hBrush = CreateSolidBrush(RGB(201, 199, 186))
    FrameRect UserControl.Hdc, rc, hBrush
    DeleteObject hBrush
    
    'Left top corner
    SetPixel .Hdc, l, t + 1, RGB(216, 213, 199)
    SetPixel .Hdc, l + 1, t + 1, RGB(216, 213, 199)
    SetPixel .Hdc, l + 1, t, RGB(216, 213, 199)
    SetPixel .Hdc, l + 1, t + 2, RGB(234, 233, 222)
    SetPixel .Hdc, l + 2, t + 1, RGB(234, 233, 222)
    'right top corner
    SetPixel .Hdc, R - 1, t, RGB(216, 213, 199)
    SetPixel .Hdc, R - 1, t + 1, RGB(216, 213, 199)
    SetPixel .Hdc, R, t + 1, RGB(216, 213, 199)
    SetPixel .Hdc, R - 2, t + 1, RGB(234, 233, 222)
    SetPixel .Hdc, R - 1, t + 2, RGB(234, 233, 222)
    'left bottom corner
    SetPixel .Hdc, l, B - 2, RGB(216, 213, 199)
    SetPixel .Hdc, l + 1, B - 2, RGB(216, 213, 199)
    SetPixel .Hdc, l + 1, B - 1, RGB(216, 213, 199)
    SetPixel .Hdc, l + 1, B - 3, RGB(234, 233, 222)
    SetPixel .Hdc, l + 2, B - 2, RGB(234, 233, 222)
    'right bottom corner
    SetPixel .Hdc, R, B - 2, RGB(216, 213, 199)
    SetPixel .Hdc, R - 1, B - 2, RGB(216, 213, 199)
    SetPixel .Hdc, R - 1, B - 1, RGB(216, 213, 199)
    SetPixel .Hdc, R - 1, B - 3, RGB(234, 233, 222)
    SetPixel .Hdc, R - 2, B - 2, RGB(234, 233, 222)
  End With

End Sub
Private Sub DrawButton2()
Dim pt As POINTAPI, Pen As Long, hPen As Long
Dim i As Long, ColorR As Long, ColorG As Long, ColorB As Long
Dim hBrush As Long

  With UserControl
  
  
    hBrush = CreateSolidBrush(RGB(0, 60, 116))
    FrameRect UserControl.Hdc, rc, hBrush
    DeleteObject hBrush
    
    'Left top corner
    SetPixel .Hdc, l, t + 1, RGB(122, 149, 168)
    SetPixel .Hdc, l + 1, t + 1, RGB(37, 87, 131)
    SetPixel .Hdc, l + 1, t, RGB(122, 149, 168)
    'SetPixel .hdc, l + 1, t + 2, RGB(191, 206, 220)
    'SetPixel .hdc, l + 2, t + 1, RGB(192, 207, 221)
    
    'right top corner
    SetPixel .Hdc, R - 1, t, RGB(122, 149, 168)
    SetPixel .Hdc, R - 1, t + 1, RGB(37, 87, 131)
    SetPixel .Hdc, R, t + 1, RGB(122, 149, 168)
    'SetPixel .hdc, r - 2, t + 1, RGB(234, 233, 222)
    'SetPixel .hdc, r - 1, t + 2, RGB(234, 233, 222)
    
    'left bottom corner
    SetPixel .Hdc, l, B - 2, RGB(122, 149, 168)
    SetPixel .Hdc, l + 1, B - 2, RGB(37, 87, 131)
    SetPixel .Hdc, l + 1, B - 1, RGB(122, 149, 168)
    'SetPixel .hdc, l + 1, b - 3, RGB(234, 233, 222)
    'SetPixel .hdc, l + 2, b - 2, RGB(234, 233, 222)
    
    'right bottom corner
    SetPixel .Hdc, R, B - 2, RGB(122, 149, 168)
    SetPixel .Hdc, R - 1, B - 2, RGB(37, 87, 131)
    SetPixel .Hdc, R - 1, B - 1, RGB(122, 149, 168)
    'SetPixel .hdc, r - 1, b - 3, RGB(234, 233, 222)
    'SetPixel .hdc, r - 2, b - 2, RGB(234, 233, 222)
  End With

End Sub
Private Sub RedrawButton(Optional ByVal Stat As Integer = -1)
  If mEnabled Then
    If Stat = 1 And LastButton = 1 Then
      DrawButtonDown
    Else
      DrawButtonFace
      If isOver = True Then
        DrawHighlight
      Else
        If flgFocus = True Then
          DrawFocus
        End If
      End If
    End If
    DrawButton2
  Else
    DrawButtonDisabled
  End If
  DrawCaption
  MakeRegion
  
End Sub
Private Sub DrawCaption()
Dim vh As Long, rcTxt As RECT
  
  With UserControl
    GetClientRect .Hwnd, rcTxt
    If mEnabled Then
      If isOver Then
        SetTextColor .Hdc, mForeHover
      Else
        SetTextColor .Hdc, .ForeColor
      End If
    Else
      SetTextColor .Hdc, RGB(161, 161, 146)
    End If
    vh = DrawText(.Hdc, mCaption, Len(mCaption), rcTxt, DT_CALCRECT Or DT_CENTER Or DT_WORDBREAK)
    'If Button = 1 Then
    '  SetRect rcTxt, 0, (.ScaleHeight * 0.5) - (vh * 0.5) + 1, .ScaleWidth, (.ScaleHeight * 0.5) + (vh * 0.5) + 1
    '  DrawText .hdc, mCaption, Len(mCaption), rcTxt, DT_CENTER Or DT_WORDBREAK
    'Else
      SetRect rcTxt, 0, (.ScaleHeight * 0.5) - (vh * 0.5), .ScaleWidth, (.ScaleHeight * 0.5) + (vh * 0.5)
      DrawText .Hdc, mCaption, Len(mCaption), rcTxt, DT_CENTER Or DT_WORDBREAK
    'End If
  End With
End Sub
Private Sub HoverTimer_Timer()
  If Not isMouseOver Then
    HoverTimer.Enabled = False
    isOver = False
    flgHover = 0
    RedrawButton 0
    RaiseEvent MouseOut
  End If
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  LastButton = 1
  Call UserControl_Click
End Sub

Private Sub UserControl_Click()
  If LastButton = 1 Then
        RedrawButton 0
'        LastButton = 1
'        RedrawButton 1
        UserControl.Refresh
        RaiseEvent Click
  End If
End Sub

Private Sub UserControl_DblClick()
  If LastButton = 1 Then
    Call UserControl_MouseDown(1, 0, 0, 0)
    SetCapture Hwnd
  End If
End Sub

Private Sub UserControl_GotFocus()
  flgFocus = True
  If mEnabled = True Then
    LastButton = 1
    UserControl.Refresh
    RedrawButton 0
  End If
End Sub

Private Sub UserControl_InitProperties()
  Set UserControl.Font = Ambient.Font
  mCaption = "Command" & Mid(Ambient.DisplayName, 9)
  mEnabled = True
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  LastKey = KeyCode
  Select Case KeyCode
    Case vbKeySpace, vbKeyReturn
      RedrawButton 1
    Case vbKeyLeft, vbKeyRight 'right and down arrows
      SendKeys "{Tab}"
    Case vbKeyDown, vbKeyUp 'left and up arrows
      SendKeys "+{Tab}"
  End Select
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  If ((KeyCode = vbKeySpace) And (LastKey = vbKeySpace)) Or ((KeyCode = vbKeyReturn) And (LastKey = vbKeyReturn)) Then
    RedrawButton 0
    LastButton = 1
    UserControl.Refresh
    RaiseEvent Click
  End If
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
  flgFocus = False
  RedrawButton 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If mEnabled = True Then
    RaiseEvent MouseDown(Button, Shift, X, Y)
    LastButton = Button
    UserControl.Refresh
    DoEvents
    RedrawButton 1
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
 '  UserControl_GotFocus
      
  If Button < 2 Then
    If Not isMouseOver Then
      If flgHover = 0 Then Exit Sub
      RedrawButton 0
    Else
      If flgHover = 1 Then Exit Sub
      flgHover = 1
      If Button = 0 And Not isOver Then
        HoverTimer.Enabled = True
        isOver = True
        flgHover = 0
        RedrawButton 0
        RaiseEvent MouseOver
      ElseIf Button = 1 Then
        isOver = True
        RedrawButton 1
        isOver = False
      End If
    End If
  End If
  RaiseEvent MouseMove(Button, Shift, X, Y)
  
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
  RedrawButton 0
  UserControl.Refresh
End Sub


Private Sub UserControl_Resize()
  GetClientRect UserControl.Hwnd, rc
  With rc
    R = .Right - 1: l = .Left: t = .Top: B = .Bottom
    W = .Right: H = .Bottom
  End With
  RedrawButton 0
End Sub
Private Function isMouseOver() As Boolean
Dim pt As POINTAPI
GetCursorPos pt
isMouseOver = (WindowFromPoint(pt.X, pt.Y) = Hwnd)
End Function
Private Sub MakeRegion()
  DeleteObject rgMain
  rgMain = CreateRectRgn(0, 0, W, H)
  rgn1 = CreateRectRgn(0, 0, 1, 1)            'Left top coner
  CombineRgn rgMain, rgMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(0, H - 1, 1, H)      'Left bottom corner
  CombineRgn rgMain, rgMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(W - 1, 0, W, 1)      'Right top corner
  CombineRgn rgMain, rgMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  rgn1 = CreateRectRgn(W - 1, H - 1, W, H) 'Right bottom corner
  CombineRgn rgMain, rgMain, rgn1, RGN_DIFF
  DeleteObject rgn1
  SetWindowRgn UserControl.Hwnd, rgMain, True
End Sub
Public Property Get Enabled() As Boolean
  Enabled = mEnabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
  mEnabled = NewValue
  PropertyChanged "Enabled"
  UserControl.Enabled = NewValue
  RedrawButton 0
End Property
Public Property Get Font() As Font
  Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal NewValue As Font)
  Set UserControl.Font = NewValue
  RedrawButton 0
  PropertyChanged "Font"
End Property
Public Property Get Caption() As String
  Caption = mCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
  mCaption = NewValue
  RedrawButton 0
  SetAccessKeys
  PropertyChanged "Caption"
End Property
Public Property Get ForeColor() As OLE_COLOR
  ForeColor = UserControl.ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
  UserControl.ForeColor = NewValue
  RedrawButton 0
  PropertyChanged "ForeColor"
End Property
Public Property Get ForeHover() As OLE_COLOR
  ForeHover = mForeHover
End Property
Public Property Let ForeHover(ByVal NewValue As OLE_COLOR)
  mForeHover = NewValue
  PropertyChanged "ForeHover"
End Property
Private Sub UserControl_Show()
  RedrawButton 0
 
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    mEnabled = .ReadProperty("Enabled", True)
    Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
    mCaption = .ReadProperty("Caption", Ambient.DisplayName)
    UserControl.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
    mForeHover = .ReadProperty("ForeHover", UserControl.ForeColor)
  End With
  UserControl.Enabled = mEnabled
  SetAccessKeys
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "Enabled", mEnabled, True
    .WriteProperty "Font", UserControl.Font, Ambient.Font
    .WriteProperty "Caption", mCaption, Ambient.DisplayName
    .WriteProperty "ForeColor", UserControl.ForeColor
    .WriteProperty "ForeHover", mForeHover, Ambient.ForeColor
  End With
End Sub
Private Sub SetAccessKeys()
Dim i As Long
UserControl.AccessKeys = ""
  If Len(mCaption) > 1 Then
    i = InStr(1, mCaption, "&", vbTextCompare)
    If (i < Len(mCaption)) And (i > 0) Then
      If Mid$(mCaption, i + 1, 1) <> "&" Then
        UserControl.AccessKeys = LCase$(Mid$(mCaption, i + 1, 1))
      Else
        i = InStr(i + 2, mCaption, "&", vbTextCompare)
        If Mid$(mCaption, i + 1, 1) <> "&" Then
          UserControl.AccessKeys = LCase$(Mid$(mCaption, i + 1, 1))
        End If
      End If
    End If
  End If
End Sub
