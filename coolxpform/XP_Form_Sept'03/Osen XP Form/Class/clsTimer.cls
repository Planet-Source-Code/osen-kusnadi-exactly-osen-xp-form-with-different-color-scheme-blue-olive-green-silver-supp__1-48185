VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Private Declare Function SetTimer& Lib "user32" (ByVal Hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
Private Declare Function KillTimer& Lib "user32" (ByVal Hwnd As Long, ByVal nIDEvent As Long)

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40

'mouse over effects
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type FunctionSPointerS
    FunctionPtr As Long
    FunctionAddress As Long
End Type

Private mint_interval As Integer
Private mint_ticks As Long
Private mb_running As Boolean
Private mlng_memhandle As Long
Private mlng_proc As Long
Private mlng_handle As Long
Private mlng_hwnd As Long
Private mb_wait As Boolean


Public Event OnTime(ByVal Int_Ticks As Long, ByVal DwTime As Long)

Public Sub TimerProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal DwTime As Long)

    mint_ticks = mint_ticks + 1
    RaiseEvent OnTime(mint_ticks, DwTime)
    mb_wait = False

End Sub

Public Property Get Interval() As Integer

    Interval = mint_interval

End Property

Public Property Let Interval(ByVal int_interval As Integer)

    If mint_interval <> int_interval Then
        mint_interval = int_interval
        If mb_running Then
            mlng_handle = SetTimer(mlng_hwnd, 0, mint_interval, mlng_proc)
        End If
    End If

End Property

Public Property Get Hwnd() As Long

    Hwnd = mlng_hwnd

End Property

Public Property Let Hwnd(ByVal lng_hwnd As Long)

    If Not mb_running Then
        mlng_hwnd = lng_hwnd
    End If

End Property

Public Sub Wait(Optional ByVal int_interval As Integer)

    mb_wait = True
    StartTimer int_interval
    Do Until Not mb_wait
        DoEvents
    Loop
    StopTimer

End Sub

Public Sub StartTimer(Optional ByVal int_interval As Integer)

    If Not mb_running Then
        mint_ticks = 0
        mb_running = True
        If int_interval <= 0 Or int_interval = mint_interval Then
            int_interval = mint_interval
        End If
        mlng_handle = SetTimer(mlng_hwnd, 0, int_interval, mlng_proc)
    End If

End Sub

Public Sub StopTimer()

    If mb_running Then
        KillTimer mlng_hwnd, mlng_handle
        mb_running = False
    End If

End Sub

Private Sub Class_Initialize()

  Dim lng_fncptr As Long
  Dim lng_objptr As Long
  Dim lng_vtable As Long
  Dim lng_ptx As Long


  Dim lng_varnum As Long
  Dim lng_objvarnum As Long
  Dim lng_funcnum As Long


    mint_interval = 1000

    lng_objptr = ObjPtr(Me)
    CopyMemory lng_vtable, ByVal lng_objptr, 4
    lng_ptx = lng_vtable + 28 + (lng_varnum * 2 * 4) + (lng_objvarnum * 3 * 4) + lng_funcnum * 4
    CopyMemory lng_fncptr, ByVal lng_ptx, 4
    mlng_memhandle = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, 105)
    mlng_proc = GlobalLock(mlng_memhandle)
    DelegateFunction mlng_proc, Me, lng_fncptr, 4

End Sub
Public Function isMouseOver(ByVal Hwnd As Long) As Boolean
    Dim pt As POINTAPI
    GetCursorPos pt
    isMouseOver = (WindowFromPoint(pt.X, pt.Y) = Hwnd)
End Function

Private Sub Class_Terminate()

    If mb_running Then StopTimer
    Call GlobalUnlock(mlng_memhandle)
    Call GlobalFree(mlng_memhandle)

End Sub

Private Function DelegateFunction(ByVal CallingADR As Long, Obj As Object, ByVal MethodAddress As Long, ByVal NumberOfParameters As Byte) As Boolean

  Dim TmpA As Long
  Dim u As Long
  Dim StackP As Byte
  Dim PERFCALL As Long

    On Error GoTo NotSuccess
    TmpA = CallingADR
    CopyMemory ByVal CallingADR, &H68EC8B55, 4
    CallingADR = CallingADR + 4
    CopyMemory ByVal CallingADR, TmpA + 31 + (NumberOfParameters * 3) - 4, 4
    CallingADR = CallingADR + 4

    StackP = 4 + 4 * NumberOfParameters

    For u = 1 To NumberOfParameters
        CopyMemory ByVal CallingADR, CInt(&H75FF), 2
        CallingADR = CallingADR + 2
        CopyMemory ByVal CallingADR, StackP, 1
        CallingADR = CallingADR + 1
        StackP = StackP - 4
    Next u

    CopyMemory ByVal CallingADR, CByte(&H68), 1
    CallingADR = CallingADR + 1
    CopyMemory ByVal CallingADR, ObjPtr(Obj), 4
    CallingADR = CallingADR + 4
    CopyMemory ByVal CallingADR, CByte(&HE8), 1
    CallingADR = CallingADR + 1
    PERFCALL = CallingADR - TmpA - 1
    PERFCALL = MethodAddress - (TmpA + (CallingADR - TmpA - 1)) - 5
    CopyMemory ByVal CallingADR, PERFCALL, 4
    CallingADR = CallingADR + 4
    CopyMemory ByVal CallingADR, CByte(&HA1), 1
    CallingADR = CallingADR + 1
    CopyMemory ByVal CallingADR, TmpA + 31 + (NumberOfParameters * 3) - 4, 4
    CallingADR = CallingADR + 4
    CopyMemory ByVal CallingADR, CInt(&HC2C9), 2

    CallingADR = CallingADR + 2
    CopyMemory ByVal CallingADR, CInt(NumberOfParameters * 4), 2
    DelegateFunction = True

Exit Function

NotSuccess:
    On Error GoTo 0

End Function

