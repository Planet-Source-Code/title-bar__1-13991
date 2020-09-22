Attribute VB_Name = "MSysMenu"
Option Explicit

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Const WM_PAINT = &HF
Const WM_NCPAINT = &H85
Const WM_NCLBUTTONDBLCLK = &HA3
Const WM_NCLBUTTONUP = &HA2
Const WM_NCHITTEST = &H84
Const WM_ACTIVATE = &H6
Const WM_NCMBUTTONDBLCLK = &HA9
Const WM_NCLBUTTONDOWN = &HA1
Const WM_NCMBUTTONDOWN = &HA7
Const WM_NCRBUTTONDOWN = &HA4
Const WM_NCRBUTTONUP = &HA5
Const WM_NCMOUSEHOVER = &H2A0
Const WM_NCMOUSEMOVE = &HA0

Const WM_MOVE = &H3
Const WM_MOUSEHOVER = &H2A1
Const WM_MOUSEMOVE = &H200
Const WM_ERASEBKGND = &H14
Const HTCAPTION = 2
Const HTCLIENT = 1
Const WM_GETMINMAXINFO = &H24
Const WM_SIZE = &H5
Const WM_SIZING = &H214
Const WM_MOVING = &H216
Const SRCCOPY = &HCC0020
Const SM_CYFRAME = 33
Const SM_CXFRAME = 32
Const SM_CYSIZE = 31
Const SM_CXSIZE = 30
Public procOld As Long
Public Const LR_LOADFROMFILE = &H10
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1
Public Const IMAGE_CURSOR = 2
Public Const IMAGE_ENHMETAFILE = 3
Public Const CF_BITMAP = 2
Public Const ScrCopy = &HCC0020
Const BDR_SUNKENOUTER = &H2
Const BDR_SUNKENINNER = &H8
Const BDR_RAISEDINNER = &H4
Const BDR_RAISEDOUTER = &H1
Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Const BF_SOFT = &H1000
Public Const BF_LEFT = &H1
Const MK_LBUTTON = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Const BF_ADJUST = &H2000
Public Const GWL_WNDPROC As Long = (-4&)
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, _
                                                              ByVal nIndex&, ByVal dwNewLong&)
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal numBytes As Long)
Public Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, _
                                                    ByVal hwnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Dim retVal As Boolean
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Dim TRect As RECT
Dim memDC As Long
Dim pDC As Long
Dim bolPressed As Boolean
Public lngBitMap As Long
Dim BtnDown As Boolean
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Function MenuProc(ByVal hwnd As Long, ByVal iMsg As Long, _
                                           ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case iMsg
    Case WM_NCMOUSEMOVE
        Call GetButtonRect(TRect)
        If PtInRect(TRect, lParam And &HFFFF&, lParam \ &H10000) <> 0 Then
           
            If wParam = HTCAPTION Then
                If BtnDown = True Then
                    bolPressed = False
                    DrawCapButton
                End If
            End If
        End If
    Case WM_MOUSEMOVE
        Call GetButtonRect(TRect)
           bolPressed = True
            DrawCapButton
    Case WM_PAINT
        DrawCapButton
    Case WM_NCPAINT
        DrawCapButton
    Case WM_ACTIVATE
        DrawCapButton
     Case WM_ERASEBKGND
        DrawCapButton
    Case WM_NCLBUTTONDOWN
        If wParam = HTCAPTION Then
            Call GetButtonRect(TRect)
            If PtInRect(TRect, lParam And &HFFFF&, lParam \ &H10000) <> 0 Then
                bolPressed = False
                DrawCapButton
                wParam = HTCLIENT
                BtnDown = True
            End If
            
        End If
    Case WM_NCLBUTTONUP
      
       If PtInRect(TRect, lParam And &HFFFF&, lParam \ &H10000) <> 0 Then
             bolPressed = True
            DrawCapButton
            BtnDown = False
        End If
       BtnDown = False
       'Call Initialise(Form1)
       'here put your code for click event
    Case WM_NCLBUTTONDBLCLK
        If PtInRect(TRect, lParam And &HFFFF&, lParam \ &H10000) <> 0 Then
            wParam = HTCLIENT 'if wParam not changed form will restore
        End If
    
        
End Select
MenuProc = CallWindowProc(procOld, hwnd, iMsg, wParam, lParam)
End Function
Public Function GetButtonRect(ByRef TRect As RECT)
Call GetWindowRect(Form1.hwnd, TRect)
TRect.Top = TRect.Top + GetSystemMetrics(SM_CYFRAME) + 2
TRect.Bottom = TRect.Top + GetSystemMetrics(SM_CYSIZE) - 4
TRect.Left = TRect.Right - GetSystemMetrics(SM_CXFRAME) - (4 * GetSystemMetrics(SM_CXSIZE)) - 1
TRect.Right = TRect.Left + GetSystemMetrics(SM_CXSIZE) - 3
End Function
Public Function DrawCapButton()
Dim RectWnd As RECT, CRect As RECT
Dim rtn As Long
If Not (Not IsWindowVisible(Form1.hwnd) And IsIconic(Form1.hwnd)) Then
    pDC = GetWindowDC(Form1.hwnd)
    memDC = CreateCompatibleDC(pDC)
    Call SelectObject(memDC, lngBitMap)
        
    Call GetButtonRect(CRect)
    Call GetWindowRect(Form1.hwnd, RectWnd)
    Call OffsetRect(CRect, -RectWnd.Left, -RectWnd.Top)
    CRect.Left = CRect.Left
    CRect.Top = CRect.Top
    CRect.Right = CRect.Right
    CRect.Bottom = CRect.Bottom
    StretchBlt pDC, CRect.Left, CRect.Top, CRect.Right - CRect.Left, CRect.Bottom - CRect.Top, memDC, 0, 0, 16, 16, SRCCOPY
    If bolPressed = True Then
         DrawEdge pDC, CRect, EDGE_RAISED, BF_RECT
         
    Else
       DrawEdge pDC, CRect, BDR_SUNKENINNER, BF_RECT
    End If
    DeleteDC (memDC)
    Call ReleaseDC(Form1.hwnd, pDC)
End If
End Function

'
