VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "System Menu Demo"
   ClientHeight    =   2964
   ClientLeft      =   2376
   ClientTop       =   1428
   ClientWidth     =   6036
   Icon            =   "frmTitleBar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   WindowState     =   1  'Minimized
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   492
      Picture         =   "frmTitleBar.frx":0442
      ScaleHeight     =   192
      ScaleWidth      =   192
      TabIndex        =   0
      Top             =   1992
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    lngBitMap = Picture2.Picture
    'lngBitMap = LoadImage(App.hInstance, "C:\Program Files\Microsoft Visual Studio\Common\Graphics\Bitmaps\OffCtlBr\Small\Color\SAVE.BMP", IMAGE_BITMAP, 16, 16, LR_LOADFROMFILE)
    procOld = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MenuProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call SetWindowLong(hwnd, GWL_WNDPROC, procOld)
End Sub

