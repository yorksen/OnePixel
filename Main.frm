VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "OnePixel"
   ClientHeight    =   402
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   462
   ControlBox      =   0   'False
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   402
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer ColorTimer 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const PIXEL_NUM = 1
Private Const TRANSPARENT_ALPHA = 127
Private Const COLOR_INTERVAL = 200

Private Sub ColorTimer_Timer()

Me.BackColor = RGB(Int((&HFF + 1) * Rnd), Int((&HFF + 1) * Rnd), Int((&HFF + 1) * Rnd))

End Sub

Private Sub Form_Load()
Dim ret As Long
Dim winExstyle As Long

' 设置窗口大小、位置以及置顶显示
ret = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, PIXEL_NUM, PIXEL_NUM, SWP_NOACTIVE Or SWP_SHOWWINDOW)
If 0 = ret Then
    MsgBox ("设置窗口位置失败")
    End
End If

' 设置窗口透明以及鼠标穿透效果
winExstyle = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
If 0 = winExstyle Then
    MsgBox ("获取窗口属性失败")
    End
End If

ret = SetWindowLong(Me.hWnd, GWL_EXSTYLE, winExstyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT)
If 0 = ret Then
    MsgBox ("设置窗口属性失败")
    End
End If

ret = SetLayeredWindowAttributes(Me.hWnd, 0, TRANSPARENT_ALPHA, LWA_ALPHA)
If 0 = ret Then
    MsgBox ("设置窗口透明且鼠标穿透失败")
    End
End If

' 启动窗口随机变色定时器
Randomize
ColorTimer.Interval = COLOR_INTERVAL

End Sub
