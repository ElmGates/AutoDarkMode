Attribute VB_Name = "Module1"
'-----------------------模块中--------------------------
Option Explicit '在系统托盘显示图标的模块
Public Const MAX_TOOLTIP As Integer = 64
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const SW_RESTORE = 9
Public Const SW_HIDE = 0
Public nfIconData As NOTIFYICONDATA
Public Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uID As Long
uFlags As Long
uCallbackMessage As Long
hIcon As Long
szTip As String * MAX_TOOLTIP
End Type
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

