VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�Զ���ɫģʽ������"
   ClientHeight    =   2730
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   6090
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command6 
      Caption         =   "������"
      Height          =   615
      Left            =   3840
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "����"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ˢ��"
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����ģʽ"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�����"
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����趨ʱ��"
      Height          =   615
      Left            =   1200
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5880
      Top             =   120
   End
   Begin VB.Label ssks 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "��ɫģʽ��ʼʱ��(ʱ):"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label qsks 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "ǳɫģʽ��ʼʱ��(ʱ):"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "��ȡ��ʱ��(ʱ):"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "��ǰʱ��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "˵��:������������Windows10��Windows11,�������Զ����趨��ʱ�佫���ĵ����л�Ϊ��ɫģʽ����ǳɫģʽ."
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Menu F00 
      Caption         =   "�Ҽ��˵�"
      Visible         =   0   'False
      Begin VB.Menu F01 
         Caption         =   "��"
      End
      Begin VB.Menu F02 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------������--------------------------
Dim WindowTop, WindowLeft
Dim timess As String
Dim timehh As String
Dim usersetday As String
Dim usersetnight As String
Dim userconfig As String
Dim dayt As Integer
Dim nightt As Integer
Dim timehhh As Integer
Dim kssj As String
Dim jssj As String
Dim zhsj As String
Dim passwd As String
Dim sfzd As Integer
Function WindowStyle() '���°ѳ������System Tray====================================System Tray Begin
With nfIconData
.hWnd = Me.hWnd
.uID = Me.Icon
.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
.uCallbackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon.Handle '��������ƶ���������ʱ��ʾ��Tip
.szTip = "�Զ���ɫģʽ,˫�����Դ�������." & vbNullChar
.cbSize = Len(nfIconData)
End With
Call Shell_NotifyIcon(NIM_ADD, nfIconData) '=========================================System Tray End
Me.Hide
End Function
Private Sub Command1_Click()
Form2.Show
End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Command3_Click()
sfzd = MsgBox("��ȷ��Ҫ���뿪��ģʽ�𣿿���ģʽ�п��ܴ���δ֪�����⣡", vbOKCancel, "��Ҫ����")
If sfzd = 1 Then
frmLogin.Show
End If
End Sub
Private Sub Command4_Click()
Name ".\userconfig.inconfig" As ".\userconfig.txt"
Open ".\userconfig.txt" For Input As #5
Line Input #5, userconfig
Close #5
usersetday = Mid(userconfig, 1, 2)
usersetnight = Mid(userconfig, 3, 2)
Name ".\userconfig.txt" As ".\userconfig.inconfig"
qsks.Caption = usersetday
ssks.Caption = usersetnight
dayt = Val(usersetday)
nightt = Val(usersetnight)
timehhh = Val(timehh)
If timehhh > dayt And timehhh < nightt Then
DayAuto
ElseIf (timehhh > nightt And timehhh) < 24 Or (timehhh > 0 And timehhh < dayt) Then
NightAuto
End If
End Sub
Private Sub Command5_Click()
frmAbout.Show
End Sub
Private Sub Command6_Click()
ppp = MsgBox("�����аٶ��������������", vbOKOnly, "Error")
End Sub
Private Sub Form_Load() '��ȡ����Ĳ���,��д����
If Dir(".\userconfig.inconfig") = "" Then
Open ".\userconfig.txt" For Output As #1
Print #1, "0718"
Close #1
Name ".\userconfig.txt" As ".\userconfig.inconfig"
Else
Name ".\userconfig.inconfig" As ".\userconfig.txt"
Open ".\userconfig.txt" For Input As #2
Line Input #2, userconfig
Close #2
usersetday = Mid(userconfig, 1, 2)
usersetnight = Mid(userconfig, 3, 2)
Name ".\userconfig.txt" As ".\userconfig.inconfig"
End If
qsks.Caption = usersetday
ssks.Caption = usersetnight
dayt = Val(usersetday)
nightt = Val(usersetnight)
timehhh = Val(timehh)
If timehhh > dayt And timehhh < nightt Then
DayAuto
ElseIf (timehhh > nightt And timehhh) < 24 Or (timehhh > 0 And timehhh < dayt) Then
NightAuto
End If
If Dir(".\none.cfgr") = "" Then
Open ".\none.txt" For Output As #1
Print #1, "12345678"
Close #1
Name ".\none.txt" As ".\none.cfgr"
Else
Name ".\none.cfgr" As ".\none.txt"
Open ".\none.txt" For Input As #2
Line Input #2, passwd
Close #2
Name ".\none.txt" As ".\none.cfgr"
End If
WindowStyle
End Sub
Private Sub Form_Resize()
WindowTop = Me.Top
WindowLeft = Me.Left
If Me.WindowState = 1 Then
WindowStyle
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single
lMsg = X / Screen.TwipsPerPixelX
Select Case lMsg
Case WM_LBUTTONDBLCLK '˫�������ʾ���壬Ҫ�ĳ������Ŀ�ģ����Ķ���
ShowWindow Me.hWnd, SW_RESTORE
Me.Top = WindowTop
Me.Left = WindowLeft 'Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
Me.SetFocus
Case WM_RBUTTONUP '������ͼ���ϵ��Ҽ���ʾ�˵�
PopupMenu F00 '�˵�����ΪF00�����˵�ʱ���Ըĳɱ�ģ�����Ҳ�øĳ���Ӧ��
End Select
End Sub
Private Sub F01_Click()
ShowWindow Me.hWnd, SW_RESTORE
Me.Top = WindowTop
Me.Left = WindowLeft
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub
Private Sub F02_Click()
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '�˳�����ʱɾ������ͼ��
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub
Function DayAuto() '�Զ��л��ռ�ģʽ
Dim WSH
Set WSH = CreateObject("WSCRIPT.SHELL")
WSH.Regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 1, "REG_DWORD"
WSH.Regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme", 1, "REG_DWORD"
End Function
Function NightAuto() '�Զ��л�ҹ��ģʽ
Dim WSH
Set WSH = CreateObject("WSCRIPT.SHELL")
WSH.Regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme", 0, "REG_DWORD"
WSH.Regwrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\SystemUsesLightTheme", 0, "REG_DWORD"
End Function
Function Autorunon() 'aoutrunon
Dim WSH
Set WSH = CreateObject("WSCRIPT.SHELL")
WSH.Regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
End Function
Private Sub Label5_Change()
timehhh = Val(timehh)
If timehhh > dayt And timehhh < nightt Then
DayAuto
ElseIf (timehhh > nightt And timehhh) < 24 Or (timehhh > 0 And timehhh < dayt) Then
NightAuto
End If
End Sub
Private Sub Timer1_Timer()
timess = Str(Time())
If Mid(timess, 2, 1) = ":" Then
timehh = Mid(timess, 1, 1)
Else
timehh = Mid(timess, 1, 2)
End If
Label3.Caption = timess
Label5.Caption = timehh
End Sub
