VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���뿪��ģʽ"
   ClientHeight    =   2730
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3600
   FillColor       =   &H000000FF&
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1612.974
   ScaleMode       =   0  'User
   ScaleWidth      =   3380.205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   2280
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   390
      Left            =   2280
      TabIndex        =   3
      Top             =   2280
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label Label5 
      Caption         =   "����˳�"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label sj 
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label cssd 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "ʣ�����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "����ע��:���뿪��ģʽ���ܻᷢ�����������,��ȷ�����������Ҫ����֮�������������,������󳬹�3��֮�󽫻��Զ�����Ϊ�������!"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "����(&P):"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cs As String
Dim ss As Integer
Dim passwd As String
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdOK_Click()
If txtPassword = passwd Then
Shell "explorer.exe " & Chr(34) & App.Path & "\resouce" & Chr(34), vbNormalFocus
kfms.Show
Unload Me
Else
If Val(cs) > 0 Then
cs = CStr(Val(cs) - 1)
cxx = MsgBox("�������,�������Գ���" & Val(cs) & "��,֮���������뽫����Ϊ�������!", vbOKOnly, "ע��!")
cssd.Caption = cs
Else
Name ".\none.cfgr" As ".\none.txt"
Open ".\none.txt" For Output As #2
Print #2, CStr(Int(Rnd() * 10000000 + 89999999))
Close #2
Name ".\none.txt" As ".\none.cfgr"
cxx = MsgBox("��������������,���������ѱ���Ϊ�������!", vbOKOnly, "ע��!")
Unload Me
End If
End If
End Sub
Private Sub Form_Load()
Randomize
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
cs = CStr(3)
ss = 30
End Sub

Private Sub Timer1_Timer()
If ss > 0 Then
ss = ss - 1
sj.Caption = CStr(ss)
Else
cxx = MsgBox("��ʱ!", vbOKOnly, "ע��!")
Unload Me
End If
End Sub
