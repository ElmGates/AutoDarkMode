VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "更改时间"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1770
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3255
   ScaleWidth      =   1770
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定更改"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox jsss 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Text            =   "18"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox kstt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Text            =   "07"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "请注意,请您输入一个两位数,24小时制,例如上午8点请输入08,下午1点请输入13."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "结束:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "开始:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kssj As String
Dim jssj As String
Dim zhsj As String
Dim usersetday As String
Dim usersetnight As String
Dim userconfig As String
Dim dayt As Integer
Dim nightt As Integer
Private Sub Command1_Click()
zhsj = ""
kssj = kstt
If Len(kssj) <> 2 Then
pppp = MsgBox("您输入的内容有误,请再试一次!", vbOKOnly, "输入错误")
Else
If Val(kssj) < 24 And Val(kssj) > 0 Then
Else
pppp = MsgBox("您输入的内容有误,请再试一次!", vbOKOnly, "输入错误")
End If
End If
zhsj = zhsj + kssj
jssj = jsss
If Len(jssj) <> 2 Then
pppp = MsgBox("您输入的内容有误,请再试一次!", vbOKOnly, "输入错误")
Else
If Val(jssj) < 24 And Val(jssj) > 0 Then
Else
pppp = MsgBox("您输入的内容有误,请再试一次!", vbOKOnly, "输入错误")
End If
End If
zhsj = zhsj + jssj
Name ".\userconfig.inconfig" As ".\userconfig.txt"
Open ".\userconfig.txt" For Output As #1
Print #1, zhsj
Close #1
Name ".\userconfig.txt" As ".\userconfig.inconfig"
pppppp = MsgBox("返回主程序之后需要您手动点击刷新设置按钮,否则设置将在下次启动程序时才生效!", vbononly, "重要提示")
Unload Me
End Sub
Private Sub Form_Load()
Name ".\userconfig.inconfig" As ".\userconfig.txt"
Open ".\userconfig.txt" For Input As #2
Line Input #2, userconfig
Close #2
usersetday = Mid(userconfig, 1, 2)
usersetnight = Mid(userconfig, 3, 2)
Name ".\userconfig.txt" As ".\userconfig.inconfig"
kstt.Text = usersetday
jsss.Text = usersetnight
End Sub
