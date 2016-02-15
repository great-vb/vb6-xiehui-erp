VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "登录平台"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   3060
   ScaleWidth      =   4875
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "忘记密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "新用户"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Picture         =   "Form1.frx":366A
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1185
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "欢 迎 使 用 会 员 管 理 系 统"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "密  码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()  '登陆按钮 ok at 11-10-11
    '初始化
    On Error Resume Next
    Call GetUserInfo(Combo1.Text, admin)
    '判断是否获得用户名的真实性
    If Combo1.Text = "" Then MsgBox "用户不存在，请创建！": Exit Sub
    If admin.uPass = Text2.Text Then
        '加载用户信息
        nowu = admin.uName
        Load Form3
        Form3.Caption = "协会管理系统欢迎" & nowu & "使用"
        Call ShowAllMem
        Form1.Hide
        
        Form3.Show
    Else
        MsgBox "密码错误，请重新输入！"
    End If
End Sub
Private Sub Command2_Click() '清空按钮 -ok at 11-10-08
    Combo1.Text = ""
    Text2 = ""
    Combo1.SetFocus
End Sub
Private Sub Command3_Click() '新用户按钮 -ok at 11-10-08
    Form1.Hide
    Form2.Show
    Form2.Timer1.Enabled = True
End Sub
Private Sub Command4_Click() '忘记密码按钮 ok at 11-10-08
    Form1.Hide
    Form4.Show
End Sub



Private Sub Form_Activate()
    'MsgBox "欢迎使用计算机协会会员管理系统！此版本为最新版，但不含权限管理功能。敬请谅解！" & vbCrLf & "我加入了搜索、导出名单等新功能，希望使用者喜欢！"
    If Dir(App.Path & "\datas.sun") <> "" Then
    Else
        Exit Sub
    End If
End Sub

Private Sub Form_Load() 'ok at 11-10-08

End Sub
Private Sub Form_Unload(Cancel As Integer) 'ok at 11-10-11
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form5
    Call CloseTable  '断开连接
    End
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Command1_Click
End Sub
