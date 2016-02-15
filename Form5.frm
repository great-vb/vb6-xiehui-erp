VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "会员详细信息"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   5520
   ScaleWidth      =   8640
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo4 
      Height          =   300
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   3360
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "Form5.frx":1272
      Left            =   1080
      List            =   "Form5.frx":128B
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "保存并返回(&R)"
      Height          =   375
      Left            =   4200
      TabIndex        =   36
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "说明"
      Height          =   615
      Left            =   5280
      TabIndex        =   34
      Top             =   3840
      Width           =   3015
      Begin VB.Label Label15 
         Caption         =   "“元老”成员在职务栏填写原职位"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "会员状态"
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   3840
      Width           =   4695
      Begin VB.OptionButton Option3 
         Caption         =   "元老"
         Height          =   180
         Left            =   3600
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "退出"
         Height          =   180
         Left            =   1920
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "正常"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   4920
      TabIndex        =   9
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   4920
      TabIndex        =   8
      Top             =   2400
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form5.frx":12C5
      Left            =   4920
      List            =   "Form5.frx":12D5
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   4920
      TabIndex        =   7
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   4920
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "返回(&Home)"
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存并继续(&C)"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改(&E)"
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form5.frx":12F5
      Left            =   1080
      List            =   "Form5.frx":12FF
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label16 
      Caption         =   "年份"
      Height          =   255
      Left            =   6480
      TabIndex        =   38
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "必填"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   30
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "必填"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   29
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "必填"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   28
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "爱   好"
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "新浪UC"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "部   门"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "E-Mail"
      Height          =   255
      Left            =   4200
      TabIndex        =   24
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "MSN"
      Height          =   255
      Left            =   4440
      TabIndex        =   23
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "QQ"
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "手    机"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "专业班级"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "学    院"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "职    务"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "性    别"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "姓    名"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   7680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "会员详细信息"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()  '修改功能 ok at 11-10-12
    Set saveData = New ADODB.Recordset
    saveData.Open "select * from MemberShip where id=" & memid, conn, adOpenStatic, adLockOptimistic
    If saveData.RecordCount = 0 Then saveData.Close: MsgBox "发生不可预知的错误！", , "警告": Exit Sub
    saveData.MoveFirst
    With saveData
        .Fields("mName") = CNull(Text1.Text)
        .Fields("mSex") = CNull(Combo1.Text)
        .Fields("mJob") = CNull(Combo3.Text)
        .Fields("mMajor") = CNull(Text3.Text)
        .Fields("mClass") = CNull(Text4.Text)
        .Fields("mCellp") = CNull(Text5.Text)
        .Fields("mQQ") = CNull(Text6.Text)
        .Fields("mMsn") = CNull(Text7.Text)
        .Fields("mEmail") = CNull(Text8.Text)
        .Fields("mSinaUC") = CNull(Text9.Text)
        .Fields("mHobie") = CNull(Text10.Text)
        .Fields("mDepart") = Combo2.Text
        If Option1.Value = True Then
            .Fields("mState") = "正常"
        ElseIf Option2.Value = True Then
            .Fields("mState") = "退出"
        ElseIf Option3.Value = True Then
            .Fields("mState") = "元老"
        End If
        .Update
        .Close
    End With
    'obtion
    '对列表中数据进行修改
    With Form3.ListView1.ListItems(Form3.ListView1.SelectedItem.Index)
        .SubItems(1) = Text1 'name
        .SubItems(2) = Combo1.Text 'sex
        .SubItems(3) = Combo3 'job
        .SubItems(4) = Text3 'depart
        .SubItems(5) = Text4 'class
        .SubItems(6) = Text5 'cellphone
    End With
    Call initAll
    Form5.Hide
    Form3.Show
End Sub
Function initAll() 'ok at 11-10-12
    Text1 = ""
    Combo3.ListIndex = 3
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    Text7 = ""
    Text8 = ""
    Text9 = ""
    Text10 = ""
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 2
    Text1.SetFocus
End Function
Private Sub Command2_Click() '保存并继续 ok at 11-10-12
    '数据检察
    '**************
    If Text1 = "" Then
        MsgBox "必填！"
        Exit Sub
    Else
    End If
    '**************
    If Combo3 = "" Then
        MsgBox "必填！"
        Exit Sub
    Else
    End If
    '**************
    If Combo1.Text = "请选择" Then
        MsgBox "请选择性别！"
        Exit Sub
    Else
    End If
    '保存数据的代码
    '执行到此处说明必填项无误，开始执行kongkill
    Call kongkill  '将未填内容自动填充
    '开始进行保存
    Set saveData = New ADODB.Recordset
    saveData.Open "MemberShip", conn, adOpenStatic, adLockOptimistic
    With saveData
        .AddNew
        .Fields("mName") = CNull(Text1.Text)
        .Fields("mSex") = CNull(Combo1.Text)
        .Fields("mJob") = CNull(Combo3.Text)
        .Fields("mMajor") = CNull(Text3.Text)
        .Fields("mClass") = CNull(Text4.Text)
        .Fields("mCellp") = CNull(Text5.Text)
        .Fields("mQQ") = CNull(Text6.Text)
        .Fields("mMsn") = CNull(Text7.Text)
        .Fields("mEmail") = CNull(Text8.Text)
        .Fields("mSinaUC") = CNull(Text9.Text)
        .Fields("mHobie") = CNull(Text10.Text)
        .Fields("mDepart") = Combo2.Text
        .Fields("mAddYear") = Val(Combo4.Text)
        If Option1.Value = True Then
            .Fields("mState") = "正常"
        ElseIf Option2.Value = True Then
            .Fields("mState") = "退出"
        ElseIf Option3.Value = True Then
            .Fields("mState") = "元老"
        End If
        .Update
        .Close
    End With
    Call initAll
End Sub
Private Sub Command3_Click() '返回按钮 ok at 11-10-12
    Form5.Hide
    Form3.Show
    Call ShowNormal
End Sub
Private Sub Command4_Click() '添加并返回 ok at 11-10-12
    '数据检察
    '**************
    If Text1 = "" Then
        MsgBox "必填！"
        Exit Sub
    Else
    End If
    '**************
    If Combo3 = "" Then
        MsgBox "必填！"
        Exit Sub
    Else
    End If
    '**************
    If Combo1.Text = "请选择" Then
        MsgBox "请选择性别！"
        Exit Sub
    Else
    End If
    '保存数据的代码
    '执行到此处说明必填项无误，开始执行kongkill
    Call kongkill  '将未填内容自动填充
    '开始进行保存
    Set saveData = New ADODB.Recordset
    saveData.Open "MemberShip", conn, adOpenStatic, adLockOptimistic
    With saveData
        .AddNew
        .Fields("mName") = CNull(Text1.Text)
        .Fields("mSex") = CNull(Combo1.Text)
        .Fields("mJob") = CNull(Combo3.Text)
        .Fields("mMajor") = CNull(Text3.Text)
        .Fields("mClass") = CNull(Text4.Text)
        .Fields("mCellp") = CNull(Text5.Text)
        .Fields("mQQ") = CNull(Text6.Text)
        .Fields("mMsn") = CNull(Text7.Text)
        .Fields("mEmail") = CNull(Text8.Text)
        .Fields("mSinaUC") = CNull(Text9.Text)
        .Fields("mHobie") = CNull(Text10.Text)
        .Fields("mDepart") = Combo2.Text
        .Fields("mAddYear") = Val(Combo4.Text)
        If Option1.Value = True Then
            .Fields("mState") = "正常"
        ElseIf Option2.Value = True Then
            .Fields("mState") = "退出"
        ElseIf Option3.Value = True Then
            .Fields("mState") = "元老"
        End If
        .Update
        .Close
    End With
    Call initAll
    Me.Hide
    Form3.Show
    Call ShowNormal
End Sub
Private Sub Form_Load() 'ok at 11-10-11
    '根据信息判断是否为查看状态
    If opt = 0 Then
        Form5.Caption = "添加会员信息 - " & nowu
        Call allowall
    ElseIf opt = 1 Then
        Form5.Caption = "查看会员信息 - " & nowu
        Call disableall
    ElseIf opt = 2 Then
        Form5.Caption = "编辑会员信息 - " & nowu
        Call editmode
    Else
    End If
End Sub
Function disableall()  'for 查看 ok at 11-10-11
    Text1.Locked = True
    Combo3.Locked = True
    Text3.Locked = True
    Text4.Locked = True
    Text5.Locked = True
    Text6.Locked = True
    Text7.Locked = True
    Text8.Locked = True
    Text9.Locked = True
    Text10.Locked = True
    Combo1.Enabled = False
    Combo2.Enabled = False
    Combo3.Enabled = False
    Command1.Enabled = False '修改
    Command2.Enabled = False '保存并继续
    Command3.Enabled = True '返回
    Command4.Enabled = False '保存并返回
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
End Function
Function allowall()  'for 添加 ok at 11-10-11
    Text1.Locked = False
    Text3.Locked = False
    Text4.Locked = False
    Text5.Locked = False
    Text6.Locked = False
    Text7.Locked = False
    Text8.Locked = False
    Text9.Locked = False
    Text10.Locked = False
    Combo1.Enabled = True
    Combo2.Enabled = True
    Combo3.Enabled = True
    Command1.Enabled = False  '修改
    Command2.Enabled = True  '保存并继续
    Command3.Enabled = True  '返回
    Command4.Enabled = True  '保存并返回
    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
    Combo1.ListIndex = 0
    Combo2.ListIndex = 2
    Combo3.ListIndex = 0
End Function
Function editmode() 'for 修改内容 ok at 11-10-11
    Text1.Locked = False
    Combo3.Locked = False
    Text3.Locked = False
    Text4.Locked = False
    Text5.Locked = False
    Text6.Locked = False
    Text7.Locked = False
    Text8.Locked = False
    Text9.Locked = False
    Text10.Locked = False
    Combo1.Enabled = True
    Combo2.Enabled = True
    Combo3.Enabled = True
    Command1.Enabled = True '修改
    Command2.Enabled = False '保存并继续
    Combo1.ListIndex = 0
    Command3.Enabled = True '返回
    Command4.Enabled = False '保存并返回
    Option1.Enabled = True
    Option2.Enabled = True
    Option3.Enabled = True
End Function
Function kongkill()  '将空的填上内容 ok at 11-10-11
    '*************
    If Text3 = "" Then
        Text3 = "暂无"
    Else
    End If
    
    If Text4 = "" Then
        Text4 = "暂无"
    Else
    End If
    If Text5 = "" Then
        Text5 = "暂无"
    Else
    End If
    
    If Text6 = "" Then
        Text6 = "暂无"
    Else
    End If
    '***********
    If Text7 = "" Then
        Text7 = "暂无"
    Else
    End If
    '***********
    If Text8 = "" Then
        Text8 = "暂无"
    Else
    End If
    '***********
    If Text9 = "" Then
        Text9 = "暂无"
    Else
    End If
    '***********
    If Text10 = "" Then
        Text10 = "暂无"
    Else
    End If
    '***********
    If Combo2.Text = "" Then
        Combo2.Text = "暂无"
    Else
    End If
    '***********
End Function
Function csh()  '根据选项，调整窗体的功能  'ok at 11-10-12
    If opt = 0 Then
        Form5.Caption = "添加会员信息 - " & nowu
        Command1.Enabled = False
        Command2.Enabled = True
        Command3.Enabled = True
        Option1.Value = True
        Option2.Value = False
        Option3.Value = False
        Text1 = ""
        Combo3.ListIndex = 2
        Text3 = ""
        Text4 = ""
        Text5 = ""
        Text6 = ""
        Text7 = ""
        Text8 = ""
        Text9 = ""
        Text10 = ""
        Combo1.ListIndex = 0
        Combo2.ListIndex = 0
        Combo3.ListIndex = 2
        'Ready for Years
        Combo4.Clear
        For i = Year(Date) - 5 To Year(Date)
        Combo4.AddItem i
        Next i
        Call allowall
    ElseIf opt = 1 Then
        Form5.Caption = "查看会员信息 - " & nowu
        Call disableall
    ElseIf opt = 2 Then
        Form5.Caption = "编辑会员信息 - " & nowu
        Call editmode
    Else
    End If
End Function
Private Sub Form_Unload(Cancel As Integer)  '关闭此编辑、查看、添加窗口时，退回到总查看窗口 11-10-12
    'ok at 11-10-12
    Form3.Show
End Sub

