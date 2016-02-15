VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "文件管理大师-密码找回"
   ClientHeight    =   2295
   ClientLeft      =   9600
   ClientTop       =   5010
   ClientWidth     =   3795
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   2295
   ScaleWidth      =   3795
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3360
      Top             =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "找回密码"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "密码答案"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "密码问题"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click() 'ok at 11-10-11
  If Text1 <> "" And Text2 <> "" And Text3 <> "" Then
       If Text3 = admin.uAnsw Then
       MsgBox "您的密码是：" & admin.uPass
       Command2_Click '清空
       Form1.Text2 = admin.uPass
       Else
       MsgBox "答案错误！"
       Text3 = ""
       Text3.SetFocus
       End If
  Else
  MsgBox "请填写正确的用户名和答案信息！"
  End If
End Sub

Private Sub Command2_Click() '清空 ok at 11-10-11
Text1 = "": Text3 = ""
Text1.SetFocus
End Sub

Private Sub Command3_Click() '返回 ok at 11-10-11
Command2_Click  '清空
Form4.Hide '隐藏本页面
Form1.Show '显示登陆界面
End Sub

Private Sub Timer1_Timer() 'ok at 11-10-11
If GetUserInfo(Text1.Text, admin) = True Then
  Text2.Text = admin.uQues
Else
    Text2.Text = ""
    End If

End Sub
