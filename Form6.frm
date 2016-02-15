VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "会员信息批量修改"
   ClientHeight    =   3315
   ClientLeft      =   7785
   ClientTop       =   4080
   ClientWidth     =   3675
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3675
   Begin VB.Frame Frame2 
      Caption         =   "修改"
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   3375
      Begin VB.ComboBox Combo3 
         Height          =   300
         ItemData        =   "Form6.frx":014A
         Left            =   1080
         List            =   "Form6.frx":0163
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "清空"
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "Form6.frx":019D
         Left            =   1080
         List            =   "Form6.frx":01AD
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox bname2 
         Height          =   270
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "职务"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "部门"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "姓名"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "搜索条件"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   2400
         Picture         =   "Form6.frx":01D1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox bname1 
         Height          =   270
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form6.frx":2973
         Left            =   960
         List            =   "Form6.frx":2986
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "姓    名"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "会员状态"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
