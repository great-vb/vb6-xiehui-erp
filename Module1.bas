Attribute VB_Name = "Module1"
Public users, i, iss, nowid, count1, lists As Long 'users记录用户个数,i作为循环量,nowid作为登录后的用户号;count1 is number of members
Public user1, pass1, ques1, ans1 As String
Public tmp As String
Public wstr As String
Public opt As Integer '操作状态：0为添加，1为查看，2为修改
Public memid As Long '当前处理的会员Id
Public Type user
  uName As String
  uPass As String
  uQues As String
  uAnsw As String
End Type
'All new dims
Public res As ADODB.Recordset
Public veryId As ADODB.Recordset
Public saveData As ADODB.Recordset
Public getData As ADODB.Recordset
Public conn As ADODB.Connection
Public nowu As String
Public admin As user
Public sqlLine As String
Sub Main()   'Changed -by 2011年10月8日
    If Dir(App.Path & "\member.mdb") = "" Then 'file does not exist
        Call OutputFileS(101, App.Path & "\member.mdb")  'release file
        MsgBox "不存在数据文件或数据文件损坏，系统自动恢复初始状态" & vbCrLf & "默认密码：admin", , "警告"
    Else
    End If
    'Creat Data Base Connection
    Call OpenTable(App.Path & "\member.mdb")
    DoEvents
    Load Form1  '登陆界面
    Load Form2
    Load Form3
    Load Form4
    Load Form5
    Form2.Timer1.Enabled = False
    Form1.Show
    Call loadMeuser '将用户名列表填充
End Sub
Function loadMeuser() '代码完成，不需改变即可适应会员管理需要-by 2011年10月8日
Form1.Combo1.Clear
Set res = New ADODB.Recordset
res.Open "users", conn, adOpenStatic, adLockOptimistic
If res.RecordCount = 0 Then res.Close: MsgBox "请创建新用户": Form1.Command3.Enabled = True: Exit Function
res.MoveFirst
        Do While Not res.EOF = True
            Form1.Combo1.AddItem res.Fields("userName")
            res.MoveNext
        Loop
res.Close
        Form1.Command3.Enabled = False
        Form1.Combo1.ListIndex = 0
End Function


Function refreshlogin()          '代码完成，适应会员管理：向列表中重新注入数据 ok at 11-10-08
    Form1.Combo1.Clear
    Set res = New ADODB.Recordset
    res.Open "users", conn, adOpenStatic, adLockOptimistic
    If res.RecordCount = 0 Then res.Close: Form1.Command3.Enabled = True: Exit Function
        Do While Not res.EOF = True
            Form1.Combo1.AddItem res.Fields("userName")
            res.MoveNext
        Loop
    res.Close
End Function

Function ShowAtSchool()   'ok at 11-10-08
Call ShowBySql("select * from MemberShip where mAddYear>=" & CStr(Year(Date) - 3) & " and mState='正常'")
End Function

Function ShowAllMem() 'ok at 11-10-08
Call ShowBySql("MemberShip")
End Function
Function GetUserInfo(ByVal uName As String, inuser As user) As Boolean
    Set res = New ADODB.Recordset
    res.Open "select * from users where userName='" & uName & "'", conn, adOpenStatic, adLockOptimistic
    If res.RecordCount = 0 Then res.Close: GetUserInfo = False: Exit Function
        res.MoveFirst
        inuser.uName = uName
        inuser.uPass = res.Fields("passWord")
        inuser.uAnsw = res.Fields("passAnswer")
        inuser.uQues = res.Fields("passQuestion")
    res.Close
    GetUserInfo = True
End Function
Function ShowNormal()  'ok at 11-10-08
Call ShowBySql("select * from MemberShip where mState='正常'")
End Function
Function loadmemfor(ByVal id As Long) As Boolean       'load information when edit,view,or input information-by 11-10-08
Set res = New ADODB.Recordset
res.Open "select * from MemberShip where id=" & id, conn, adOpenStatic, adLockOptimistic
If res.RecordCount = 0 Then res.Close: loadmemfor = False
    With Form5
        .Text1 = CNull(res.Fields("mName"))
        Select Case CNull(res.Fields("mSex"))
        Case "男"
        .Combo1.ListIndex = 0
        Case "女"
        .Combo1.ListIndex = 1
        End Select
        For i = 0 To .Combo3.ListCount - 1
        If .Combo3.List(i) = CNull(res.Fields("mJob")) Then
            .Combo3.ListIndex = i
            Exit For
        End If
        Next i
        .Text3 = CNull(res.Fields("mMajor"))
        .Text4 = CNull(res.Fields("mClass"))
        .Text5 = CNull(res.Fields("mCellp"))
        .Text6 = CNull(res.Fields("mQQ"))
        .Text7 = CNull(res.Fields("mMsn"))
        .Text8 = CNull(res.Fields("mEmail"))
        .Text9 = CNull(res.Fields("mSinaUC"))
        .Text10 = CNull(res.Fields("mHobie"))
        .Combo4.AddItem res.Fields("mAddYear")
        For i = 0 To .Combo2.ListCount - 1
        If .Combo2.List(i) = CNull(res.Fields("mDepart")) Then
            .Combo2.ListIndex = i
            Exit For
        End If
        Next i
        Select Case CNull(res.Fields("mState"))
            Case "正常"
                .Option1.Value = True
            Case "退出"
                .Option2.Value = True
            Case "元老"
                .Option3.Value = True
            Case Else
                MsgBox CNull(res.Fields("mState"))
                MsgBox "状态有误，默认设定为“正常”，请校正数据！"
                .Option1.Value = True
        End Select
    End With
res.Close
End Function

Function OpenTable(ByVal txtPath As String) '【功能：建立数据库连接；状态：完成】'ok at 11-10-07
Set conn = New ADODB.Connection
conn.CursorLocation = adUseClient
conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtPath & ";"
End Function
Function CloseTable() '【功能：关闭数据库连接；状态：完成】'ok at 11-10-07
conn.Close
End Function

Function OutputFileS(ByVal sId As Long, ByVal sFile As String) 'ok at 11-10-07
Dim sTemp() As Byte
sTemp = LoadResData(sId, "CUSTOM")
Open sFile For Binary As #1
        Put #1, , sTemp
Close #1
End Function

Function CNull(ByVal sTxt As Variant) As String   'ok at 11-10-08
If IsNull(sTxt) = True Then
    CNull = ""
Else
    If Len(sTxt) > 30 Then
    CNull = Mid(sTxt, 1, 30)
    Else
    CNull = sTxt
    End If
End If
End Function

Function sqlRemove(ByVal idOf As Long)   '根据id删除会员 ok at 11-10-11
Set res = New ADODB.Recordset
res.Open "delete * from MemberShip where id=" & idOf, conn, adOpenStatic, adLockOptimistic
End Function

Function ShowBySql(ByVal sqlStr As String) 'ok at 11-10-12
    Form3.ListView1.ListItems.Clear
    Dim im As Long
    Set res = New ADODB.Recordset
    res.Open sqlStr, conn, adOpenStatic, adLockOptimistic
    If res.RecordCount = 0 Then res.Close: Exit Function
        Do While Not res.EOF = True
            With Form3.ListView1.ListItems.Add()
                .Text = CNull(res.Fields("id"))
                .SubItems(1) = CNull(res.Fields("mName"))
                .SubItems(2) = CNull(res.Fields("mSex"))
                .SubItems(3) = CNull(res.Fields("mJob"))
                .SubItems(4) = CNull(res.Fields("mMajor"))
                .SubItems(5) = CNull(res.Fields("mClass"))
                .SubItems(6) = CNull(res.Fields("mCellp"))
            End With
        res.MoveNext
        Loop
    res.Close
End Function

Function SelectFile(ByVal sPath As String)  'ok at 11-10-12
Dim k As Long
k = Shell("C:\WINDOWS\explorer.exe /n ,/select ," & sPath, vbNormalFocus)
End Function

Function LenBS(ByVal strTxt As String)
Dim i As Long, c As Long
For i = 1 To Len(strTxt)
If Asc(Mid(strTxt, i, 1)) < 0 Then
c = c + 2
Else
c = c + 1
End If
Next i
LenBS = c
End Function
