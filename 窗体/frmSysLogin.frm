VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSysLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "系统登陆"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4005
   StartUpPosition =   2  '屏幕中心
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin 工程1.ucTextComboBox ucTC 
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
   End
   Begin VB.TextBox Text1 
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
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登陆"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "账号"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   580
      Width           =   360
   End
End
Attribute VB_Name = "frmSysLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mconDot As String = ","


Private Function mfVersionCheck() As Boolean
    '版本检查
    
    Dim fsoVer As FileSystemObject
    Dim arrNet() As String
    Dim arrLoc() As String
    Dim I As Long
    Dim blnNew As Boolean
    Dim strOut As String
    
    If Not gfFileExist(gID.FileAppNet) Then Exit Function   '网络上的文件是否存在
    
    On Error GoTo LineErr
    
    If GetAttr(gID.FileAppNet) <> vbNormal Then SetAttr gID.FileAppNet, vbNormal    '修改成正常属性
    If GetAttr(gID.FileAppLoc) <> vbNormal Then SetAttr gID.FileAppLoc, vbNormal    '
    
    Set fsoVer = New FileSystemObject
    arrNet = Split(fsoVer.GetFileVersion(gID.FileAppNet), ".")
    arrLoc = Split(fsoVer.GetFileVersion(gID.FileAppLoc), ".")
    For I = 0 To UBound(arrNet)
        If Val(arrNet(I)) > Val(arrLoc(I)) Then
            blnNew = True
            Exit For
        End If
    Next
    
    If blnNew Then
        If Not gfFileExist(gID.FileSetupNet) Then
            MsgBox "更新程序异常！", vbCritical
            Exit Function
        End If
        
        If GetAttr(gID.FileSetupNet) <> vbNormal Then SetAttr gID.FileSetupNet, vbNormal
        FileCopy gID.FileSetupNet, gID.FileSetupLoc
        Shell gID.FileSetupLoc
        
        Set gMDI = Nothing
        End
        
    End If
    
    Set fsoVer = Nothing
    mfVersionCheck = True
    
    Exit Function
    
LineErr:
    Call gsAlarmAndLog("版本检测异常")
    
End Function


Private Sub msLoadUserList()
    '加载曾经登陆过的用户名列表
    
    Dim strReg As String
    Dim strList() As String
    Dim strName As String
    Dim I As Long
    
    strReg = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserList, "")
    If Len(strReg) > 0 Then
        strList = Split(strReg, mconDot)
        For I = 0 To UBound(strList)
            strName = Trim(strList(I))  '防止注册表中信息被人为两边加入空格
            If Len(strName) > 0 Then ucTC.AddItem strName  '清理空格
        Next
    End If
    
End Sub

Private Sub msLoadUserAuthority(ByVal strUID As String)
    '权限控制
    
    Dim cbsAction As CommandBarAction
    Dim strSQL As String, strKey As String, strSys As String
    Const strFRM As String = "frm"
    
    strUID = Trim(strUID)
    If Len(strUID) = 0 Then Exit Sub
    
    strSys = LCase(gID.UserLoginName)
    If strSys = LCase(gID.UserAdmin) Or strSys = LCase(gID.UserSystem) Then   '程序内定两个用户拥有所有权限
        For Each cbsAction In gMDI.cBS.Actions
            cbsAction.Enabled = True
        Next
        Exit Sub
    End If
    
    strSQL = "SELECT DISTINCT t1.UserAutoID ,t1.UserLoginName ,t1.UserFullName " & _
             ",t5.FuncAutoID ,t5.FuncCaption ,t5.FuncName ,t5.FuncType " & _
             ",t6.FuncName AS [FuncFormName] FROM tb_Test_Sys_User AS [t1] " & _
             "INNER JOIN tb_Test_Sys_UserRole AS [t2] ON t1.UserAutoID =t2.UserAutoID " & _
             "INNER JOIN tb_Test_Sys_RoleFunc AS [t4] ON t2.RoleAutoID =t4.RoleAutoID " & _
             "INNER JOIN tb_Test_Sys_Func AS [t5] ON t4.FuncAutoID =t5.FuncAutoID " & _
             "INNER JOIN tb_Test_Sys_Func AS [t6] ON t5.FuncParentID =t6.FuncAutoID " & _
             "WHERE t1.UserAutoID =" & strUID
    Set gID.rsRF = gfBackRecordset(strSQL)
    With gID.rsRF
        If .State = adStateOpen Then
            If .RecordCount > 0 Then
                For Each cbsAction In gMDI.cBS.Actions
                    strKey = LCase(cbsAction.Key)
                    If Len(strKey) > 0 Then
                        If Left(strKey, 3) = strFRM Then
                            .MoveFirst
                            Do While Not .EOF
                                If LCase(.Fields("FuncName")) = strKey Then
                                    cbsAction.Enabled = True
                                End If
                                .MoveNext
                            Loop
                        End If
                    End If
                Next
            End If
        End If
    End With
    
End Sub

Private Sub msSaveUserList()
    '将当前用户名保存至登陆过的用户名列表中
    '并提升至列表的第一位，即表示越近登陆过的用户名越靠近列表前面
    
    Dim strName As String
    Dim strReg As String
    Dim strSave As String
    Dim strList() As String
    Dim I As Long
    
    strName = Trim(ucTC.Text)
    If Len(strName) = 0 Then Exit Sub
    strReg = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserList, "")
        
    If Len(strReg) > 0 Then
        strList = Split(strReg, mconDot)
        If LCase(strName) = LCase(Trim(strList(0))) Then Exit Sub
        
        For I = 0 To UBound(strList)    '此循环清理一下原来用户名两端的空格，并去掉此次登陆过的
            If LCase(strName) = LCase(Trim(strList(I))) Then
                strList(I) = ""
            Else
                strList(I) = Trim(strList(I))
            End If
        Next
        
        strSave = strName & mconDot '当前登陆的用户名保存在最前面
        For I = 0 To UBound(strList)    '将有效的用户名拼接起来
            If Len(strList(I)) > 0 Then strSave = strSave & strList(I) & mconDot
        Next
        
        strSave = Left(strSave, Len(strSave) - 1)   '去年最右边的mconDot
    Else
        strSave = strName
    End If
    
    SaveSetting gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserList, strSave
    
End Sub


Private Sub Command1_Click()
    Dim strName As String, strPWD As String
    Dim strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset
    
    
    strName = Trim(ucTC.Text)
    strPWD = Trim(Text1.Text)
    ucTC.Text = strName
    Text1.Text = strPWD
    
    If Len(strName) = 0 Then
        MsgBox "账号不能为空，且首尾不能有空格！", vbExclamation
        ucTC.SetFocus
        Exit Sub
    End If
    
    If Len(strPWD) = 0 Then
        MsgBox "密码不能为空，且首尾不能有空格！", vbExclamation
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strName)
    If Len(strMsg) > 0 Then
        MsgBox "账号中不能含有特殊字符【" & strMsg & "】！", vbExclamation
        ucTC.SetFocus
        Exit Sub
    End If
    
    strSQL = "EXEC sp_Test_Sys_UserLogin " & strName
    Set rsUser = gfBackRecordset(strSQL)
    
    If rsUser.State = adStateClosed Then GoTo LineEnd
    
    If rsUser.RecordCount = 0 Then
        strMsg = "账号不存在，请重新输入或联系管理员！"
        ucTC.SetFocus
        GoTo LineEnd
    End If
    
    If rsUser.RecordCount > 1 Then
        strMsg = "账号信息重复，禁止登陆，请联系管理员！"
        ucTC.SetFocus
        GoTo LineEnd
    End If
    
    If Not (LCase(strName) = LCase(gID.UserAdmin) Or LCase(strName) = LCase(gID.UserSystem)) Then
        If rsUser.Fields("UserState") & "" <> "启用" Then
            strMsg = "账号【" & strName & "】状态已停用，禁止登陆。启用请联系管理员！"
            ucTC.SetFocus
            GoTo LineEnd
        End If
    End If
    If gfDecryptSimple(rsUser.Fields("UserPassword") & "") <> strPWD Then
        strMsg = "密码输入错误！"
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        GoTo LineEnd
    End If
    
    gID.UserAutoID = rsUser.Fields("UserAutoID")
    gID.UserLoginName = strName
    gID.UserPassword = strPWD
    gID.UserFullName = rsUser.Fields("UserFullName") & ""
    gMDI.cBS.StatusBar.FindPane(gID.StatusBarPaneUserInfo).Text = gID.UserFullName
        
    rsUser.Close
    
    SaveSetting gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserLast, strName
    Call msSaveUserList
    Call gsLogAdd(Me, udSelect, "tb_Test_Sys_User", "【" & strName & "】登陆系统")
    Call msLoadUserAuthority(gID.UserAutoID) '******加载用户拥有的权限******
    
    gMDI.Show
    Unload Me
    
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
    If Len(strMsg) > 0 Then MsgBox strMsg, vbCritical
    
End Sub

Private Sub Form_Load()
    
    gID.UserComputerName = Winsock1.LocalHostName
    gID.UserLoginIP = Winsock1.LocalIP
   
    If App.PrevInstance Then
        If MsgBox("该程序在进程中存在、已经被打开！" & vbCrLf & vbCrLf _
            & "不建议开启多个相同程序端，是否仍要继续？", vbExclamation + vbYesNo) = vbNo Then
            Set gMDI = Nothing
            End
        End If
    End If
        
    If Not mfVersionCheck Then
        If MsgBox("软件版本检测失败！是否继续登陆？", vbExclamation + vbYesNo) = vbNo Then
            Set gMDI = Nothing
            End
        End If
    End If
    
    Set Me.Icon = gMDI.Icon
    
'''    ucTC.Text = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserLast, "")
    Call msLoadUserList
    If ucTC.ListCount > 0 Then ucTC.Text = ucTC.List(0)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not gMDI.Visible Then
        Unload gMDI
    End If
End Sub


Private Sub Text1_GotFocus()
    With Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
