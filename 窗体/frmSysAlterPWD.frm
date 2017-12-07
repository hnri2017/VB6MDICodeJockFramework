VERSION 5.00
Begin VB.Form frmSysAlterPWD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "密码修改"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4875
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   300
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   300
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "修改"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "密码只能包含数字或大小字母，且长度在20位以内"
      ForeColor       =   &H000000FF&
      Height          =   1020
      Index           =   5
      Left            =   3480
      TabIndex        =   12
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新密码确认"
      Height          =   180
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Top             =   1850
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "新密码"
      Height          =   180
      Index           =   3
      Left            =   840
      TabIndex        =   10
      Top             =   1500
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "旧密码"
      Height          =   180
      Index           =   2
      Left            =   840
      TabIndex        =   9
      Top             =   1150
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "账号"
      Height          =   180
      Index           =   1
      Left            =   1000
      TabIndex        =   8
      Top             =   765
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   0
      Left            =   1000
      TabIndex        =   7
      Top             =   400
      Width           =   360
   End
End
Attribute VB_Name = "frmSysAlterPWD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    '修改密码
    
    Dim strULID As String, strName As String
    Dim strPwdPre As String, strPwdNew As String, strPwdNewB As String
    Dim strSQL As String, strMsg As String
    Dim rsEdit As ADODB.Recordset
    Dim K As Long
    
    For K = 2 To 4
        If Len(Trim(Text1.Item(K).Text)) = 0 Then
            MsgBox "请输入【" & Label1.Item(K).Caption & "】且不可以有空格！", vbExclamation, "空密码警告"
            Text1.Item(K).SetFocus
            Exit Sub
        End If
    Next
    
    strName = Trim(Text1.Item(0).Text)
    strULID = Trim(Text1.Item(1).Text)
    strPwdPre = Trim(Text1.Item(2).Text)
    strPwdNew = Trim(Text1.Item(3).Text)
    strPwdNewB = Trim(Text1.Item(4).Text)
    Text1.Item(2).Text = strPwdPre
    Text1.Item(3).Text = strPwdNew
    Text1.Item(4).Text = strPwdNewB
    
    If strPwdNew <> strPwdNewB Then
        MsgBox "两次输入的新密码不一致！", vbExclamation
        Text1.Item(4).SetFocus
        Text1.Item(4).SelStart = 0
        Text1.Item(4).SelLength = Len(strPwdNewB)
        Exit Sub
    End If
    
    If strPwdNew = strPwdPre Then
        MsgBox "新密码不能与旧密码相同！", vbExclamation
        Text1.Item(3).SetFocus
        Text1.Item(3).SelStart = 0
        Text1.Item(3).SelLength = Len(strPwdNew)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strPwdPre)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(2).Caption & " 不能含有特殊字符【" & strMsg & "】！", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPwdPre)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strPwdNew)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(3).Caption & " 不能含有特殊字符【" & strMsg & "】！", vbExclamation
        Text1.Item(3).SetFocus
        Text1.Item(3).SelStart = 0
        Text1.Item(3).SelLength = Len(strPwdNew)
        Exit Sub
    End If
    
    strSQL = "SELECT UserLoginName ,UserPassword ,UserFullName FROM tb_Test_Sys_User " & _
             "WHERE UserLoginName='" & strULID & "'"

    Set rsEdit = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsEdit.State = adStateClosed Then GoTo LineEnd
    If rsEdit.RecordCount = 0 Then
        strMsg = "用户信息丢失！请联系系统管理员。"
        GoTo LineEnd
    End If
    If rsEdit.RecordCount > 1 Then
        strMsg = "用户信息重复！请联系系统管理员。"
        GoTo LineEnd
    End If
    If rsEdit.Fields("UserFullName") <> strName Then
        strMsg = "用户信息已被变更，请关闭该窗口后重新打开！"
        GoTo LineEnd
    End If
    If gfDecryptSimple(rsEdit.Fields("UserPassword")) <> strPwdPre Then '解密数据库中的密码
        strMsg = "旧密码输入错误，请重新输入！"
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPwdPre)    '简单加密 密码再存储
        GoTo LineEnd
    End If
    
    On Error GoTo LineErr
    
    If MsgBox("确定修改当前密码吗？", vbQuestion + vbOKCancel, "密码修改询问") = vbOK Then
        rsEdit.Fields("UserPassword") = gfEncryptSimple(strPwdNew)
        rsEdit.Update
        rsEdit.Close
        Call gsLogAdd(Me, udUpdate, "tb_Test_Sys_User", "修改用户ID【" & strULID & "】的登陆密码")
        MsgBox "密码已修改成功，重新登陆后生效。", vbInformation
    End If
    
    GoTo LineEnd
    
LineErr:
    Call gsAlarmAndLog("密码修改异常")
LineEnd:
    If rsEdit.State = adStateOpen Then rsEdit.Close
    Set rsEdit = Nothing
    If Len(strMsg) > 0 Then MsgBox strMsg, vbCritical
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '
    Set Me.Icon = gMDI.imgListCommandBars.ListImages("SysPassword").Picture
    Text1.Item(0).Text = gID.UserFullName
    Text1.Item(1).Text = gID.UserLoginName

End Sub
