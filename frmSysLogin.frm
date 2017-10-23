VERSION 5.00
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

Private Function mfVersionCheck() As Boolean
    '版本检查

    mfVersionCheck = True
    
End Function


Private Sub Command1_Click()
    Dim strName As String
    Dim frmNew As Form
    Dim I As Long
    
    strName = Trim(ucTC.Text)
    
    For I = 1 To 1
        Set frmNew = New frmSysTest
        frmNew.Caption = "Form" & I
        frmNew.Command1.Caption = frmNew.Caption & "cmd1"
        frmNew.Show
    Next
    
    SaveSetting gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserLast, strName
    Call msSaveUserList
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    If Not mfVersionCheck Then
        Unload Me
    End If
    
    Set Me.Icon = gMDI.Icon
    
    ucTC.Text = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserLast, "")
    Call msLoadUserList
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not gMDI.Visible Then
        Unload gMDI
    End If
End Sub
