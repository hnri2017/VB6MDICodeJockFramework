VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSysLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ϵͳ��½"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4005
   StartUpPosition =   2  '��Ļ����
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ����1.ucTextComboBox ucTC 
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��½"
      Default         =   -1  'True
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�"
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
    '�汾���
    
    Dim fsoVer As FileSystemObject
    Dim arrNet() As String
    Dim arrLoc() As String
    Dim I As Long
    Dim blnNew As Boolean
    Dim strOut As String
    
    If Not gfFileExist(gID.FileAppNet) Then Exit Function   '�����ϵ��ļ��Ƿ����
    
    On Error GoTo LineErr
    
    If GetAttr(gID.FileAppNet) <> vbNormal Then SetAttr gID.FileAppNet, vbNormal    '�޸ĳ���������
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
            MsgBox "���³����쳣��", vbCritical
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
    Call gsAlarmAndLog("�汾����쳣")
    
End Function


Private Sub msLoadUserList()
    '����������½�����û����б�
    
    Dim strReg As String
    Dim strList() As String
    Dim strName As String
    Dim I As Long
    
    strReg = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserList, "")
    If Len(strReg) > 0 Then
        strList = Split(strReg, mconDot)
        For I = 0 To UBound(strList)
            strName = Trim(strList(I))  '��ֹע�������Ϣ����Ϊ���߼���ո�
            If Len(strName) > 0 Then ucTC.AddItem strName  '����ո�
        Next
    End If
    
End Sub

Private Sub msLoadUserAuthority(ByVal strUID As String)
    'Ȩ�޿���
    
    Dim cbsAction As CommandBarAction
    Dim strSQL As String, strKey As String, strSys As String
    Const strFRM As String = "frm"
    
    strUID = Trim(strUID)
    If Len(strUID) = 0 Then Exit Sub
    
    strSys = LCase(gID.UserLoginName)
    If strSys = LCase(gID.UserAdmin) Or strSys = LCase(gID.UserSystem) Then   '�����ڶ������û�ӵ������Ȩ��
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
    '����ǰ�û�����������½�����û����б���
    '���������б�ĵ�һλ������ʾԽ����½�����û���Խ�����б�ǰ��
    
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
        
        For I = 0 To UBound(strList)    '��ѭ������һ��ԭ���û������˵Ŀո񣬲�ȥ���˴ε�½����
            If LCase(strName) = LCase(Trim(strList(I))) Then
                strList(I) = ""
            Else
                strList(I) = Trim(strList(I))
            End If
        Next
        
        strSave = strName & mconDot '��ǰ��½���û�����������ǰ��
        For I = 0 To UBound(strList)    '����Ч���û���ƴ������
            If Len(strList(I)) > 0 Then strSave = strSave & strList(I) & mconDot
        Next
        
        strSave = Left(strSave, Len(strSave) - 1)   'ȥ�����ұߵ�mconDot
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
        MsgBox "�˺Ų���Ϊ�գ�����β�����пո�", vbExclamation
        ucTC.SetFocus
        Exit Sub
    End If
    
    If Len(strPWD) = 0 Then
        MsgBox "���벻��Ϊ�գ�����β�����пո�", vbExclamation
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strName)
    If Len(strMsg) > 0 Then
        MsgBox "�˺��в��ܺ��������ַ���" & strMsg & "����", vbExclamation
        ucTC.SetFocus
        Exit Sub
    End If
    
    strSQL = "EXEC sp_Test_Sys_UserLogin " & strName
    Set rsUser = gfBackRecordset(strSQL)
    
    If rsUser.State = adStateClosed Then GoTo LineEnd
    
    If rsUser.RecordCount = 0 Then
        strMsg = "�˺Ų����ڣ��������������ϵ����Ա��"
        ucTC.SetFocus
        GoTo LineEnd
    End If
    
    If rsUser.RecordCount > 1 Then
        strMsg = "�˺���Ϣ�ظ�����ֹ��½������ϵ����Ա��"
        ucTC.SetFocus
        GoTo LineEnd
    End If
    
    If Not (LCase(strName) = LCase(gID.UserAdmin) Or LCase(strName) = LCase(gID.UserSystem)) Then
        If rsUser.Fields("UserState") & "" <> "����" Then
            strMsg = "�˺š�" & strName & "��״̬��ͣ�ã���ֹ��½����������ϵ����Ա��"
            ucTC.SetFocus
            GoTo LineEnd
        End If
    End If
    If gfDecryptSimple(rsUser.Fields("UserPassword") & "") <> strPWD Then
        strMsg = "�����������"
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
    Call gsLogAdd(Me, udSelect, "tb_Test_Sys_User", "��" & strName & "����½ϵͳ")
    Call msLoadUserAuthority(gID.UserAutoID) '******�����û�ӵ�е�Ȩ��******
    
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
        If MsgBox("�ó����ڽ����д��ڡ��Ѿ����򿪣�" & vbCrLf & vbCrLf _
            & "�����鿪�������ͬ����ˣ��Ƿ���Ҫ������", vbExclamation + vbYesNo) = vbNo Then
            Set gMDI = Nothing
            End
        End If
    End If
        
    If Not mfVersionCheck Then
        If MsgBox("����汾���ʧ�ܣ��Ƿ������½��", vbExclamation + vbYesNo) = vbNo Then
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
