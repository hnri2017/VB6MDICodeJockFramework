VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysUser 
   Caption         =   "�û�����"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13215
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   13215
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6255
      Left            =   4080
      TabIndex        =   12
      Top             =   120
      Width           =   8175
      Begin VB.OptionButton Option1 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   2040
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ů"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   5
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   4
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   120
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   4680
         TabIndex        =   11
         Text            =   "Combo2"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����û�"
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�޸��û���Ϣ"
         Height          =   495
         Left            =   3240
         TabIndex        =   10
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "����ֻ�ܰ������ֻ��С��ĸ���ҳ�����20λ����"
         ForeColor       =   &H000000FF&
         Height          =   420
         Index           =   7
         Left            =   4680
         TabIndex        =   20
         Top             =   1080
         Width           =   2100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ע"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   650
         TabIndex        =   19
         Top             =   3060
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   650
         TabIndex        =   18
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   650
         TabIndex        =   17
         Top             =   2100
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   650
         TabIndex        =   16
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "��ʶ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   650
         TabIndex        =   15
         Top             =   180
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�˺�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   650
         TabIndex        =   14
         Top             =   660
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   650
         TabIndex        =   13
         Top             =   1140
         Width           =   450
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7223
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSysUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlngID As Long

Private Const mKeyDept As String = "k"
Private Const mKeyUser As String = "u"
Private Const mOtherKey As String = "kOther"
Private Const mOtherText As String = "������Ա"



Private Sub msLoadDept(ByRef tvwDept As MSComctlLib.TreeView)
    '���ز�����TreeView�ؼ���
    'Ҫ��1�����ݿ��в�����Ϣ��Dept����DeptID(Not Null)��DeptName(Not Null)��ParentID(Null)�����ֶΡ�
    'Ҫ��2�����ű���ֻ���������ţ�����Ϊ��˾���ƣ���ParentIDΪNull���������ŵ�ParentID������ΪNull��
    
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim arrDept() As String 'ע���±�Ҫ��0��ʼ
    Dim I As Long, lngCount As Long, lngOneCompany As Long
    Dim blnLoop As Boolean
    
    
    strSQL = "SELECT t1.DeptID ,t1.DeptName ,t1.ParentID ,t2.DeptName AS [ParentName] " & _
             "FROM tb_Test_Sys_Department AS [t1] " & _
             "LEFT JOIN tb_Test_Sys_Department AS [t2] " & _
             "ON t1.ParentID = t2.DeptID " & _
             "ORDER BY t1.ParentID ,t1.DeptName"    'ע���ֶ�˳�򲻿ɱ�
    Set rsDept = gfBackRecordset(strSQL)
    If rsDept.State = adStateClosed Then Exit Sub
    If rsDept.RecordCount > 0 Then
        
        tvwDept.Nodes.Clear
        Combo1.Item(0).Clear
        Combo1.Item(1).Clear
        
        While Not rsDept.EOF
            If IsNull(rsDept.Fields(3).Value) Then
                lngOneCompany = lngOneCompany + 1
                tvwDept.Nodes.Add , , mKeyDept & rsDept.Fields(0).Value, rsDept.Fields(1).Value, "SysCompany"
                tvwDept.Nodes.Item(mKeyDept & rsDept.Fields(0).Value).Expanded = True
            Else
                ReDim Preserve arrDept(3, lngCount)
                For I = 0 To 3
                    arrDept(I, lngCount) = rsDept.Fields(I).Value
                Next
                lngCount = lngCount + 1
                blnLoop = True
            End If
            
            Combo1.Item(0).AddItem rsDept.Fields(1).Value
            Combo1.Item(1).AddItem rsDept.Fields(0).Value
            
            rsDept.MoveNext
        Wend
        
    End If
    rsDept.Close
    Set rsDept = Nothing
    
    If blnLoop Then Call msLoadDeptTree(tvwDept, arrDept)

End Sub

Private Sub msLoadDeptTree(ByRef tvwTree As MSComctlLib.TreeView, ByRef arrLoad() As String)
    '������msLoadDept�������ʹ�������ز����б�
    
    Dim arrOther() As String    '����ʣ���
    Dim blnOther As Boolean     'ʣ���ʶ
    Dim I As Long, J As Long, K As Long, lngCount As Long
    Static C As Long
    
    With tvwTree
        For J = LBound(arrLoad, 2) To UBound(arrLoad, 2)
            For I = 1 To .Nodes.Count   'ע��˴��±��1��ʼ
                If .Nodes.Item(I).Key = mKeyDept & arrLoad(2, J) Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, mKeyDept & arrLoad(0, J), arrLoad(1, J), "threemen"
                    .Nodes.Item(mKeyDept & arrLoad(0, J)).Expanded = True
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(3, lngCount)
                For K = 0 To 3
                    arrOther(K, lngCount) = arrLoad(K, J)
                Next
                lngCount = lngCount + 1
            End If
            
        Next
    End With
    
    C = C + 1
    If C > 64 Then Exit Sub '��ֹ�ݹ����̫��¶�ջ������������
    
    If blnOther Then
        Call msLoadDeptTree(tvwTree, arrOther)
    End If

End Sub

Private Sub msLoadUser(ByRef tvwUser As MSComctlLib.TreeView)
    '�����û���ǰ�����Ѽ��غò���
    
    Dim strSQL As String
    Dim rsUser As ADODB.Recordset
    Dim arrOther() As String    '����ʣ���
    Dim blnOther As Boolean     'ʣ���ʶ
    Dim I As Long, J As Long, K As Long, lngCount As Long
    
    If tvwUser.Nodes.Count = 0 Then Exit Sub
    
    strSQL = "SELECT UserAutoID ,UserFullName ,UserSex ,DeptID FROM tb_Test_Sys_User"
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then GoTo LineEnd

    With tvwUser
        While Not rsUser.EOF
            For I = 1 To .Nodes.Count
                If .Nodes(I).Key = mKeyDept & rsUser.Fields("DeptID").Value Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, _
                        mKeyUser & rsUser.Fields("UserAutoID").Value, rsUser.Fields("UserFullName").Value, _
                        IIf(rsUser.Fields("UserSex") = "��", "man", "woman")
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(3, lngCount)
                For K = 0 To 3
                    arrOther(K, lngCount) = rsUser.Fields(K).Value & ""
                Next
                lngCount = lngCount + 1
            End If
            
            rsUser.MoveNext
        Wend
        
        If blnOther Then
            .Nodes.Add 1, tvwChild, mOtherKey, mOtherText, "unknown"
            .Nodes(mOtherKey).Expanded = True
            For I = LBound(arrOther, 2) To UBound(arrOther, 2)
                .Nodes.Add mOtherKey, tvwChild, mKeyUser & arrOther(0, I), _
                    arrOther(1, I), IIf(arrOther(2, I) = "��", "man", "woman")
            Next
        End If

    End With
    
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
End Sub


Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            Combo1.Item(Index).ListIndex = -1
        End If
    End If
End Sub

Private Sub Command1_Click()
    '���
    
    Dim strLoginName As String, strPWD As String, strFullName As String
    Dim strSex As String, strMemo As String
    Dim strDept As Variant
    Dim strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset
    
    strLoginName = Trim(Text1.Item(1).Text)
    strPWD = Trim(Text1.Item(2).Text)
    strFullName = Trim(Text1.Item(3).Text)
    strMemo = Trim(Text1.Item(4).Text)
    
    strLoginName = Left(strLoginName, 50)
    strPWD = Left(strPWD, 20)
    strFullName = Left(strFullName, 50)
    strMemo = Left(strMemo, 500)
    
    Text1.Item(1).Text = strLoginName
    Text1.Item(2).Text = strPWD
    Text1.Item(3).Text = strFullName
    Text1.Item(4).Text = strMemo
    
    If Option1.Item(0).Value Then strSex = Option1.Item(0).Caption
    If Option1.Item(1).Value Then strSex = Option1.Item(1).Caption
    strDept = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    If Len(strLoginName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strLoginName)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strLoginName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " ���ܺ��������ַ���" & strMsg & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strLoginName)
        Exit Sub
    End If
    
    If Len(strPWD) = 0 Then
        MsgBox Label1.Item(2).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strPWD)
        Exit Sub
    End If
    
    If Len(strFullName) = 0 Then
        MsgBox Label1.Item(3).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(3).SetFocus
        Text1.Item(3).SelStart = 0
        Text1.Item(3).SelLength = Len(strFullName)
        Exit Sub
    End If
    
    If Len(strSex) = 0 Then Option1.Item(0).Value = True
    
    If Len(strDept) = 0 Then strDept = Null
    
    If MsgBox("�Ƿ�����û���" & strLoginName & "����" & strFullName & "����", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    strSQL = "SELECT UserAutoID ,UserLoginName ,UserPassword ," & _
             "UserFullName ,UserSex ,DeptID ,UserMemo " & _
             "From tb_Test_Sys_User " & _
             "WHERE UserLoginName = '" & strLoginName & "'"
    Set rsUser = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount > 0 Then
        strMsg = "�˺��Ѵ��ڣ��������"
        GoTo LineBrk
    Else
        On Error GoTo LineErr
        
        rsUser.AddNew
        rsUser.Fields("UserLoginName") = strLoginName
        rsUser.Fields("UserPassword") = gfEncryptSimple(strPWD)
        rsUser.Fields("UserFullName") = strFullName
        rsUser.Fields("UserSex") = strSex
        rsUser.Fields("DeptID") = strDept
        rsUser.Fields("UserMemo") = strMemo
        rsUser.Update
        strMsg = rsUser.Fields("UserAutoID").Value
        Text1.Item(0).Text = strMsg
        rsUser.Close
        strMsg = "����û���" & strMsg & "����" & strLoginName & "����" & strFullName & "��"
        Call gsLogAdd(Me, udInsert, "tb_Test_Sys_User", strMsg)
        MsgBox "�û���" & strLoginName & "����" & strFullName & "����ӳɹ���", vbInformation
        Call msLoadDept(TreeView1)
        Call msLoadUser(TreeView1)
    End If
    
    GoTo LineEnd
    
LineBrk:
    rsUser.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineErr:
    Call gsAlarmAndLog("����û��쳣")
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
End Sub

Private Sub Command2_Click()
    '�޸�
    
    Dim strUID As String, strLoginName As String, strPWD As String
    Dim strFullName As String, strSex As String, strDept As String, strMemo As String
    Dim blnLoginName As Boolean, blnPwd As Boolean, blnFullName As Boolean
    Dim blnSex As Boolean, blnDept As Boolean, blnMemo As Boolean
    Dim strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset
    
    strUID = Trim(Text1.Item(0).Text)
    strLoginName = Trim(Text1.Item(1).Text)
    strPWD = Trim(Text1.Item(2).Text)
    strFullName = Trim(Text1.Item(3).Text)
    strMemo = Trim(Text1.Item(4).Text)
    
    strLoginName = Left(strLoginName, 50)
    strPWD = Left(strPWD, 20)
    strFullName = Left(strFullName, 50)
    strMemo = Left(strMemo, 500)
    
    Text1.Item(1).Text = strLoginName
    Text1.Item(2).Text = strPWD
    Text1.Item(3).Text = strFullName
    Text1.Item(4).Text = strMemo
    
    If Option1.Item(0).Value Then strSex = Option1.Item(0).Caption
    If Option1.Item(1).Value Then strSex = Option1.Item(1).Caption
    strDept = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    If Len(strLoginName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strLoginName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " ���ܺ��������ַ���" & strMsg & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    If Len(strPWD) = 0 Then
        MsgBox Label1.Item(2).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(Text1.Item(2).Text)
        Exit Sub
    End If
    
    If Len(strFullName) = 0 Then
        MsgBox Label1.Item(3).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(3).SetFocus
        Text1.Item(3).SelStart = 0
        Text1.Item(3).SelLength = Len(Text1.Item(3).Text)
        Exit Sub
    End If
    
    If Len(strDept) = 0 Then
        MsgBox Label1.Item(5).Caption & " ����Ϊ�գ�", vbExclamation
        Combo1.Item(0).SetFocus
        Exit Sub
    End If
    
    If Len(strSex) = 0 Then Option1.Item(0).Value = True
    
    strSQL = "SELECT UserAutoID ,UserLoginName ,UserPassword ," & _
             "UserFullName ,UserSex ,DeptID ,UserMemo " & _
             "From tb_Test_Sys_User " & _
             "WHERE UserAutoID = '" & strUID & "'"
    Set rsUser = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then
        strMsg = "���˺������Ϣ�Ѷ�ʧ������ϵ����Ա��"
        GoTo LineBrk
    ElseIf rsUser.RecordCount > 1 Then
        strMsg = "���˺������Ϣ�쳣������ϵ����Ա��"
        GoTo LineBrk
    Else
        If strLoginName <> rsUser.Fields("UserLoginName").Value Then blnLoginName = True
        If strPWD <> gfDecryptSimple(rsUser.Fields("UserPassword").Value) Then blnPwd = True
        If strFullName <> rsUser.Fields("UserFullName").Value Then blnFullName = True
        If IsNull(rsUser.Fields("UserSex").Value) Or strSex <> rsUser.Fields("UserSex").Value Then blnSex = True
        If IsNull(rsUser.Fields("DeptID").Value) Or strDept <> rsUser.Fields("DeptID").Value Then blnDept = True
        If IsNull(rsUser.Fields("UserMemo").Value) Or strMemo <> rsUser.Fields("UserMemo").Value Then blnMemo = True
        
        If Not (blnLoginName Or blnPwd Or blnFullName Or blnSex Or blnDept Or blnMemo) Then
            strMsg = "û��ʵ���ԵĸĶ����������޸ģ�"
            GoTo LineBrk
        End If
        
        strMsg = "ȷ��Ҫ�޸�" & Label1.Item(0).Caption & "��" & strUID & "�����û���Ϣ��"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then GoTo LineEnd
        
        On Error GoTo LineErr
        
        If blnLoginName Then rsUser.Fields("UserLoginName") = strLoginName
        If blnPwd Then rsUser.Fields("UserPassword") = gfEncryptSimple(strPWD)
        If blnFullName Then rsUser.Fields("UserFullName") = strFullName
        If blnSex Then rsUser.Fields("UserSex") = strSex
        If blnDept Then rsUser.Fields("DeptID") = strDept
        If blnMemo Then rsUser.Fields("UserMemo") = strMemo
        
        rsUser.Update
        rsUser.Close
        
        strMsg = "�޸�ID��" & strUID & "����"
        If blnLoginName Then strMsg = strMsg & "��" & Label1.Item(1).Caption & "��"
        If blnPwd Then strMsg = strMsg & "��" & Label1.Item(2).Caption & "��"
        If blnFullName Then strMsg = strMsg & "[" & Label1.Item(3).Caption & "��"
        If blnSex Then strMsg = strMsg & "��" & Label1.Item(4).Caption & "��"
        If blnDept Then strMsg = strMsg & "��" & Label1.Item(5).Caption & "��"
        If blnMemo Then strMsg = strMsg & "��" & Label1.Item(6).Caption & "��"
        Call gsLogAdd(Me, udUpdate, "tb_Test_Sys_User", strMsg)
        
        MsgBox "�ѳɹ�" & strMsg & "��", vbInformation
        
        If blnFullName Or blnSex Or blnDept Then
            Call msLoadDept(TreeView1)
            Call msLoadUser(TreeView1)
        End If
        
    End If
    
    GoTo LineEnd
    
LineBrk:
    rsUser.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineErr:
    Call gsAlarmAndLog("�û���Ϣ�޸��쳣")
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
    
End Sub

Private Sub Form_Load()

    Set Me.Icon = gMDI.imgListCommandBars.ListImages("SysUser").Picture
    Me.Caption = gMDI.cBS.Actions(gID.SysUser).Caption
    
    For mlngID = Text1.LBound To Text1.UBound
        Text1.Item(mlngID).Text = ""
    Next
    
    TreeView1.Nodes.Clear
    TreeView1.ImageList = gMDI.imgListCommandBars
    
    Call msLoadDept(TreeView1)  '�����ȼ���
    Call msLoadUser(TreeView1)  '��Ա�����
    
End Sub

Private Sub Form_Resize()

    Const conHeight As Long = 9000
    Const conEdge As Long = 120
    
    If Me.WindowState <> vbMinimized Then
        If Me.Height > conHeight Then
            If Me.ScaleHeight > conEdge * 2 Then
                TreeView1.Height = Me.ScaleHeight - conEdge * 2
            End If
        End If
    End If
    
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim lngLen As Long, I As Long
    Dim strKey As String, strUID As String, strSQL As String, strMsg As String
    Dim rsUser As ADODB.Recordset
    
    strKey = Node.Key
    lngLen = Len(strKey)
    If lngLen < Len(mKeyUser) Then Exit Sub
    If Left(strKey, Len(mKeyDept)) = mKeyDept Then
        For mlngID = Text1.LBound To Text1.UBound
            Text1.Item(mlngID).Text = ""
        Next
        Option1.Item(1).Value = True
        Combo1.Item(0).ListIndex = -1
        Exit Sub
    End If
    If Left(strKey, Len(mKeyUser)) <> mKeyUser Then Exit Sub
    
    strUID = Right(Node.Key, lngLen - Len(mKeyUser))
    strSQL = "EXEC sp_Test_Sys_UserInfo '" & strUID & "'"
    Set rsUser = gfBackRecordset(strSQL)
    If rsUser.State = adStateClosed Then GoTo LineEnd
    If rsUser.RecordCount = 0 Then
        strMsg = "�û���Ϣ��ʧ�ˣ�����ϵ����Ա��"
        rsUser.Close
        GoTo LineBreak
    ElseIf rsUser.RecordCount > 1 Then
        strMsg = "�û���Ϣ�쳣������ϵ����Ա��"
        rsUser.Close
        GoTo LineBreak
    Else
        Text1.Item(0).Text = strUID
        Text1.Item(1).Text = rsUser.Fields("UserLoginName").Value & ""
        Text1.Item(2).Text = gfDecryptSimple(rsUser.Fields("UserPassword").Value & "")
        Text1.Item(3).Text = rsUser.Fields("UserFullName").Value & ""
        Text1.Item(4).Text = rsUser.Fields("UserMemo").Value & ""
        
        Option1.Item(0).Value = IIf(rsUser.Fields("UserSex").Value = "Ů", True, False)
        Option1.Item(1).Value = IIf(rsUser.Fields("UserSex").Value = "��", True, False)
        
        If IsNull(rsUser.Fields("DeptID").Value) Then
            Combo1.Item(0).ListIndex = -1
        Else
            For I = 0 To Combo1.Item(1).ListCount - 1
                If rsUser.Fields("DeptID").Value = Combo1.Item(1).List(I) Then
                    Combo1.Item(0).ListIndex = I
                    Exit For
                End If
            Next
            If I = Combo1.Item(1).ListCount Then Combo1.Item(0).ListIndex = -1
        End If

        Node.SelectedImage = "SelectedMen"
    End If

    GoTo LineEnd
    
LineBreak:
    MsgBox strMsg, vbExclamation
LineEnd:
    If rsUser.State = adStateOpen Then rsUser.Close
    Set rsUser = Nothing
End Sub
