VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysRole 
   Caption         =   "��ɫ����"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16725
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   16725
   Begin VB.HScrollBar Hsb 
      Height          =   255
      Left            =   15360
      TabIndex        =   21
      Top             =   6120
      Width           =   1455
   End
   Begin VB.VScrollBar Vsb 
      Height          =   1935
      Left            =   16440
      TabIndex        =   20
      Top             =   4080
      Width           =   255
   End
   Begin VB.Frame ctlMove 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   15855
      Begin VB.Frame Frame1 
         Caption         =   "��ɫ����"
         ForeColor       =   &H00FF0000&
         Height          =   5535
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   7815
         Begin VB.CommandButton Command2 
            Caption         =   "�޸Ľ�ɫ��Ϣ"
            Height          =   495
            Left            =   1560
            TabIndex        =   4
            Top             =   2760
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "��ӽ�ɫ"
            Height          =   495
            Left            =   1560
            TabIndex        =   3
            Top             =   2040
            Width           =   1455
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
            TabIndex        =   1
            Text            =   "Text2"
            Top             =   840
            Width           =   2500
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   1
            Left            =   360
            TabIndex        =   16
            Text            =   "Combo2"
            Top             =   3240
            Visible         =   0   'False
            Width           =   2295
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
            TabIndex        =   0
            Text            =   "Text1"
            Top             =   360
            Width           =   2500
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
            TabIndex        =   2
            Top             =   1320
            Width           =   2500
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4095
            Left            =   3840
            TabIndex        =   5
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7223
            _Version        =   393217
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��������"
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
            Left            =   195
            TabIndex        =   19
            Top             =   1380
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��ɫ����"
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
            Left            =   195
            TabIndex        =   18
            Top             =   900
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��ɫ��ʶ"
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
            Left            =   195
            TabIndex        =   17
            Top             =   420
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "��ɫȨ�޷���"
         ForeColor       =   &H00FF0000&
         Height          =   5535
         Index           =   1
         Left            =   8040
         TabIndex        =   11
         Top             =   0
         Width           =   7695
         Begin VB.CommandButton Command3 
            Caption         =   "�����ɫȨ�޷�����"
            Height          =   495
            Left            =   4920
            TabIndex        =   9
            Top             =   1920
            Width           =   2175
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
            ForeColor       =   &H000000FF&
            Height          =   330
            Index           =   2
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   720
            Width           =   2500
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
            Index           =   2
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1320
            Width           =   2500
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   3
            Left            =   5160
            TabIndex        =   12
            Text            =   "Combo2"
            Top             =   2640
            Visible         =   0   'False
            Width           =   2295
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   4095
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   7223
            _Version        =   393217
            Indentation     =   441
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��ѡ��ɫ"
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
            Height          =   210
            Index           =   4
            Left            =   4080
            TabIndex        =   14
            Top             =   750
            Width           =   900
         End
         Begin VB.Label Label1 
            Caption         =   "����������ɫȨ��"
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
            Height          =   450
            Index           =   3
            Left            =   4080
            TabIndex        =   13
            Top             =   1260
            Width           =   960
         End
      End
   End
End
Attribute VB_Name = "frmSysRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mKeyDept As String = "k"
Private Const mKeyRole As String = "r"
Private Const mOtherKey As String = "kOther"
Private Const mOtherText As String = "������ɫ"
Private Const mKeyFunc As String = "f"
Private Const mHeadKey As String = "kHeadKey"
Private Const mHeadText As String = "���ƹ����б�"
Private Const mTwoBar As String = "--"



Private Function mfCheckRoleFunc() As Boolean
    '����ǰ���
    
    Dim strRID As String, strTemp As String
    
    If TreeView1.Nodes.Count = 0 Then
        MsgBox "������ӽ�ɫ!", vbExclamation
        Exit Function
    End If
    If TreeView2.Nodes.Count = 0 Then
        MsgBox "������ӹ���!", vbExclamation
        Exit Function
    End If
    If TreeView1.SelectedItem Is Nothing Then
        MsgBox "����ѡ��һ����ɫ!", vbExclamation
        Exit Function
    End If
    strTemp = Trim(Text1.Item(2).Text)
    If Len(strTemp) = 0 Then
        MsgBox "����ѡ��һ����ɫ!", vbExclamation
        Exit Function
    End If
    strRID = Left(strTemp, InStr(strTemp, mTwoBar) - 1)
    If strRID <> Trim(Text1.Item(0).Text) Then
        MsgBox "��ⲻһ�£�������ѡ��һ�ν�ɫ!", vbExclamation
        Exit Function
    End If
    
    mfCheckRoleFunc = True
    
End Function

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

Private Sub msLoadFunc(ByRef tvwLoad As MSComctlLib.TreeView)
    '���ع����б�
    
    Dim rsFunc As ADODB.Recordset
    Dim strSQL As String
    Dim arrFunc() As String
    Dim I As Long, lngCount As Long
    Dim blnLoop As Boolean
    
    tvwLoad.Nodes.Clear
    tvwLoad.Nodes.Add , , mHeadKey, mHeadText, "FuncHead"   '����׽��
    tvwLoad.Nodes(mHeadKey).Expanded = True     'չ�����
        
    strSQL = "SELECT FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID " & _
             "FROM tb_Test_Sys_Func ORDER BY FuncType ,FuncName "
    Set rsFunc = gfBackRecordset(strSQL)
    If rsFunc.State = adStateClosed Then Exit Sub
    
    If rsFunc.RecordCount > 0 Then
        While Not rsFunc.EOF
            If rsFunc.Fields("FuncType") = gID.FuncMainMenu Then
                tvwLoad.Nodes.Add mHeadKey, tvwChild, mKeyFunc & rsFunc.Fields("FuncAutoID"), rsFunc.Fields("FuncCaption"), "FuncMainMenu"
                tvwLoad.Nodes.Item(mKeyFunc & rsFunc.Fields("FuncAutoID")).Expanded = True
            Else
                ReDim Preserve arrFunc(4, lngCount)
                For I = 0 To 4
                    arrFunc(I, lngCount) = rsFunc.Fields(I).Value
                Next
                lngCount = lngCount + 1
                blnLoop = True
            End If
            rsFunc.MoveNext
        Wend
    End If
    
    If blnLoop Then Call msLoadFuncTree(tvwLoad, arrFunc)
    
    If rsFunc.State = adStateOpen Then rsFunc.Close
    Set rsFunc = Nothing
    
End Sub

Private Sub msLoadFuncTree(ByRef tvwTree As MSComctlLib.TreeView, ByRef arrLoad() As String)
    '������msLoadFunc�������ʹ���������б�
    
    Dim arrOther() As String    '����ʣ���
    Dim blnOther As Boolean     'ʣ���ʶ
    Dim I As Long, J As Long, K As Long, lngCount As Long
    Dim strImage As String
    
    With tvwTree
        For J = LBound(arrLoad, 2) To UBound(arrLoad, 2)
            For I = 1 To .Nodes.Count   'ע��˴��±��1��ʼ
                If .Nodes.Item(I).Key = mKeyFunc & arrLoad(4, J) Then   ' FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID
                    If arrLoad(3, J) = gID.FuncButton Then
                        strImage = "FuncButton"
                    ElseIf arrLoad(3, J) = gID.FuncForm Then
                        strImage = "FuncForm"
                    Else
                        strImage = "FuncControl"
                    End If
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, mKeyFunc & arrLoad(0, J), arrLoad(2, J), strImage
                    If arrLoad(3, J) = gID.FuncForm Then .Nodes(mKeyFunc & arrLoad(0, J)).Expanded = True
                    Exit For
                End If
            Next

            If I = .Nodes.Count + 1 Then
                If arrLoad(3, J) = gID.FuncForm Then
                    .Nodes.Add mHeadKey, tvwChild, mKeyFunc & arrLoad(0, J), arrLoad(2, J), "FuncMainMenu"
                    .Nodes(mKeyFunc & arrLoad(0, J)).Expanded = True
                Else
                    blnOther = True
                    ReDim Preserve arrOther(4, lngCount)
                    For K = 0 To 4
                        arrOther(K, lngCount) = arrLoad(K, J)
                    Next
                    lngCount = lngCount + 1
                End If
            End If
            
        Next
    End With
    
    If blnOther Then
        Call msLoadFuncTree(tvwTree, arrOther)
'        MsgBox mHeadText & "���ز���ȫ����֪ͨ����Ա��", vbCritical
    End If

End Sub


Private Sub msLoadRole(ByRef tvwUser As MSComctlLib.TreeView)
    '���ؽ�ɫ��ǰ�����Ѽ��غò���
    
    Dim strSQL As String
    Dim rsRole As ADODB.Recordset
    Dim arrOther() As String    '����ʣ���
    Dim blnOther As Boolean     'ʣ���ʶ
    Dim I As Long, J As Long, K As Long, lngCount As Long
    
    If tvwUser.Nodes.Count = 0 Then Exit Sub
    
    strSQL = "SELECT RoleAutoID ,RoleName ,DeptID FROM tb_Test_Sys_Role "
    Set rsRole = gfBackRecordset(strSQL)
    If rsRole.State = adStateClosed Then GoTo LineEnd
    If rsRole.RecordCount = 0 Then GoTo LineEnd
    
    Combo1.Item(2).Clear
    Combo1.Item(3).Clear
    
    With tvwUser
        While Not rsRole.EOF
            For I = 1 To .Nodes.Count
                If .Nodes(I).Key = mKeyDept & rsRole.Fields("DeptID").Value Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, _
                        mKeyRole & rsRole.Fields("RoleAutoID").Value, rsRole.Fields("RoleName").Value, "SysRole"
                    Exit For
                End If
            Next
            
            Combo1.Item(2).AddItem rsRole.Fields("RoleName")
            Combo1.Item(3).AddItem rsRole.Fields("RoleAutoID")
            
            If I = .Nodes.Count + 1 Then
                blnOther = True
                ReDim Preserve arrOther(2, lngCount)
                For K = 0 To 2
                    arrOther(K, lngCount) = rsRole.Fields(K).Value & ""
                Next
                lngCount = lngCount + 1
            End If
            
            rsRole.MoveNext
        Wend
        
        If blnOther Then
            .Nodes.Add 1, tvwChild, mOtherKey, mOtherText, "unknown"
            .Nodes(mOtherKey).Expanded = True
            For I = LBound(arrOther, 2) To UBound(arrOther, 2)
                .Nodes.Add mOtherKey, tvwChild, mKeyRole & arrOther(0, I), _
                    arrOther(1, I), "SysRole"
            Next
        End If

    End With
    
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    
End Sub

Private Sub msLoadRoleFunc(ByVal strRID As String)
    '���ض�ӦȨ��
    
    Dim I As Long
    Dim strSQL As String
    Dim rsRole As ADODB.Recordset
    
    strSQL = "SELECT RoleAutoID ,FuncAutoID " & _
             "From tb_Test_Sys_RoleFunc " & _
             "WHERE RoleAutoID = " & strRID & "  ORDER BY FuncAutoID"
    Set rsRole = gfBackRecordset(strSQL)
    If rsRole.State = adStateOpen Then
        With TreeView2.Nodes
            For I = 2 To .Count 'ע��˫��ѭ��Ұ��˳���ܵߵ��������������
                If rsRole.RecordCount > 0 Then
                    rsRole.MoveFirst
                    Do While (Not rsRole.EOF)
                        If mKeyFunc & rsRole.Fields("FuncAutoID") = .Item(I).Key Then
                            .Item(I).Checked = True
                            Exit Do
                        End If
                        rsRole.MoveNext
                    Loop
                End If
                If rsRole.EOF Then
                    .Item(I).Checked = False
                End If
            Next
        End With
    End If
    
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    
End Sub


Private Sub Combo1_Click(Index As Integer)
    '����
    
    Dim strRID As String, strTemp As String, strSQL As String, strMsg As String
    Dim strLeadText As String, strLeadID As String
    Dim cnRole As ADODB.Connection
    Dim rsRole As ADODB.Recordset, rsLead As ADODB.Recordset, rsHave As ADODB.Recordset
    Dim blnTran As Boolean
    Dim I As Long
    
    If Index <> 2 Then Exit Sub
    If Not mfCheckRoleFunc Then Exit Sub
    
    If Combo1.Item(2).ListCount = 0 Then Exit Sub
    strLeadText = Trim(Combo1.Item(2).Text)
    strLeadID = Trim(Combo1.Item(3).List(Combo1.Item(2).ListIndex))
    If Len(strLeadText) = 0 Then Exit Sub
    
    strTemp = Trim(Text1.Item(2).Text)
    strRID = Left(strTemp, InStr(strTemp, mTwoBar) - 1)
    
    strSQL = "SELECT RoleAutoID ,FuncAutoID FROM tb_Test_Sys_RoleFunc " & _
             "WHERE RoleAutoID =" & strLeadID
    Set rsLead = gfBackRecordset(strSQL)
    If rsLead.State = adStateClosed Then
        Set rsLead = Nothing
        Exit Sub
    End If
    If rsLead.RecordCount = 0 Then
        rsLead.Close
        Set rsLead = Nothing
        MsgBox "��ɫ��" & strLeadID & mTwoBar & strLeadText & "������û�з����κ�Ȩ��Ӵ��", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("ȷ��Ҫ����ɫ��" & strLeadID & mTwoBar & strLeadText & _
        "����Ȩ�޵������ɫ��" & strTemp & "����", vbQuestion + vbOKCancel, "����ѯ��") = vbCancel Then Exit Sub
    
    Set cnRole = New ADODB.Connection
    Set rsRole = New ADODB.Recordset
    Set rsHave = New ADODB.Recordset
    cnRole.CursorLocation = adUseClient
    
    On Error GoTo LineErr
    
    cnRole.Open gID.CnString
    cnRole.BeginTrans
    blnTran = True
    
    strSQL = "SELECT RoleAutoID ,FuncAutoID FROM tb_Test_Sys_RoleFunc " & _
             "WHERE RoleAutoID =" & strRID
    rsRole.Open strSQL, cnRole, adOpenStatic, adLockBatchOptimistic
    rsHave.Open strSQL, cnRole, adOpenStatic, adLockReadOnly
    
    If rsRole.RecordCount = 0 Then
        While Not rsLead.EOF
            rsRole.AddNew
            rsRole.Fields("RoleAutoID") = strRID
            rsRole.Fields("FuncAutoID") = rsLead.Fields("FuncAutoID")
            rsLead.MoveNext
        Wend
    Else
        If rsHave.RecordCount = 0 Then
            strMsg = "��ɫ��" & strTemp & "��Ȩ����Ϣ�쳣�������Ի���ϵ����Ա��"
            GoTo LineBreak
        End If
        
        While Not rsLead.EOF
            rsHave.MoveFirst
            Do While Not rsHave.EOF
                If rsHave.Fields("FuncAutoID") = rsLead.Fields("FuncAutoID") Then Exit Do
                rsHave.MoveNext
            Loop
            If rsHave.EOF Then
                rsRole.AddNew
                rsRole.Fields("RoleAutoID") = strRID
                rsRole.Fields("FuncAutoID") = rsLead.Fields("FuncAutoID")
            End If
            rsLead.MoveNext
        Wend
        
    End If
    
    rsRole.UpdateBatch
    cnRole.CommitTrans
    rsLead.Close
    rsRole.Close
    rsHave.Close
    cnRole.Close
    
    Call gsLogAdd(Me, udInsertBatch, "tb_Test_Sys_RoleFunc", "����ɫ��" & strLeadID & mTwoBar & strLeadText & "����Ȩ�޵������ɫ��" & strTemp & "��")
    Call msLoadRoleFunc(strRID)
    MsgBox "�ѳɹ�����ɫ��" & strLeadID & mTwoBar & strLeadText & "����Ȩ�޵������ɫ��" & strTemp & "��", vbInformation
    
    GoTo LineEnd
    
LineErr:
    If blnTran Then cnRole.RollbackTrans
    If rsLead.State = adStateOpen Then rsLead.Close
    If rsRole.State = adStateOpen Then rsRole.Close
    If rsHave.State = adStateOpen Then rsHave.Close
    If cnRole.State = adStateOpen Then cnRole.Close
    Call gsAlarmAndLog("Ϊ��" & strTemp & "�������ɫȨ���쳣")
    GoTo LineEnd
    
LineBreak:
    If blnTran Then cnRole.RollbackTrans
    If rsLead.State = adStateOpen Then rsLead.Close
    If rsRole.State = adStateOpen Then rsRole.Close
    If rsHave.State = adStateOpen Then rsHave.Close
    If cnRole.State = adStateOpen Then cnRole.Close
    MsgBox strMsg, vbExclamation
    
LineEnd:
    If rsLead.State = adStateOpen Then rsLead.Close
    Set rsLead = Nothing
    If rsRole.State = adStateOpen Then rsRole.Close
    If rsHave.State = adStateOpen Then rsHave.Close
    If cnRole.State = adStateOpen Then cnRole.Close
    Set rsRole = Nothing
    Set rsHave = Nothing
    Set cnRole = Nothing
    
End Sub

Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            Combo1.Item(Index).ListIndex = -1
        End If
    End If
End Sub

Private Sub Command1_Click()
    '��ӽ�ɫ

    Dim rsRole As ADODB.Recordset
    Dim strRoleName As String, strSQL As String, strCheck As String, strMsg As String
    Dim strDeptID As Variant
    Dim lngRID As Long

    
    strRoleName = Trim(Text1.Item(1).Text)
    strRoleName = Left(strRoleName, 50)
    Text1.Item(1).Text = strRoleName
    
    If Len(strRoleName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ���ַ���", vbExclamation
        Text1.Item(1).SetFocus
        Exit Sub
    End If
    
    strCheck = gfStringCheck(strRoleName)
    If Len(strCheck) > 0 Then
        MsgBox Label1.Item(1).Caption & "�в��ܰ��������ַ���" & strCheck & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(0).SelLength = Len(strRoleName)
        Exit Sub
    End If
    
    If Combo1.Item(0).ListIndex = -1 Then
        strDeptID = Null
    Else
        strDeptID = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    End If
    
    strMsg = "ȷ�����" & Label1.Item(1).Caption & "��" & strRoleName & "����"
    If MsgBox(strMsg, vbQuestion + vbOKCancel, "ȷ��ѯ��") = vbCancel Then Exit Sub
    
    strSQL = "SELECT RoleAutoID ,RoleName ,DeptID FROM tb_Test_Sys_Role WHERE 1<>1 "
    Set rsRole = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    
    On Error GoTo LineErr
    
    If rsRole.State = adStateOpen Then
        rsRole.AddNew
        rsRole.Fields("RoleName") = strRoleName
        rsRole.Fields("DeptID") = strDeptID
        rsRole.Update
        lngRID = rsRole.Fields("RoleAutoID")
        rsRole.Close
        Call msLoadDept(TreeView1)
        Call msLoadRole(TreeView1)
        Call gsLogAdd(Me, udInsert, "tb_Test_Sys_Role", "����½�ɫ��" & lngRID & "����" & strRoleName & "��")
        MsgBox "��ɫ��" & strRoleName & "����ӳɹ���", vbInformation
    End If
    
    GoTo LineEnd
    
LineErr:
    Call gsAlarmAndLog("��ɫ����쳣")
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    Text1.Item(2).Text = ""
    
End Sub

Private Sub Command2_Click()
    '�޸Ľ�ɫ
    
    Dim rsRole As ADODB.Recordset
    Dim strSQL As String, strMsg As String, strCheck As String
    Dim strRID As String, strRoleName As String, strLastDept As String
    Dim strDeptID As Variant
    Dim blnName As Boolean, blnDeptID As Boolean
    
    
    strRID = Trim(Text1.Item(0).Text)
    
    strRoleName = Trim(Text1.Item(1).Text)
    Text1.Item(1).Text = strRoleName
    
    If Combo1.Item(0).ListIndex = -1 Then
        strDeptID = Null
    Else
        strDeptID = Trim(Combo1.Item(1).List(Combo1.Item(0).ListIndex))
    End If
    
    If Len(strRID) = 0 Then
        MsgBox "����ѡ��һ����ɫ��", vbExclamation
        Exit Sub
    End If
    
    If Len(strRoleName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ���ַ���", vbExclamation
        Text1.Item(1).SetFocus
        Exit Sub
    End If
    strCheck = gfStringCheck(strRoleName)
    If Len(strCheck) > 0 Then
        MsgBox Label1.Item(1).Caption & " ���ܺ��������ַ���" & strCheck & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strRoleName)
        Exit Sub
    End If

    With TreeView1
        If .SelectedItem Is Nothing Then
            MsgBox "�ڲ�����쳣��������ѡ���ɫ��", vbExclamation
            Exit Sub
        End If
        
        If .SelectedItem.Key <> mKeyRole & strRID Then
            MsgBox "�ڲ�����쳣��������ѡ��һ����ɫ��", vbExclamation
            Exit Sub
        End If
        
        If .SelectedItem.Text <> strRoleName Then blnName = True
        
        If .SelectedItem.Parent.Key = mOtherKey Then
            If Not IsNull(strDeptID) Then blnDeptID = True
            strLastDept = ""
        Else
            If IsNull(strDeptID) Then
                blnDeptID = True
            Else
               If (mKeyDept & strDeptID) <> .SelectedItem.Parent.Key Then blnDeptID = True
            End If
            strLastDept = .SelectedItem.Parent.Text
        End If
                
        If Not (blnName Or blnDeptID) Then
            MsgBox "û��ʵ���ԵĸĶ��������޸ġ�", vbExclamation
            Exit Sub
        End If
        
        strMsg = "���ζ�IDΪ��" & strRID & "���Ľ�ɫ��Ϣ�޸�������£�" & vbCrLf & vbCrLf & _
                 Space(6) & "�޸�λ��" & vbTab & "�޸�ǰ" & vbTab & vbTab & "�޸ĺ�" & vbCrLf & vbCrLf & _
                 Space(6) & Label1(1).Caption & vbTab & .SelectedItem.Text & vbTab & vbTab & strRoleName & vbCrLf & _
                 Space(6) & Label1(2).Caption & vbTab & strLastDept & vbTab & vbTab & Combo1.Item(0).Text & vbCrLf & vbCrLf & _
                 "�Ƿ��������޸ģ�"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
    End With
    
    strSQL = "SELECT RoleAutoID ,RoleName ,DeptID " & _
             "FROM tb_Test_Sys_Role " & _
             "WHERE RoleAutoID =" & strRID & ""
    Set rsRole = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    
    On Error GoTo LineErr
    
    If rsRole.State = adStateOpen Then
        If rsRole.RecordCount = 1 Then
            strMsg = "��" & strRID & "����ɫ��Ϣ�޸ģ�"
            If blnName Then
                strMsg = strMsg & Label1.Item(1).Caption & "[" & rsRole.Fields("RoleName") & "]-->[" & strRoleName & "];"
                rsRole.Fields("RoleName") = strRoleName
            End If
            If blnDeptID Then
                rsRole.Fields("DeptID") = strDeptID
                strMsg = strMsg & Label1.Item(2).Caption & "[" & strLastDept & "]-->[" & Combo1.Item(0).Text & "];"
            End If
            rsRole.Update
            rsRole.Close
            Call msLoadDept(TreeView1)
            Call msLoadRole(TreeView1)
            Call gsLogAdd(Me, udUpdate, "tb_Test_Sys_Role", strMsg)
            MsgBox "��ɫ��Ϣ�޸���ɣ�", vbInformation
        Else
            rsRole.Close
            MsgBox "��̨�����쳣������ϵ����Ա��", vbCritical
        End If

    End If
    
    GoTo LineEnd
    
LineErr:
    Call gsAlarmAndLog("��ɫ�޸��쳣")
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    Text1.Item(2).Text = ""
    
End Sub

Private Sub Command3_Click()
    '����
    
    Dim strRID As String, strTemp As String, strSQL As String, strMsg As String
    Dim cnRole As ADODB.Connection
    Dim rsRole As ADODB.Recordset
    Dim blnTran As Boolean
    Dim I As Long
    
    If Not mfCheckRoleFunc Then Exit Sub
    
    strTemp = Trim(Text1.Item(2).Text)
    strRID = Left(strTemp, InStr(strTemp, mTwoBar) - 1)
    
    If MsgBox("ȷ�����桾" & strTemp & "����Ȩ�޷�����Ϣ��", vbQuestion + vbOKCancel, "����ѯ��") = vbCancel Then Exit Sub
    
    Set cnRole = New ADODB.Connection
    Set rsRole = New ADODB.Recordset
    cnRole.CursorLocation = adUseClient
    
    On Error GoTo LineErr
    
    cnRole.Open gID.CnString
    cnRole.BeginTrans
    blnTran = True
    
    'ɾ������Ȩ��
    strSQL = "DELETE FROM tb_Test_Sys_RoleFunc WHERE RoleAutoID =" & strRID
    cnRole.Execute strSQL
    
    '�����·���Ȩ��
    strSQL = "SELECT RoleAutoID ,FuncAutoID FROM tb_Test_Sys_RoleFunc WHERE RoleAutoID =" & strRID
    rsRole.Open strSQL, cnRole, adOpenStatic, adLockBatchOptimistic
    If rsRole.RecordCount > 0 Then
        strMsg = "��" & Label1.Item(0).Caption & strRID & "�� �ĺ�̨Ȩ����Ϣ�쳣�������Ի���ϵ����Ա��"
        GoTo LineBreak
    End If
    With TreeView2.Nodes
        For I = 2 To .Count
            If .Item(I).Checked Then
                rsRole.AddNew
                rsRole.Fields("RoleAutoID") = strRID
                rsRole.Fields("FuncAutoID") = Right(.Item(I).Key, Len(.Item(I).Key) - Len(mKeyFunc))
            End If
        Next
    End With
    rsRole.UpdateBatch
    cnRole.CommitTrans
    rsRole.Close
    cnRole.Close
    Call gsLogAdd(Me, udInsertBatch, "tb_Test_Sys_RoleFunc", "���桾" & strTemp & "����Ȩ�޷�����Ϣ")
    MsgBox strTemp & " ��Ȩ�޷�����Ϣ����ɹ���", vbInformation
    
    GoTo LineEnd
    
LineErr:
    If blnTran Then cnRole.RollbackTrans
    If rsRole.State = adStateOpen Then rsRole.Close
    If cnRole.State = adStateOpen Then cnRole.Close
    Call gsAlarmAndLog(Command3.Caption & " �쳣")
    GoTo LineEnd
    
LineBreak:
    If blnTran Then cnRole.RollbackTrans
    If rsRole.State = adStateOpen Then rsRole.Close
    If cnRole.State = adStateOpen Then cnRole.Close
    MsgBox strMsg, vbExclamation
    
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    If cnRole.State = adStateOpen Then cnRole.Close
    Set rsRole = Nothing
    Set cnRole = Nothing
    
End Sub

Private Sub Form_Load()
    
    Me.Icon = gMDI.imgListCommandBars.ListImages("SysRole").Picture
    Me.Caption = gMDI.cBS.Actions(gID.SysRole).Caption
    Frame1.Item(0).Caption = Me.Caption
    
    Text1.Item(0).Text = ""
    Text1.Item(1).Text = ""
    Text1.Item(2).Text = ""
    TreeView1.Nodes.Clear
    TreeView2.Nodes.Clear
    TreeView1.ImageList = gMDI.imgListCommandBars
    TreeView2.ImageList = gMDI.imgListCommandBars
    
    Call msLoadDept(TreeView1)
    Call msLoadRole(TreeView1)  '��ɫ�����
    Call msLoadFunc(TreeView2)
    
    Call gsLoadAuthority(Me, TreeView1)
    Call gsLoadAuthority(Me, Command1)
    Call gsLoadAuthority(Me, Command2)
    Call gsLoadAuthority(Me, Command3)
    Call gsLoadAuthority(Me, Combo1.Item(2))
    
End Sub

Private Sub Form_Resize()

    Const conHeight As Long = 6000
    Const conEdge As Long = 120
    Const conTB As Long = 400
    
    If Me.WindowState <> vbMinimized Then
        If Me.Height > conHeight Then
            If Me.ScaleHeight > conEdge * 2 Then
                Frame1.Item(0).Height = Me.ScaleHeight - conEdge * 2
                Frame1.Item(1).Height = Frame1.Item(0).Height
                TreeView1.Height = Frame1.Item(0).Height - conTB
                TreeView2.Height = TreeView1.Height
                ctlMove.Height = Frame1.Item(0).Height
            End If
        End If
    End If
    
    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 16300, 9000)  'ע�ⳤ������޸�
    
End Sub

Private Sub Hsb_Change()
    ctlMove.Left = -Hsb.Value
End Sub

Private Sub Hsb_Scroll()
    Call Hsb_Change    '�������������еĻ���ʱ��ͬʱ���¶�Ӧ���ݣ�����ͬ��
End Sub

Private Sub Vsb_Change()
    ctlMove.Top = -Vsb.Value
End Sub

Private Sub Vsb_Scroll()
    Call Vsb_Change
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim lngLen As Long, I As Long
    Dim strKey As String, strRID As String, strSQL As String, strMsg As String
    Dim rsRole As ADODB.Recordset
    
    Text1.Item(2).Text = ""
    strKey = Node.Key
    lngLen = Len(strKey)
    If lngLen < Len(mKeyRole) Then Exit Sub
    If Left(strKey, Len(mKeyDept)) = mKeyDept Then
        Text1.Item(0).Text = ""
        Text1.Item(1).Text = ""
        Text1.Item(2).Text = ""
        Combo1.Item(0).ListIndex = -1
        For I = 1 To TreeView2.Nodes.Count
            TreeView2.Nodes.Item(I).Checked = False
        Next
        Exit Sub
    End If
    If Left(strKey, Len(mKeyRole)) <> mKeyRole Then Exit Sub
    
    strRID = Right(Node.Key, lngLen - Len(mKeyRole))
    strSQL = "SELECT RoleAutoID ,RoleName ,DeptID FROM tb_Test_Sys_Role " & _
             "WHERE RoleAutoID =" & strRID
    Set rsRole = gfBackRecordset(strSQL)
    If rsRole.State = adStateClosed Then GoTo LineEnd
    If rsRole.RecordCount = 0 Then
        strMsg = "��ɫ��Ϣ��ʧ�ˣ�����ϵ����Ա��"
        GoTo LineBreak
    ElseIf rsRole.RecordCount > 1 Then
        strMsg = "��ɫ��Ϣ�쳣������ϵ����Ա��"
        GoTo LineBreak
    Else
        Text1.Item(0).Text = strRID
        Text1.Item(2).Text = strRID & mTwoBar & rsRole.Fields("RoleName")
        Text1.Item(1).Text = rsRole.Fields("RoleName").Value
        If IsNull(rsRole.Fields("DeptID").Value) Then
            Combo1.Item(0).ListIndex = -1
        Else
            For I = 0 To Combo1.Item(1).ListCount - 1
                If rsRole.Fields("DeptID").Value = Combo1.Item(1).List(I) Then
                    Combo1.Item(0).ListIndex = I
                    Exit For
                End If
            Next
            If I = Combo1.Item(1).ListCount Then Combo1.Item(0).ListIndex = -1
        End If
        Node.SelectedImage = "RoleSelect"
    End If
    
    rsRole.Close
    Call msLoadRoleFunc(strRID)
    
    GoTo LineEnd
    
LineBreak:
    rsRole.Close
    MsgBox strMsg, vbExclamation
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    
End Sub

Private Sub TreeView2_NodeCheck(ByVal Node As MSComctlLib.Node)
    Call gsNodeCheckCascade(Node, Node.Checked)
End Sub
