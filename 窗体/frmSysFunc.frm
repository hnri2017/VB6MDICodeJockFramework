VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysFunc 
   Caption         =   "��������"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   17115
   Begin VB.VScrollBar Vsb 
      Height          =   1935
      Left            =   16200
      TabIndex        =   22
      Top             =   4080
      Width           =   255
   End
   Begin VB.HScrollBar Hsb 
      Height          =   255
      Left            =   15120
      TabIndex        =   21
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Frame ctlMove 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5895
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   15975
      Begin VB.Frame Frame1 
         Caption         =   "�����ָ����ɫ"
         ForeColor       =   &H00FF0000&
         Height          =   5535
         Index           =   1
         Left            =   8040
         TabIndex        =   19
         Top             =   0
         Width           =   7695
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
            Index           =   3
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   720
            Width           =   2500
         End
         Begin VB.CommandButton Command3 
            Caption         =   "�����ָ����ɫ�������"
            Height          =   495
            Left            =   4680
            TabIndex        =   10
            Top             =   1920
            Width           =   2415
         End
         Begin MSComctlLib.TreeView TreeView2 
            Height          =   4095
            Left            =   120
            TabIndex        =   8
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
            Caption         =   "��ѡ����"
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
            Index           =   5
            Left            =   4080
            TabIndex        =   20
            Top             =   750
            Width           =   900
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "����"
         ForeColor       =   &H00FF0000&
         Height          =   5535
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   7815
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
            TabIndex        =   4
            Top             =   2160
            Width           =   2500
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
            Top             =   240
            Width           =   2500
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Index           =   1
            Left            =   1920
            TabIndex        =   13
            Text            =   "Combo2"
            Top             =   2760
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
            TabIndex        =   1
            Text            =   "Text2"
            Top             =   720
            Width           =   2500
         End
         Begin VB.CommandButton Command1 
            Caption         =   "��ӹ���"
            Height          =   495
            Left            =   600
            TabIndex        =   5
            Top             =   3360
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
            Caption         =   "�޸Ĺ�����Ϣ"
            Height          =   495
            Left            =   2280
            TabIndex        =   6
            Top             =   3360
            Width           =   1335
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
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1680
            Width           =   2500
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
            Index           =   2
            Left            =   1200
            TabIndex        =   2
            Text            =   "Text2"
            Top             =   1200
            Width           =   2500
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4095
            Left            =   3840
            TabIndex        =   7
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
            Caption         =   "�ϼ�����"
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
            Left            =   255
            TabIndex        =   18
            Top             =   2220
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�������"
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
            Left            =   255
            TabIndex        =   17
            Top             =   1740
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�Զ����"
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
            Left            =   255
            TabIndex        =   16
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "���ܱ�ʶ"
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
            Left            =   255
            TabIndex        =   15
            Top             =   780
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "���ܱ���"
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
            Left            =   255
            TabIndex        =   14
            Top             =   1260
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmSysFunc"
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
    
    Dim strFID As String, strTemp As String
    
    If TreeView1.Nodes.Count = 0 Then
        MsgBox "������ӹ���!", vbExclamation
        Exit Function
    End If
    If TreeView2.Nodes.Count = 0 Then
        MsgBox "������ӽ�ɫ!", vbExclamation
        Exit Function
    End If
    If TreeView1.SelectedItem Is Nothing Then
        MsgBox "����ѡ��һ������!", vbExclamation
        Exit Function
    End If
    strTemp = Trim(Text1.Item(3).Text)
    If Len(strTemp) = 0 Then
        MsgBox "����ѡ��һ������!", vbExclamation
        Exit Function
    End If
    strFID = Left(strTemp, InStr(strTemp, mTwoBar) - 1)
    If strFID <> Trim(Text1.Item(0).Text) Then
        MsgBox "��ⲻһ�£�������ѡ��һ������!", vbExclamation
        Exit Function
    End If
    
    mfCheckRoleFunc = True
    
End Function

Private Function mfFuncTypeCheck(ByVal strType As String) As Boolean
    '��鹦������Ƿ���ȷ
    
    Select Case strType
        Case gID.FuncButton, gID.FuncControl, gID.FuncForm
            mfFuncTypeCheck = True
        Case Else
    End Select
    
End Function

Private Sub msFuncTypeCheck()
    '����������ϼ�����֮����໥Լ��
    
    If Combo1.Item(2).Text = gID.FuncMainMenu Then
        Combo1.Item(0).Text = mHeadText
    End If

    If Combo1.Item(0).Text = mHeadText Then
        Combo1.Item(2).Text = gID.FuncMainMenu
    End If

End Sub

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
             "FROM tb_Test_Sys_Func ORDER BY FuncType ,FuncName " & _
             "SELECT FuncAutoID ,FuncCaption FROM tb_Test_Sys_Func " & _
             "WHERE FuncType ='" & gID.FuncMainMenu & "' OR FuncType ='" & gID.FuncForm & "' ORDER BY FuncCaption "
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
    
    Set rsFunc = rsFunc.NextRecordset
    If rsFunc.State = adStateOpen Then
    
        Combo1.Item(0).Clear
        Combo1.Item(1).Clear
        Combo1.Item(0).AddItem mHeadText
        Combo1.Item(1).AddItem mHeadKey
        If rsFunc.RecordCount > 0 Then
            While Not rsFunc.EOF
                Combo1.Item(0).AddItem rsFunc.Fields("FuncCaption")
                Combo1.Item(1).AddItem rsFunc.Fields("FuncAutoID")
                rsFunc.MoveNext
            Wend
        End If
    End If
    
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
    
    With tvwUser
        While Not rsRole.EOF
            For I = 1 To .Nodes.Count
                If .Nodes(I).Key = mKeyDept & rsRole.Fields("DeptID").Value Then
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, _
                        mKeyRole & rsRole.Fields("RoleAutoID").Value, rsRole.Fields("RoleName").Value, "SysRole"
                    Exit For
                End If
            Next
            
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

Private Sub msLoadRoleFunc(ByVal strFID As String)
    '���ض�ӦȨ��
    
    Dim I As Long
    Dim strSQL As String
    Dim rsRole As ADODB.Recordset
    
    strSQL = "SELECT RoleAutoID ,FuncAutoID " & _
             "From tb_Test_Sys_RoleFunc " & _
             "WHERE FuncAutoID = " & strFID & "  ORDER BY RoleAutoID"
    Set rsRole = gfBackRecordset(strSQL)
    If rsRole.State = adStateOpen Then
        With TreeView2.Nodes
            For I = 2 To .Count 'ע��˫��ѭ��Ұ��˳���ܵߵ��������������
                If rsRole.RecordCount > 0 Then
                    rsRole.MoveFirst
                    Do While (Not rsRole.EOF)
                        If mKeyRole & rsRole.Fields("RoleAutoID") = .Item(I).Key Then
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


Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        Combo1.Item(Index).ListIndex = -1
    End If
End Sub

Private Sub Command1_Click()
    '���
    
    Dim strName As String, strCaption As String
    Dim strType As String, strParent As Variant
    Dim strSQL As String, strMsg As String
    Dim rsFunc As ADODB.Recordset
    
    Call msFuncTypeCheck
    
    strName = Trim(Text1.Item(1).Text)
    strCaption = Trim(Text1.Item(2).Text)
    strType = Combo1.Item(2).Text
    strParent = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    strName = Left(strName, 50)
    strCaption = Left(strCaption, 50)
    
    Text1.Item(1).Text = strName
    Text1.Item(2).Text = strCaption

    If Len(strName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strName)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " ���ܺ��������ַ���" & strMsg & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strName)
        Exit Sub
    End If
    
    If Len(strCaption) = 0 Then
        MsgBox Label1.Item(2).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strCaption)
        Exit Sub
    End If
    
    If Len(strType) = 0 Then
        MsgBox Label1.Item(3).Caption & " ����Ϊ�գ�", vbExclamation
        Combo1.Item(2).SetFocus
        Exit Sub
    End If
    
    If Len(strParent) = 0 Then
        MsgBox Label1.Item(4).Caption & " ����Ϊ�գ�", vbExclamation
        Combo1.Item(0).SetFocus
        Exit Sub
    End If
        
    If MsgBox("�Ƿ���ӹ��ܡ�" & strName & "����" & strCaption & "����" & strType & "����", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    If strType = gID.FuncForm Then
        strSQL = "SELECT FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID FROM tb_Test_Sys_Func " & _
                 "WHERE FuncName = '" & strName & "'"
    Else
        strSQL = "SELECT FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID FROM tb_Test_Sys_Func " & _
                 "WHERE  FuncParentID =" & Val(strParent) & " AND FuncName ='" & strName & "'"
    End If
    
    Set rsFunc = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsFunc.State = adStateClosed Then GoTo LineEnd
    If rsFunc.RecordCount > 0 Then
        If strType = gID.FuncForm Then
            strMsg = Label1.Item(1).Caption & "�Ѵ��ڣ��������"
        Else
            strMsg = Label1.Item(1).Caption & " �� " & Label1.Item(4).Caption & " ���Ѵ��ڣ��������"
        End If
        GoTo LineBrk
    Else
        On Error GoTo LineErr
        
        rsFunc.AddNew
        rsFunc.Fields("FuncName") = strName
        rsFunc.Fields("FuncCaption") = strCaption
        rsFunc.Fields("FuncType") = strType
        rsFunc.Fields("FuncParentID") = Val(strParent)
        rsFunc.Update
        strMsg = rsFunc.Fields("FuncAutoID").Value
        Text1.Item(0).Text = strMsg
        rsFunc.Close
        strMsg = "��ӹ��ܡ�" & strMsg & "����" & strName & "����" & strCaption & "����" & strType & "��"
        Call gsLogAdd(Me, udInsert, "tb_Test_Sys_Func", strMsg)
        MsgBox "��" & strName & "����" & strCaption & "����" & strType & "����ӳɹ���", vbInformation
        Call msLoadFunc(TreeView1)
    End If
    
    GoTo LineEnd
    
LineBrk:
    rsFunc.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineErr:
    Call gsAlarmAndLog("��ӹ����쳣")
LineEnd:
    If rsFunc.State = adStateOpen Then rsFunc.Close
    Set rsFunc = Nothing
    
End Sub

Private Sub Command2_Click()
    '�޸�
    
    Dim strFID As String, strName As String, strCaption As String
    Dim strType As String, strParent As String, strSQL As String, strMsg As String
    Dim blnName As Boolean, blnCaption As Boolean
    Dim blnType As Boolean, blnParent As Boolean
    Dim rsFunc As ADODB.Recordset
    
    Call msFuncTypeCheck
    
    strFID = Trim(Text1.Item(0).Text)
    strName = Trim(Text1.Item(1).Text)
    strCaption = Trim(Text1.Item(2).Text)
    strType = Combo1.Item(2).Text
    strParent = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    strName = Left(strName, 50)
    strCaption = Left(strCaption, 50)
    
    Text1.Item(1).Text = strName
    Text1.Item(2).Text = strCaption

    If Len(strName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " ���ܺ��������ַ���" & strMsg & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    If Len(strCaption) = 0 Then
        MsgBox Label1.Item(2).Caption & " ����Ϊ�գ�", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(Text1.Item(2).Text)
        Exit Sub
    End If
    
    If Len(strType) = 0 Then
        MsgBox Label1.Item(3).Caption & " ����Ϊ�գ�", vbExclamation
        Combo1.Item(2).SetFocus
        Exit Sub
    End If
    
    If Len(strParent) = 0 Then
        MsgBox Label1.Item(4).Caption & " ����Ϊ�գ�", vbExclamation
        Combo1.Item(0).SetFocus
        Exit Sub
    End If
    
    strSQL = "SELECT FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID FROM tb_Test_Sys_Func " & _
             "Where FuncAutoID = " & strFID
    Set rsFunc = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsFunc.State = adStateClosed Then GoTo LineEnd
    If rsFunc.RecordCount = 0 Then
        strMsg = "�ù��������Ϣ�Ѷ�ʧ������ϵ����Ա��"
        GoTo LineBrk
    ElseIf rsFunc.RecordCount > 1 Then
        strMsg = "�ù��������Ϣ�쳣������ϵ����Ա��"
        GoTo LineBrk
    Else
        If strParent = strFID Then
            strMsg = Label1.Item(4).Caption & " �����Ǳ����ţ��������"
            Combo1.Item(0).SetFocus
            GoTo LineBrk
        End If
        
        If strName <> rsFunc.Fields("FuncName") Then blnName = True
        If strCaption <> rsFunc.Fields("FuncCaption") Then blnCaption = True
        If strType <> rsFunc.Fields("FuncType") Then blnType = True
        If Val(strParent) <> rsFunc.Fields("FuncParentID") Then blnParent = True
        
        If Not (blnName Or blnCaption Or blnType Or blnParent) Then
            strMsg = "û��ʵ���ԵĸĶ����������޸ģ�"
            GoTo LineBrk
        End If
        
        strMsg = "ȷ��Ҫ�޸�" & Label1.Item(0).Caption & "��" & strFID & "���Ĺ�����Ϣ��"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then GoTo LineEnd
        
        On Error GoTo LineErr
        
        If blnName Then rsFunc.Fields("FuncName") = strName
        If blnCaption Then rsFunc.Fields("FuncCaption") = strCaption
        If blnType Then rsFunc.Fields("FuncType") = strType
        If blnParent Then rsFunc.Fields("FuncParentID") = Val(strParent)
        rsFunc.Update
        rsFunc.Close
        
        strMsg = "�޸�ID��" & strFID & "����"
        If blnName Then strMsg = strMsg & "��" & Label1.Item(1).Caption & "��"
        If blnCaption Then strMsg = strMsg & "��" & Label1.Item(2).Caption & "��"
        If blnType Then strMsg = strMsg & "��" & Label1.Item(3).Caption & "��"
        If blnParent Then strMsg = strMsg & "��" & Label1.Item(4).Caption & "��"
        
        Call gsLogAdd(Me, udUpdate, "tb_Test_Sys_Func", strMsg)
        MsgBox "�ѳɹ�" & strMsg & "��", vbInformation
        
        Call msLoadFunc(TreeView1)
        
    End If
    
    GoTo LineEnd
    
LineBrk:
    rsFunc.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineErr:
    Call gsAlarmAndLog("������Ϣ�޸��쳣")
LineEnd:
    If rsFunc.State = adStateOpen Then rsFunc.Close
    Set rsFunc = Nothing
    
End Sub

Private Sub Command3_Click()
    '����
    
    Dim strFID As String, strTemp As String, strSQL As String, strMsg As String
    Dim cnRole As ADODB.Connection
    Dim rsRole As ADODB.Recordset
    Dim blnTran As Boolean
    Dim I As Long
    
    If Not mfCheckRoleFunc Then Exit Sub
    
    strTemp = Trim(Text1.Item(3).Text)
    strFID = Left(strTemp, InStr(strTemp, mTwoBar) - 1)
    
    If MsgBox("ȷ�����桾" & strTemp & "���Ľ�ɫָ����Ϣ��", vbQuestion + vbOKCancel, "����ѯ��") = vbCancel Then Exit Sub
    
    Set cnRole = New ADODB.Connection
    Set rsRole = New ADODB.Recordset
    cnRole.CursorLocation = adUseClient
    
    On Error GoTo LineErr
    
    cnRole.Open gID.CnString
    cnRole.BeginTrans
    blnTran = True
    
    'ɾ������Ȩ��
    strSQL = "DELETE FROM tb_Test_Sys_RoleFunc WHERE FuncAutoID =" & strFID
    cnRole.Execute strSQL
    
    '�����·���Ȩ��
    strSQL = "SELECT RoleAutoID ,FuncAutoID FROM tb_Test_Sys_RoleFunc WHERE FuncAutoID =" & strFID
    rsRole.Open strSQL, cnRole, adOpenStatic, adLockBatchOptimistic
    If rsRole.RecordCount > 0 Then
        strMsg = "��" & strTemp & "�� �ĺ�̨Ȩ����Ϣ�쳣�������Ի���ϵ����Ա��"
        GoTo LineBreak
    End If
    With TreeView2.Nodes
        For I = 2 To .Count
            If .Item(I).Checked And (Left(.Item(I).Key, Len(mKeyRole)) = mKeyRole) Then
                rsRole.AddNew
                rsRole.Fields("RoleAutoID") = Right(.Item(I).Key, Len(.Item(I).Key) - Len(mKeyRole))
                rsRole.Fields("FuncAutoID") = strFID
            End If
        Next
    End With
    rsRole.UpdateBatch
    cnRole.CommitTrans
    rsRole.Close
    cnRole.Close
    Call gsLogAdd(Me, udInsertBatch, "tb_Test_Sys_RoleFunc", "���桾" & strTemp & "���Ľ�ɫָ����Ϣ")
    MsgBox strTemp & " �Ľ�ɫָ����Ϣ����ɹ���", vbInformation
    
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
    
    Dim I As Long
    
    Me.Icon = gMDI.imgListCommandBars.ListImages("SysFunc").Picture
    Me.Caption = gMDI.cBS.Actions(gID.SysFunc).Caption
    Frame1.Item(0).Caption = Me.Caption
    
    For I = Text1.LBound To Text1.UBound
        Text1.Item(I).Text = ""
        If I < 3 Then Combo1.Item(I).ListIndex = -1
    Next
    TreeView1.Nodes.Clear
    TreeView2.Nodes.Clear
    TreeView1.ImageList = gMDI.imgListCommandBars
    TreeView2.ImageList = gMDI.imgListCommandBars
    
    Combo1.Item(2).Clear
    Combo1.Item(2).AddItem gID.FuncMainMenu
    Combo1.Item(2).AddItem gID.FuncForm
    Combo1.Item(2).AddItem gID.FuncButton
    Combo1.Item(2).AddItem gID.FuncControl
    
    Call msLoadFunc(TreeView1)
    Call msLoadDept(TreeView2)
    Call msLoadRole(TreeView2)
    
    Call gsLoadAuthority(Me, TreeView1)
    Call gsLoadAuthority(Me, Command1)
    Call gsLoadAuthority(Me, Command2)
    Call gsLoadAuthority(Me, Command3)
    
    
End Sub

Private Sub Form_Resize()

    Const conHeight As Long = 9000
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
    Dim strKey As String, strFID As String, strSQL As String, strMsg As String
    Dim rsFunc As ADODB.Recordset
    
    Text1.Item(3).Text = ""
    strKey = Node.Key
    lngLen = Len(strKey)
    If lngLen < Len(mKeyFunc) Then Exit Sub
    If Left(strKey, Len(mHeadKey)) = mHeadKey Then
        For I = Text1.LBound To Text1.UBound
            Text1.Item(I).Text = ""
            If I < 3 Then Combo1.Item(I).ListIndex = -1
        Next
        Exit Sub
    End If
    If Left(strKey, Len(mKeyFunc)) <> mKeyFunc Then Exit Sub
    
    strFID = Right(Node.Key, lngLen - Len(mKeyFunc))
    strSQL = "SELECT FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID FROM tb_Test_Sys_Func " & _
             "Where FuncAutoID = " & strFID & ""
    Set rsFunc = gfBackRecordset(strSQL)
    If rsFunc.State = adStateClosed Then GoTo LineEnd
    If rsFunc.RecordCount = 0 Then
        strMsg = "������Ϣ��ʧ������ϵ����Ա��"
        GoTo LineBreak
    ElseIf rsFunc.RecordCount > 1 Then
        strMsg = "������Ϣ�쳣������ϵ����Ա��"
        GoTo LineBreak
    Else
        Text1.Item(0).Text = strFID
        Text1.Item(1).Text = rsFunc.Fields("FuncName").Value
        Text1.Item(2).Text = rsFunc.Fields("FuncCaption").Value
        Text1.Item(3).Text = strFID & mTwoBar & rsFunc.Fields("FuncCaption")
        Combo1.Item(2).Text = rsFunc.Fields("FuncType").Value
        
        For I = 0 To Combo1.Item(1).ListCount - 1
            If rsFunc.Fields("FuncParentID").Value = Combo1.Item(1).List(I) Then
                Combo1.Item(0).ListIndex = I
                Exit For
            End If
        Next
        If I = Combo1.Item(1).ListCount Then
            Combo1.Item(0).ListIndex = IIf(Node.Parent.Key = mHeadKey, 0, -1)
        End If

        Node.SelectedImage = "FuncSelect"
    End If
    
    Call msLoadRoleFunc(strFID)
    
    GoTo LineEnd
    
LineBreak:
    rsFunc.Close
    MsgBox strMsg, vbExclamation
LineEnd:
    If rsFunc.State = adStateOpen Then rsFunc.Close
    Set rsFunc = Nothing
    
End Sub

Private Sub TreeView2_NodeCheck(ByVal Node As MSComctlLib.Node)
    Call gsNodeCheckCascade(Node, Node.Checked)
End Sub

