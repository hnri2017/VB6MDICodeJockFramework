VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysDepartment 
   Caption         =   "���Ź���"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   10155
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   4200
      TabIndex        =   7
      Top             =   240
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "�޸Ĳ�����Ϣ"
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��Ӳ���"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   1800
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
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   6
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
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   120
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
         TabIndex        =   3
         Top             =   1080
         Width           =   3375
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
         Index           =   2
         Left            =   200
         TabIndex        =   10
         Top             =   1140
         Width           =   900
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
         Index           =   1
         Left            =   200
         TabIndex        =   9
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����ID"
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
         Left            =   400
         TabIndex        =   8
         Top             =   180
         Width           =   690
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
Attribute VB_Name = "frmSysDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mKeyDept As String = "k"



Private Function mfIsChild(ByRef nodeDad As MSComctlLib.Node, ByVal strKey As String) As Boolean
    '�жϴ���Keyֵ�ǲ����Լ����ӽ��
    
    Dim I As Long, c As Long
    Dim nodeSon As MSComctlLib.Node
    
    c = nodeDad.Children
    If c = 0 Then Exit Function

    For I = 1 To c
        If I = 1 Then
            Set nodeSon = nodeDad.Child
        Else
            Set nodeSon = nodeSon.Next
        End If

'Debug.Print nodeSon.Text & "--" & nodeSon.Key

        If nodeSon.Key = strKey Then
            mfIsChild = True
            Exit Function
        End If
        If nodeSon.Children > 0 Then
            If mfIsChild(nodeSon, strKey) Then
                mfIsChild = True
                Exit Function
            End If
        End If
    Next

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
    
'''    If lngOneCompany = 1 Then    '���ֻ����һ����˾����
        If blnLoop Then Call msLoadDeptTree(tvwDept, arrDept)
'''    End If
    
End Sub

Private Sub msLoadDeptTree(ByRef tvwTree As MSComctlLib.TreeView, ByRef arrLoad() As String)
    '������msLoadDept�������ʹ�������ز����б�
    
    Dim arrOther() As String    '����ʣ���
    Dim blnOther As Boolean     'ʣ���ʶ
    Dim I As Long, J As Long, K As Long, lngCount As Long
    Static c As Long
    
    With tvwTree
'''        For I = LBound(arrLoad, 2) To UBound(arrLoad, 2)
'''            Debug.Print I & "--" & arrLoad(0, I) & "--" & arrLoad(1, I) & "--" & arrLoad(2, I) & "--" & arrLoad(3, I)
'''        Next
'''        Debug.Print "��������=" & (UBound(arrLoad, 2) - LBound(arrLoad, 2) + 1)
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
    
    c = c + 1
    If c > 64 Then Exit Sub '��ֹ�ݹ����̫��¶�ջ������������
    
    If blnOther Then
        Call msLoadDeptTree(tvwTree, arrOther)
    End If

End Sub



Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            Combo1.Item(Index).ListIndex = -1
        End If
    End If
End Sub

Private Sub Command1_Click()
    '��Ӳ���
    'Ҫ�����ݿ��в��ű�Ĳ���ID(�ֶ�DeptID)��������ʶ��
    
    Dim rsAdd As ADODB.Recordset
    Dim strSQL As String, strCheck As String, strMsg As String
    Dim strDeptName As String, strParentID As String
    Dim blnCompany As Boolean
    Dim lngDeptID As Long

    
    strDeptName = Trim(Text1.Item(1).Text)
    strDeptName = Left(strDeptName, 50)
    Text1.Item(1).Text = strDeptName
    
    If Len(strDeptName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ���ַ���", vbExclamation
        Text1.Item(1).SetFocus
        Exit Sub
    End If
    
    strCheck = gfStringCheck(strDeptName)
    If Len(strCheck) > 0 Then
        MsgBox Label1.Item(1).Caption & "�в��ܰ��������ַ���" & strCheck & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(0).SelLength = Len(strDeptName)
        Exit Sub
    End If
    
    If Combo1.Item(0).ListIndex = -1 Then
        blnCompany = True
    Else
        strParentID = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    End If
    
    If blnCompany Then
        strMsg = "����ѡ�ϼ����ţ���Ĭ�ϴ���Ϊ��˾���ƣ�����" & _
                 vbCrLf & "ȷ����ӹ�˾���ơ�" & strDeptName & "����"
    Else
        strMsg = "ȷ����Ӳ������ơ�" & strDeptName & "����"
    End If
    If MsgBox(strMsg, vbQuestion + vbOKCancel, "ȷ��ѯ��") = vbCancel Then Exit Sub
    
    strSQL = "SELECT DeptID ,DeptName ,ParentID FROM tb_Test_Sys_Department WHERE 1<>1 "
    Set rsAdd = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    
    On Error GoTo LineErr
    
    If rsAdd.State = adStateOpen Then
        rsAdd.AddNew
        rsAdd.Fields("DeptName") = strDeptName
        If Not blnCompany Then rsAdd.Fields("ParentID") = strParentID
        rsAdd.Update
        lngDeptID = rsAdd.Fields("DeptID")
        rsAdd.Close
        Call msLoadDept(TreeView1)
        Call gsLogAdd(Me, udInsert, "tb_Test_Sys_Department", "����²��š�" & lngDeptID & "����" & strDeptName & "��")
        MsgBox "���š�" & strDeptName & "����ӳɹ���", vbInformation
    End If
    
    GoTo LineEnd
    
LineErr:
    Call gsAlarmAndLog("��������쳣")
LineEnd:
    If rsAdd.State = adStateOpen Then rsAdd.Close
    Set rsAdd = Nothing
End Sub

Private Sub Command2_Click()
    '�޸Ĳ���
    
    Dim rsEdit As ADODB.Recordset
    Dim strSQL As String, strMsg As String, strCheck As String
    Dim strDeptID As String, strDeptName As String, strParentID As String, strLastPN As String
    Dim blnName As Boolean, blnParent As Boolean, blnCompany As Boolean
    
    
    strDeptID = Trim(Text1.Item(0).Text)
    strDeptName = Trim(Text1.Item(1).Text)
    Text1.Item(1).Text = strDeptName
    strParentID = Trim(Combo1.Item(1).List(Combo1.Item(0).ListIndex))
    
    If Len(strDeptID) = 0 Then
        MsgBox "����ѡ��һ�����ţ�", vbExclamation
        Exit Sub
    End If
    
    If Len(strDeptName) = 0 Then
        MsgBox Label1.Item(1).Caption & " ����Ϊ���ַ���", vbExclamation
        Text1.Item(1).SetFocus
        Exit Sub
    End If
    strCheck = gfStringCheck(strDeptName)
    If Len(strCheck) > 0 Then
        MsgBox Label1.Item(1).Caption & " ���ܺ��������ַ���" & strCheck & "����", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    If strDeptID = strParentID Then
        MsgBox Label1.Item(2).Caption & " �����Ǳ����ţ�", vbExclamation
        Exit Sub
    End If
    If Len(strParentID) = 0 Then blnCompany = True
    
    With TreeView1
        If .SelectedItem Is Nothing Then
            MsgBox "�ڲ�����쳣��������ѡ���ţ�", vbExclamation
            Exit Sub
        End If
        
        If .SelectedItem.Key <> mKeyDept & strDeptID Then
            MsgBox "�ڲ�����쳣��������ѡ��һ�����ţ�", vbExclamation
            Exit Sub
        End If
        
        If .SelectedItem.Text <> strDeptName Then blnName = True
        If .SelectedItem.Parent Is Nothing Then
            strLastPN = ""
            If Len(strParentID) > 0 Then blnParent = True
        Else
            strLastPN = .SelectedItem.Parent.Text
            If .SelectedItem.Parent.Key <> mKeyDept & strParentID Then blnParent = True
        End If
        
        If blnParent And (Len(strParentID) > 0) Then '�����޸ĵ��Լ����Ӳ�����
            If mfIsChild(.SelectedItem, mKeyDept & strParentID) Then
                MsgBox Label1.Item(2).Caption & " �����Ǳ����ŵ��Ӳ��ţ�", vbExclamation
                Exit Sub
            End If
        End If
                
        If Not (blnName Or blnParent) Then
            MsgBox "û��ʵ���ԵĸĶ��������޸ġ�", vbExclamation
            Exit Sub
        End If
        
        strMsg = "���ζ�IDΪ��" & strDeptID & "���Ĳ�����Ϣ�޸�������£�" & vbCrLf & vbCrLf & _
                 Space(6) & "�޸�λ��" & vbTab & "�޸�ǰ" & vbTab & vbTab & "�޸ĺ�" & vbCrLf & vbCrLf & _
                 Space(6) & Label1(1).Caption & vbTab & .SelectedItem.Text & vbTab & vbTab & strDeptName & vbCrLf & _
                 Space(6) & Label1(2).Caption & vbTab & strLastPN & vbTab & vbTab & Combo1.Item(0).Text & vbCrLf & vbCrLf & _
                 "�Ƿ��������޸ģ�"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
    End With
    
    strSQL = "SELECT DeptID ,DeptName ,ParentID " & _
             "FROM tb_Test_Sys_Department " & _
             "WHERE DeptID ='" & strDeptID & "'"
    Set rsEdit = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    
    On Error GoTo LineErr
    
    If rsEdit.State = adStateOpen Then
        If rsEdit.RecordCount = 1 Then
            strMsg = "��" & strDeptID & "��������Ϣ�޸ģ�"
            If blnName Then
                strMsg = strMsg & "��������[" & rsEdit.Fields("DeptName") & "]-->[" & strDeptName & "];"
                rsEdit.Fields("DeptName") = strDeptName
            End If
            If blnParent Then
                If blnCompany Then
                    rsEdit.Fields("ParentID") = Null
                Else
                    rsEdit.Fields("ParentID") = strParentID
                End If
                strMsg = strMsg & "�ϼ�����[" & strLastPN & "]-->[" & Combo1.Item(0).Text & "];"
            End If
            rsEdit.Update
            rsEdit.Close
            Call msLoadDept(TreeView1)
            Call gsLogAdd(Me, udUpdate, "tb_Test_Sys_Department", strMsg)
            MsgBox "������Ϣ�޸���ɣ�", vbInformation
        Else
            rsEdit.Close
            MsgBox "��̨�����쳣������ϵ����Ա��", vbCritical
        End If

    End If
    
    GoTo LineEnd
    
LineErr:
    Call gsAlarmAndLog("�����޸��쳣")
LineEnd:
    If rsEdit.State = adStateOpen Then rsEdit.Close
    Set rsEdit = Nothing
End Sub

Private Sub Form_Load()
        
    Me.Icon = frmSysMDI.imgListCommandBars.ListImages("SysDepartment").Picture
    Text1.Item(0).Text = ""
    Text1.Item(1).Text = ""
    TreeView1.Nodes.Clear
    TreeView1.ImageList = gMDI.imgListCommandBars
    
    Call msLoadDept(TreeView1)
    
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
    
    lngLen = Len(Node.Key)
    If lngLen < Len(mKeyDept) Then Exit Sub
    If Combo1.Item(0).ListCount < 1 Then Exit Sub
    
    If Node.Parent Is Nothing Then
        Combo1.Item(0).ListIndex = -1
    Else
        For I = 0 To Combo1.Item(0).ListCount - 1
            If Node.Parent.Text = Combo1.Item(0).List(I) Then
                Combo1.Item(0).ListIndex = I
                Exit For
            End If
        Next
    End If
    
    Text1.Item(0).Text = Right(Node.Key, lngLen - Len(mKeyDept))
    Text1.Item(1).Text = Node.Text
    
    Node.SelectedImage = "SysDepartment"
    
End Sub
