VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysDepartment 
   Caption         =   "部门管理"
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
      TabIndex        =   1
      Top             =   240
      Width           =   5775
      Begin VB.CommandButton Command2 
         Caption         =   "修改部门"
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加部门"
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Text            =   "Combo2"
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   120
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "宋体"
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
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上级部门"
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   6
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部门名称"
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   3
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部门ID"
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   2
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
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSysDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mKeyHead As String = "k"


Private Sub msLoadDept(ByRef tvwDept As MSComctlLib.TreeView)
    '加载部门至TreeView控件中
    '要求：1、数据库中部门信息表Dept包含DeptID(Not Null)、DeptName(Not Null)、ParentID(Null)三个字段。
    '要求：2、部门表中只允许顶级部门（锁定为公司名称）的ParentID为Null，其它部门的ParentID都不能为Null。
    
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim arrDept() As String '注意下标要从0开始
    Dim I As Long, lngCount As Long, lngOneCompany As Long
    
    
    strSQL = "SELECT t1.DeptID ,t1.DeptName ,t1.ParentID ,t2.DeptName AS [ParentName] " & _
             "FROM tb_Test_Department AS [t1] " & _
             "LEFT JOIN tb_Test_Department AS [t2] " & _
             "ON t1.ParentID = t2.DeptID " & _
             "ORDER BY t1.ParentID ,t1.DeptName"    '注意字段顺序不可变
    Set rsDept = gfBackRecordset(strSQL)
    If rsDept.State = adStateClosed Then Exit Sub
    If rsDept.RecordCount > 0 Then
        
        tvwDept.Nodes.Clear
        Combo1.Item(0).Clear
        Combo1.Item(1).Clear
        
        While Not rsDept.EOF
            If IsNull(rsDept.Fields(3).Value) Then
                lngOneCompany = lngOneCompany + 1
                tvwDept.Nodes.Add , , mKeyHead & rsDept.Fields(0).Value, rsDept.Fields(1).Value
                tvwDept.Nodes.Item(1).Expanded = True
            Else
                lngCount = lngCount + 1
                ReDim Preserve arrDept(4, lngCount)
                For I = 0 To 3
                    arrDept(I, lngCount - 1) = rsDept.Fields(I).Value
                Next
            End If
            
            Combo1.Item(0).AddItem rsDept.Fields(1).Value
            Combo1.Item(1).AddItem rsDept.Fields(0).Value
            
            rsDept.MoveNext
        Wend
        
    End If
    rsDept.Close
    Set rsDept = Nothing
    
'''    If lngOneCompany = 1 Then
        Call msLoadDeptTree(tvwDept, arrDept)
'''    End If
    
End Sub

Private Sub msLoadDeptTree(ByRef tvwTree As MSComctlLib.TreeView, ByRef arrLoad() As String)
    '必须与msLoadDept过程配合使用来加载部门列表
    
    Dim arrOther() As String    '保存剩余的
    Dim blnOther As Boolean     '剩余标识
    Dim I As Long, J As Long, lngCount As Long
    Static C As Long
    
    With tvwTree
        For J = LBound(arrLoad, 2) To UBound(arrLoad, 2)
            C = C + 1
            For I = 1 To .Nodes.Count   '注意此处下标从1开始
                If .Nodes.Item(I).Key = mKeyHead & arrLoad(2, J) Then
                    .Nodes.Add I, tvwChild, mKeyHead & arrLoad(0, J), arrLoad(1, J)
                    .Nodes.Item(I + 1).Expanded = True
                    Exit For
                End If
            Next
            
            If I = .Nodes.Count Then
                blnOther = True
                lngCount = lngCount + 1
                ReDim Preserve arrOther(4, lngCount)
                For I = 0 To 3
                    arrOther(I, lngCount - 1) = arrLoad(I, J)
                Next
            End If
            
        Next
    End With
    
    If C > 10000000 Then Exit Sub
    
    If blnOther Then
        Call msLoadDeptTree(tvwTree, arrOther)
    End If

End Sub



Private Sub Command1_Click()
    '添加部门
    '要求数据库中部门表的部门ID(字段DeptID)是自增标识列
    
    Dim rsAdd As ADODB.Recordset
    Dim strSQL As String, strCheck As String, strMsg As String
    Dim strDeptName As String, strParentID As String
    Dim blnCompany As Boolean
    Dim lngDeptID As Long

    
    strDeptName = Trim(Text1.Item(1).Text)
    Text1.Item(1).Text = strDeptName
    
    If Len(strDeptName) = 0 Then
        MsgBox Label1.Item(1).Caption & " 不能为空字符！", vbExclamation
        Text1.Item(1).SetFocus
        Exit Sub
    End If
    
    strCheck = gfStringCheck(strDeptName)
    If Len(strCheck) > 0 Then
        MsgBox Label1.Item(1).Caption & "中不能包含特殊字符【" & strCheck & "】！", vbExclamation
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
        strMsg = "不勾选上级部门，则默认创建为公司名称！！！" & _
                 vbCrLf & "确定添加公司名称【" & strDeptName & "】吗？"
    Else
        strMsg = "确定添加部门名称【" & strDeptName & "】吗？"
    End If
    If MsgBox(strMsg, vbQuestion + vbOKCancel, "确认询问") = vbCancel Then Exit Sub
    
    strSQL = "SELECT DeptID ,DeptName ,ParentID FROM tb_Test_Department WHERE 1<>1 "
    Set rsAdd = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsAdd.State = adStateOpen Then
        rsAdd.AddNew
        rsAdd.Fields("DeptName") = strDeptName
        If Not blnCompany Then rsAdd.Fields("ParentID") = strParentID
        rsAdd.Update
        lngDeptID = rsAdd.Fields("DeptID")
        rsAdd.Close
        Set rsAdd = Nothing
        Call msLoadDept(TreeView1)
        MsgBox "部门【" & strDeptName & "】添加成功！", vbInformation
    End If
    Set rsAdd = Nothing
    
End Sub

Private Sub Command2_Click()
    '修改部门
    
    '不能修改到自己的子部门中
    
    Call msLoadDept(TreeView1)
    
End Sub

Private Sub Form_Load()
    
    Text1.Item(0).Text = ""
    Text1.Item(1).Text = ""
    TreeView1.Nodes.Clear
    
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
    If lngLen < Len(mKeyHead) Then Exit Sub
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
    
    Text1.Item(0).Text = Right(Node.Key, lngLen - Len(mKeyHead))
    Text1.Item(1).Text = Node.Text
    
End Sub
