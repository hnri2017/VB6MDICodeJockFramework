VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysRole 
   Caption         =   "角色设置"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   10650
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
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
         TabIndex        =   6
         Top             =   1080
         Width           =   3375
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
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Text            =   "Combo2"
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
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
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加角色"
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改角色信息"
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "角色标识"
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
         Left            =   200
         TabIndex        =   9
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "角色名称"
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
         TabIndex        =   8
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所属部门"
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
         TabIndex        =   7
         Top             =   1140
         Width           =   900
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4095
      Left            =   6120
      TabIndex        =   10
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
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Private Const mOtherText As String = "其他角色"



Private Function mfIsChild(ByRef nodeDad As MSComctlLib.Node, ByVal strKey As String) As Boolean
    '判断传入Key值是不是自己的子结点
    
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
    '加载部门至TreeView控件中
    '要求：1、数据库中部门信息表Dept包含DeptID(Not Null)、DeptName(Not Null)、ParentID(Null)三个字段。
    '要求：2、部门表中只允许顶级部门（锁定为公司名称）的ParentID为Null，其它部门的ParentID都不能为Null。
    
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim arrDept() As String '注意下标要从0开始
    Dim I As Long, lngCount As Long, lngOneCompany As Long
    Dim blnLoop As Boolean
    
    
    strSQL = "SELECT t1.DeptID ,t1.DeptName ,t1.ParentID ,t2.DeptName AS [ParentName] " & _
             "FROM tb_Test_Sys_Department AS [t1] " & _
             "LEFT JOIN tb_Test_Sys_Department AS [t2] " & _
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
    '必须与msLoadDept过程配合使用来加载部门列表
    
    Dim arrOther() As String    '保存剩余的
    Dim blnOther As Boolean     '剩余标识
    Dim I As Long, J As Long, K As Long, lngCount As Long
    Static c As Long
    
    With tvwTree
        For J = LBound(arrLoad, 2) To UBound(arrLoad, 2)
            For I = 1 To .Nodes.Count   '注意此处下标从1开始
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
    If c > 64 Then Exit Sub '防止递归层数太深导致堆栈溢出而程序崩溃
    
    If blnOther Then
        Call msLoadDeptTree(tvwTree, arrOther)
    End If

End Sub

Private Sub msLoadRole(ByRef tvwUser As MSComctlLib.TreeView)
    '加载角色，前提是已加载好部门
    
    Dim strSQL As String
    Dim rsRole As ADODB.Recordset
    Dim arrOther() As String    '保存剩余的
    Dim blnOther As Boolean     '剩余标识
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


Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 Then
        If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
            Combo1.Item(Index).ListIndex = -1
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Me.Icon = gMDI.imgListCommandBars.ListImages("SysRole").Picture
    Me.Caption = gMDI.cBS.Actions(gID.SysRole).Caption
    
    Text1.Item(0).Text = ""
    Text1.Item(1).Text = ""
    TreeView1.Nodes.Clear
    TreeView1.ImageList = gMDI.imgListCommandBars
    
    Call msLoadDept(TreeView1)
    Call msLoadRole(TreeView1)  '角色后加载
    
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

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

    Dim lngLen As Long, I As Long
    Dim strKey As String, strUID As String, strSQL As String, strMsg As String
    Dim rsRole As ADODB.Recordset
    
    strKey = Node.Key
    lngLen = Len(strKey)
    If lngLen < Len(mKeyRole) Then Exit Sub
    If Left(strKey, Len(mKeyDept)) = mKeyDept Then
        For mlngID = Text1.LBound To Text1.UBound
            Text1.Item(mlngID).Text = ""
        Next
        Option1.Item(1).Value = True
        Combo1.Item(0).ListIndex = -1
        Exit Sub
    End If
    If Left(strKey, Len(mKeyRole)) <> mKeyRole Then Exit Sub
    
    strUID = Right(Node.Key, lngLen - Len(mKeyRole))
    strSQL = "EXEC sp_Test_Sys_UserInfo '" & strUID & "'"
    Set rsRole = gfBackRecordset(strSQL)
    If rsRole.State = adStateClosed Then GoTo LineEnd
    If rsRole.RecordCount = 0 Then
        strMsg = "用户信息丢失了，请联系管理员！"
        rsRole.Close
        GoTo LineBreak
    ElseIf rsRole.RecordCount > 1 Then
        strMsg = "用户信息异常，请联系管理员！"
        rsRole.Close
        GoTo LineBreak
    Else
        Text1.Item(0).Text = strUID
        Text1.Item(1).Text = rsRole.Fields("UserLoginName").Value & ""
        Text1.Item(2).Text = gfDecryptSimple(rsRole.Fields("UserPassword").Value & "")
        Text1.Item(3).Text = rsRole.Fields("UserFullName").Value & ""
        Text1.Item(4).Text = rsRole.Fields("UserMemo").Value & ""
        
        Option1.Item(0).Value = IIf(rsRole.Fields("UserSex").Value = "女", True, False)
        Option1.Item(1).Value = IIf(rsRole.Fields("UserSex").Value = "男", True, False)
        
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

        Node.SelectedImage = "SelectedMen"
    End If

    GoTo LineEnd
    
LineBreak:
    MsgBox strMsg, vbExclamation
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    
End Sub
