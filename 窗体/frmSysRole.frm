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

Private Sub Command1_Click()
    '添加角色

    Dim rsRole As ADODB.Recordset
    Dim strRoleName As String, strSQL As String, strCheck As String, strMsg As String
    Dim strDeptID As Variant
    Dim lngRID As Long

    
    strRoleName = Trim(Text1.Item(1).Text)
    strRoleName = Left(strRoleName, 50)
    Text1.Item(1).Text = strRoleName
    
    If Len(strRoleName) = 0 Then
        MsgBox Label1.Item(1).Caption & " 不能为空字符！", vbExclamation
        Text1.Item(1).SetFocus
        Exit Sub
    End If
    
    strCheck = gfStringCheck(strRoleName)
    If Len(strCheck) > 0 Then
        MsgBox Label1.Item(1).Caption & "中不能包含特殊字符【" & strCheck & "】！", vbExclamation
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
    
    strMsg = "确定添加" & Label1.Item(1).Caption & "【" & strRoleName & "】吗？"
    If MsgBox(strMsg, vbQuestion + vbOKCancel, "确认询问") = vbCancel Then Exit Sub
    
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
        Call gsLogAdd(Me, udInsert, "tb_Test_Sys_Role", "添加新角色【" & lngRID & "】【" & strRoleName & "】")
        MsgBox "角色【" & strRoleName & "】添加成功！", vbInformation
    End If
    
    GoTo LineEnd
    
LineErr:
    Call gsAlarmAndLog("角色添加异常")
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing

End Sub

Private Sub Command2_Click()
    '修改角色
    
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
        MsgBox "请先选择一个角色！", vbExclamation
        Exit Sub
    End If
    
    If Len(strRoleName) = 0 Then
        MsgBox Label1.Item(1).Caption & " 不能为空字符！", vbExclamation
        Text1.Item(1).SetFocus
        Exit Sub
    End If
    strCheck = gfStringCheck(strRoleName)
    If Len(strCheck) > 0 Then
        MsgBox Label1.Item(1).Caption & " 不能含有特殊字符【" & strCheck & "】！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strRoleName)
        Exit Sub
    End If

    With TreeView1
        If .SelectedItem Is Nothing Then
            MsgBox "内部检测异常，请重新选择角色！", vbExclamation
            Exit Sub
        End If
        
        If .SelectedItem.Key <> mKeyRole & strRID Then
            MsgBox "内部检测异常，请重新选择一个角色！", vbExclamation
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
            MsgBox "没有实质性的改动，不作修改。", vbExclamation
            Exit Sub
        End If
        
        strMsg = "本次对ID为【" & strRID & "】的角色信息修改情况如下：" & vbCrLf & vbCrLf & _
                 Space(6) & "修改位置" & vbTab & "修改前" & vbTab & vbTab & "修改后" & vbCrLf & vbCrLf & _
                 Space(6) & Label1(1).Caption & vbTab & .SelectedItem.Text & vbTab & vbTab & strRoleName & vbCrLf & _
                 Space(6) & Label1(2).Caption & vbTab & strLastDept & vbTab & vbTab & Combo1.Item(0).Text & vbCrLf & vbCrLf & _
                 "是否对其进行修改？"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
    End With
    
    strSQL = "SELECT RoleAutoID ,RoleName ,DeptID " & _
             "FROM tb_Test_Sys_Role " & _
             "WHERE RoleAutoID =" & strRID & ""
    Set rsRole = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    
    On Error GoTo LineErr
    
    If rsRole.State = adStateOpen Then
        If rsRole.RecordCount = 1 Then
            strMsg = "【" & strRID & "】角色信息修改："
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
            MsgBox "角色信息修改完成！", vbInformation
        Else
            rsRole.Close
            MsgBox "后台数据异常，请联系管理员！", vbCritical
        End If

    End If
    
    GoTo LineEnd
    
LineErr:
    Call gsAlarmAndLog("角色修改异常")
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    
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

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim lngLen As Long, I As Long
    Dim strKey As String, strRID As String, strSQL As String, strMsg As String
    Dim rsRole As ADODB.Recordset
    
    strKey = Node.Key
    lngLen = Len(strKey)
    If lngLen < Len(mKeyRole) Then Exit Sub
    If Left(strKey, Len(mKeyDept)) = mKeyDept Then
        Text1.Item(0).Text = ""
        Text1.Item(1).Text = ""
        Combo1.Item(0).ListIndex = -1
        Exit Sub
    End If
    If Left(strKey, Len(mKeyRole)) <> mKeyRole Then Exit Sub
    
    strRID = Right(Node.Key, lngLen - Len(mKeyRole))
    strSQL = "SELECT RoleAutoID ,RoleName ,DeptID FROM tb_Test_Sys_Role " & _
             "WHERE RoleAutoID =" & strRID
    Set rsRole = gfBackRecordset(strSQL)
    If rsRole.State = adStateClosed Then GoTo LineEnd
    If rsRole.RecordCount = 0 Then
        strMsg = "角色信息丢失了，请联系管理员！"
        GoTo LineBreak
    ElseIf rsRole.RecordCount > 1 Then
        strMsg = "角色信息异常，请联系管理员！"
        GoTo LineBreak
    Else
        Text1.Item(0).Text = strRID
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

    GoTo LineEnd
    
LineBreak:
    rsRole.Close
    MsgBox strMsg, vbExclamation
LineEnd:
    If rsRole.State = adStateOpen Then rsRole.Close
    Set rsRole = Nothing
    
End Sub
