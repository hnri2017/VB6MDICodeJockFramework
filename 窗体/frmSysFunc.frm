VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysFunc 
   Caption         =   "功能设置"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   10035
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5175
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
         Index           =   2
         Left            =   1200
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   1080
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
         Index           =   2
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改功能信息"
         Height          =   495
         Left            =   3120
         TabIndex        =   5
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加功能"
         Height          =   495
         Left            =   1200
         TabIndex        =   4
         Top             =   3480
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
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   3720
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
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
         TabIndex        =   0
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
         TabIndex        =   3
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "功能标题"
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
         Left            =   250
         TabIndex        =   13
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "功能标识"
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
         Left            =   250
         TabIndex        =   12
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "自动编号"
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
         Left            =   250
         TabIndex        =   11
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "功能类别"
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
         Index           =   3
         Left            =   250
         TabIndex        =   10
         Top             =   1620
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "上级功能"
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
         Index           =   4
         Left            =   250
         TabIndex        =   9
         Top             =   2100
         Width           =   900
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4095
      Left            =   5640
      TabIndex        =   6
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
Attribute VB_Name = "frmSysFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mKeyFunc As String = "f"
Private Const mHeadKey As String = "kHeadKey"
Private Const mHeadText As String = "控制功能列表"




Private Function mfFuncTypeCheck(ByVal strType As String) As Boolean
    '检查功能类别是否正确
    
    Select Case strType
        Case gID.FuncButton, gID.FuncControl, gID.FuncForm
            mfFuncTypeCheck = True
        Else
    End Select
    
End Function

Private Sub msLoadFunc(ByRef tvwLoad As MSComctlLib.TreeView)
    '加载功能列表
    
    Dim rsFunc As ADODB.Recordset
    Dim strSQL As String
    Dim arrFunc() As String
    Dim I As Long, lngCount As Long
    Dim blnLoop As Boolean
    
    tvwLoad.Nodes.Clear
    tvwLoad.Nodes.Add , , mHeadKey, mHeadText, "FuncHead"   '添加首结点
    tvwLoad.Nodes(mHeadKey).Expanded = True     '展开结点
        
    strSQL = "SELECT FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID " & _
             "FROM tb_Test_Sys_Func ORDER BY FuncType ,FuncName " & _
             "SELECT FuncAutoID ,FuncCaption FROM tb_Test_Sys_Func " & _
             "WHERE FuncType ='" & gID.FuncForm & "' ORDER BY FuncCaption "
    Set rsFunc = gfBackRecordset(strSQL)
    If rsFunc.State = adStateClosed Then Exit Sub
    
    If rsFunc.RecordCount > 0 Then
        While Not rsFunc.EOF
            If rsFunc.Fields("FuncType") = gID.FuncForm Then
                tvwLoad.Nodes.Add mHeadKey, tvwChild, mKeyFunc & rsFunc.Fields("FuncAutoID"), rsFunc.Fields("FuncCaption"), "FuncForm"
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
    '必须与msLoadFunc过程配合使用来加载列表
    
    Dim arrOther() As String    '保存剩余的
    Dim blnOther As Boolean     '剩余标识
    Dim I As Long, J As Long, K As Long, lngCount As Long

    With tvwTree
        For J = LBound(arrLoad, 2) To UBound(arrLoad, 2)
            For I = 1 To .Nodes.Count   '注意此处下标从1开始
                If .Nodes.Item(I).Key = mKeyFunc & arrLoad(4, J) Then   ' FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID
                    .Nodes.Add .Nodes.Item(I).Key, tvwChild, mKeyFunc & arrLoad(0, J), arrLoad(2, J), IIf(arrLoad(3, J) = gID.FuncButton, "FuncButton", "FuncControl")
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
    
    If blnOther Then
        'Call msLoadFuncTree(tvwTree, arrOther)
        MsgBox mHeadText & "加载不完全，请通知管理员！", vbCritical
    End If

End Sub


Private Sub Combo1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        Combo1.Item(Index).ListIndex = -1
    End If
End Sub

Private Sub Command1_Click()
    '添加
    
    Dim strName As String, strCaption As String
    Dim strType As String, strParent As Variant
    Dim strSQL As String, strMsg As String
    Dim rsFunc As ADODB.Recordset
    
    If Combo1.Item(2).Text = gID.FuncForm Then
        Combo1.Item(0).Text = mHeadText
    End If
    If Combo1.Item(0).Text = mHeadText Then
        Combo1.Item(2).Text = gID.FuncForm
    End If
    
    strName = Trim(Text1.Item(1).Text)
    strCaption = Trim(Text1.Item(2).Text)
    strType = Combo1.Item(2).Text
    strParent = Combo1.Item(1).List(Combo1.Item(0).ListIndex)
    
    strName = Left(strName, 50)
    strCaption = Left(strCaption, 50)
    
    Text1.Item(1).Text = strName
    Text1.Item(2).Text = strCaption

    If Len(strName) = 0 Then
        MsgBox Label1.Item(1).Caption & " 不能为空！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strName)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " 不能含有特殊字符【" & strMsg & "】！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(strName)
        Exit Sub
    End If
    
    If Len(strCaption) = 0 Then
        MsgBox Label1.Item(2).Caption & " 不能为空！", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(strCaption)
        Exit Sub
    End If
    
    If Len(strType) = 0 Then
        MsgBox Label1.Item(3).Caption & " 不能为空！", vbExclamation
        Combo1.Item(2).SetFocus
        Exit Sub
    End If
    
    If Len(strParent) = 0 Then
        MsgBox Label1.Item(4).Caption & " 不能为空！", vbExclamation
        Combo1.Item(0).SetFocus
        Exit Sub
    End If
        
    If MsgBox("是否添加功能【" & strName & "】【" & strCaption & "】【" & strType & "】？", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
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
            strMsg = Label1.Item(1).Caption & "已存在，请更换！"
        Else
            strMsg = Label1.Item(1).Caption & " 在 " & Label1.Item(4).Caption & " 中已存在，请更换！"
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
        strMsg = "添加功能【" & strMsg & "】【" & strName & "】【" & strCaption & "】【" & strType & "】"
        Call gsLogAdd(Me, udInsert, "tb_Test_Sys_Func", strMsg)
        MsgBox "【" & strName & "】【" & strCaption & "】【" & strType & "】添加成功！", vbInformation
        Call msLoadFunc(TreeView1)
    End If
    
    GoTo LineEnd
    
LineBrk:
    rsFunc.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineErr:
    Call gsAlarmAndLog("添加功能异常")
LineEnd:
    If rsFunc.State = adStateOpen Then rsFunc.Close
    Set rsFunc = Nothing
    
End Sub

Private Sub Command2_Click()
    '修改
    
    Dim strFID As String, strName As String, strCaption As String
    Dim strType As String, strParent As String, strSQL As String, strMsg As String
    Dim blnName As Boolean, blnCaption As Boolean
    Dim blnType As Boolean, blnParent As Boolean
    Dim rsFunc As ADODB.Recordset
    
    If Combo1.Item(2).Text = gID.FuncForm Then
        Combo1.Item(0).Text = mHeadText
    End If
    If Combo1.Item(0).Text = mHeadText Then
        Combo1.Item(2).Text = gID.FuncForm
    End If
    
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
        MsgBox Label1.Item(1).Caption & " 不能为空！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    strMsg = gfStringCheck(strName)
    If Len(strMsg) > 0 Then
        MsgBox Label1.Item(1).Caption & " 不能含有特殊字符【" & strMsg & "】！", vbExclamation
        Text1.Item(1).SetFocus
        Text1.Item(1).SelStart = 0
        Text1.Item(1).SelLength = Len(Text1.Item(1).Text)
        Exit Sub
    End If
    
    If Len(strCaption) = 0 Then
        MsgBox Label1.Item(2).Caption & " 不能为空！", vbExclamation
        Text1.Item(2).SetFocus
        Text1.Item(2).SelStart = 0
        Text1.Item(2).SelLength = Len(Text1.Item(2).Text)
        Exit Sub
    End If
    
    If Len(strType) = 0 Then
        MsgBox Label1.Item(3).Caption & " 不能为空！", vbExclamation
        Combo1.Item(2).SetFocus
        Exit Sub
    End If
    
    If Len(strParent) = 0 Then
        MsgBox Label1.Item(4).Caption & " 不能为空！", vbExclamation
        Combo1.Item(0).SetFocus
        Exit Sub
    End If
    
    strSQL = "SELECT FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID FROM tb_Test_Sys_Func " & _
             "Where FuncAutoID = " & strFID
    Set rsFunc = gfBackRecordset(strSQL, adOpenStatic, adLockOptimistic)
    If rsFunc.State = adStateClosed Then GoTo LineEnd
    If rsFunc.RecordCount = 0 Then
        strMsg = "该功能相关信息已丢失，请联系管理员！"
        GoTo LineBrk
    ElseIf rsFunc.RecordCount > 1 Then
        strMsg = "该功能相关信息异常，请联系管理员！"
        GoTo LineBrk
    Else
        If strType = gID.FuncForm And strParent = strFID Then
            strMsg = Label1.Item(4).Caption & " 不能是本部门，请更换！"
            Combo1.Item(0).SetFocus
            GoTo LineBrk
        End If
        
        If strName <> rsFunc.Fields("FuncName") Then blnName = True
        If strCaption <> rsFunc.Fields("FuncCaption") Then blnCaption = True
        If strType <> rsFunc.Fields("FuncType") Then blnType = True
        If Val(strParent) <> rsFunc.Fields("FuncParentID") Then blnParent = True
        
        If Not (blnName Or blnCaption Or blnType Or blnParent) Then
            strMsg = "没有实质性的改动，不进行修改！"
            GoTo LineBrk
        End If
        
        strMsg = "确定要修改" & Label1.Item(0).Caption & "【" & strFID & "】的功能信息吗？"
        If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then GoTo LineEnd
        
        On Error GoTo LineErr
        
        If blnName Then rsFunc.Fields("FuncName") = strName
        If blnCaption Then rsFunc.Fields("FuncCaption") = strCaption
        If blnType Then rsFunc.Fields("FuncType") = strType
        If blnParent Then rsFunc.Fields("FuncParentID") = Val(strParent)
        rsFunc.Update
        rsFunc.Close
        
        strMsg = "修改ID【" & strFID & "】的"
        If blnName Then strMsg = strMsg & "【" & Label1.Item(1).Caption & "】"
        If blnCaption Then strMsg = strMsg & "【" & Label1.Item(2).Caption & "】"
        If blnType Then strMsg = strMsg & "【" & Label1.Item(4).Caption & "】"
        If blnParent Then strMsg = strMsg & "【" & Label1.Item(5).Caption & "】"
        
        Call gsLogAdd(Me, udUpdate, "tb_Test_Sys_Func", strMsg)
        MsgBox "已成功" & strMsg & "。", vbInformation
        
        Call msLoadFunc(TreeView1)
        
    End If
    
    GoTo LineEnd
    
LineBrk:
    rsFunc.Close
    MsgBox strMsg, vbExclamation
    GoTo LineEnd
LineErr:
    Call gsAlarmAndLog("功能信息修改异常")
LineEnd:
    If rsFunc.State = adStateOpen Then rsFunc.Close
    Set rsFunc = Nothing
    
End Sub

Private Sub Form_Load()
    
    Dim I As Long
    
    Me.Icon = frmSysMDI.imgListCommandBars.ListImages("SysFunc").Picture
    Me.Caption = frmSysMDI.cBS.Actions(gID.SysFunc).Caption
    
    For I = Text1.LBound To Text1.UBound
        Text1.Item(I).Text = ""
        Combo1.Item(I).ListIndex = -1
    Next
    TreeView1.Nodes.Clear
    TreeView1.ImageList = gMDI.imgListCommandBars
    
    Combo1.Item(2).Clear
    Combo1.Item(2).AddItem gID.FuncForm
    Combo1.Item(2).AddItem gID.FuncButton
    Combo1.Item(2).AddItem gID.FuncControl
    
    Call msLoadFunc(TreeView1)

    
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
    Dim strKey As String, strFID As String, strSQL As String, strMsg As String
    Dim rsFunc As ADODB.Recordset
    
    strKey = Node.Key
    lngLen = Len(strKey)
    If lngLen < Len(mKeyFunc) Then Exit Sub
    If Left(strKey, Len(mHeadKey)) = mHeadKey Then
        For I = Text1.LBound To Text1.UBound
            Text1.Item(I).Text = ""
            Combo1.Item(I).ListIndex = -1
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
        strMsg = "功能信息丢失，请联系管理员！"
        GoTo LineBreak
    ElseIf rsFunc.RecordCount > 1 Then
        strMsg = "功能信息异常，请联系管理员！"
        GoTo LineBreak
    Else
        Text1.Item(0).Text = strFID
        Text1.Item(1).Text = rsFunc.Fields("FuncName").Value
        Text1.Item(2).Text = rsFunc.Fields("FuncCaption").Value
        Combo1.Item(2).Text = rsFunc.Fields("FuncType").Value
        
        For I = 0 To Combo1.Item(1).ListCount - 1
            If rsFunc.Fields("FuncParentID").Value = Combo1.Item(1).List(I) Then
                Combo1.Item(0).ListIndex = I
                Exit For
            End If
        Next
        If I = Combo1.Item(1).ListCount Then
            Combo1.Item(0).ListIndex = IIf(rsFunc.Fields("FuncType") = gID.FuncForm, 0, -1)
        End If

        Node.SelectedImage = "FuncSelect"
    End If

    GoTo LineEnd
    
LineBreak:
    rsFunc.Close
    MsgBox strMsg, vbExclamation
LineEnd:
    If rsFunc.State = adStateOpen Then rsFunc.Close
    Set rsFunc = Nothing
    
End Sub
