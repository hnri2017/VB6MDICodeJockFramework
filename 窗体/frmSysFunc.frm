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
      TabIndex        =   0
      Top             =   120
      Width           =   5175
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
         TabIndex        =   14
         Top             =   1560
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改功能信息"
         Height          =   495
         Left            =   3120
         TabIndex        =   7
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加功能"
         Height          =   495
         Left            =   1200
         TabIndex        =   6
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
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   3720
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   2040
         Width           =   3375
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
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text2"
         Top             =   1080
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         Index           =   5
         Left            =   250
         TabIndex        =   8
         Top             =   2100
         Width           =   900
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4095
      Left            =   5640
      TabIndex        =   13
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
Private Const mHeadKey As String = "fHeadKey"
Private Const mHeadText As String = "控制功能列表"




Private Sub msLoadFunc(ByRef tvwLoad As MSComctlLib.TreeView)
    '加载功能列表
    
    Dim rsFunc As ADODB.Recordset
    Dim strSQL As String
    Dim arrFunc() As String
    Dim I As Long, lngCount As Long
    Dim blnLoop As Boolean
    
    
    strSQL = "SELECT FuncAutoID ,FuncName ,FuncCaption ,FuncType ,FuncParentID " & _
             "FROM tb_Test_Sys_Func ORDER BY FuncType ,FuncName " & _
             "SELECT DISTINCT FuncType FROM tb_Test_Sys_Func ORDER BY FuncType " & _
             "SELECT FuncAutoID ,FuncCaption FROM tb_Test_Sys_Func " & _
             "WHERE FuncType ='" & gID.FuncForm & "' ORDER BY FuncCaption "
    Set rsFunc = gfBackRecordset(strSQL)
    If rsFunc.State = adStateClosed Then Exit Sub
    
    If rsFunc.RecordCount > 0 Then
        tvwLoad.Nodes.Clear
        tvwLoad.Nodes.Add , , mHeadKey, mHeadText, "FuncHead"   '添加首结点
        tvwLoad.Nodes(mHeadKey).Expanded = True     '展开结点
        
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
        Combo1.Item(2).Clear
        If rsFunc.RecordCount > 0 Then
            While Not rsFunc.EOF
                Combo1.Item(2).AddItem rsFunc.Fields("FuncType")
                rsFunc.MoveNext
            Wend
        End If
    End If
    
    Set rsFunc = rsFunc.NextRecordset
    If rsFunc.State = adStateOpen Then
    
        Combo1.Item(0).Clear
        Combo1.Item(1).Clear
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
