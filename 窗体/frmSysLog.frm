VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSysLog 
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   13530
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar Hsb 
      Height          =   255
      Left            =   11760
      TabIndex        =   26
      Top             =   5760
      Width           =   1455
   End
   Begin VB.VScrollBar Vsb 
      Height          =   1935
      Left            =   12840
      TabIndex        =   25
      Top             =   3720
      Width           =   255
   End
   Begin VB.Frame ctlMove 
      Caption         =   "Frame3"
      Height          =   6495
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   615
         Left            =   1320
         TabIndex        =   16
         Top             =   5880
         Width           =   8175
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   0
            Left            =   7320
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   1
            Left            =   360
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   2
            Left            =   1440
            TabIndex        =   20
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   3
            Left            =   2520
            TabIndex        =   19
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Index           =   4
            Left            =   3480
            TabIndex        =   18
            Top             =   120
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   270
            Left            =   6240
            TabIndex        =   17
            Text            =   "Text2"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   21
            Left            =   4560
            TabIndex        =   24
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   22
            Left            =   5520
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9855
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   2640
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   720
            TabIndex        =   8
            Text            =   "Combo1"
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   0
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Index           =   1
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   375
            Left            =   7680
            TabIndex        =   4
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   375
            Left            =   8640
            TabIndex        =   3
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   270
            Left            =   4560
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   840
            Width           =   615
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   5
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   92733441
            CurrentDate     =   42628
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   255
            Index           =   1
            Left            =   6000
            TabIndex        =   10
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Format          =   92733441
            CurrentDate     =   42628
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   1
            Left            =   5160
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Label1"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   11
            Top             =   840
            Width           =   975
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   2775
         Left            =   480
         TabIndex        =   15
         Top             =   2280
         Width           =   7215
         _cx             =   12726
         _cy             =   4895
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
End
Attribute VB_Name = "frmSysLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim lngPageSize As Long
Dim lngPageCount As Long
Dim lngPageCur As Long
Dim lngAddWith As Long
Dim rsLog As New ADODB.Recordset

Private Type typeInitialSize
    frmWidth As Long
    frmHeight As Long
    vsWidth As Long
    vsHeight As Long
    frameLeft As Long
    frameTop As Long
    rowHeight As Long
    pageSize As Long
End Type
Dim lngSize As typeInitialSize

Dim strLastTxt As String    '保存单元格编辑之前值

Private Sub Check1_Click()
    '时间
    Check1.ForeColor = IIf(Check1.Value, vbBlue, vbRed)
    DTPicker1.Item(0).Enabled = IIf(Check1.Value, True, False)
    DTPicker1.Item(1).Enabled = DTPicker1.Item(0).Enabled
    
End Sub


Private Sub Combo2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '清除类别
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        Combo2.Item(Index).ListIndex = -1
    End If
    
End Sub

Private Sub Command1_Click()
    '''查询

    Dim strSQL As String
    Dim strMen As String
    Dim strClass As String
    Dim strDateA As String
    Dim strDateB As String
    Dim strInfo As String
    Dim strCK As String
    
    strMen = Trim(Combo1.Text)
    strCK = gfStringCheck(strMen)
    If Len(strCK) > 0 Then
        MsgBox Label1(0).Caption & "中不能包含字符【" & strCK & "】！", vbExclamation, "敏感字符警告"
        Combo1.SetFocus
        Exit Sub
    End If
    
    strClass = Trim(Combo2.Item(0).Text)

    strInfo = Trim(Text1.Text)
    strCK = gfStringCheck(strInfo)
    If Len(strCK) > 0 Then
        MsgBox Label1(3).Caption & "中不能包含字符【" & strCK & "】！", vbExclamation, "敏感字符警告"
        Text1.SetFocus
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1.Text)
        Exit Sub
    End If
    
    If Check1.Value Then
        strDateA = Format(DTPicker1.Item(0).Value, "yyyy-MM-dd hh:mm:ss")
        strDateB = Format(DTPicker1.Item(1).Value, "yyyy-MM-dd hh:mm:ss")
    End If
    
    strSQL = "EXEC sp_Test_Sys_LogQuery '" & strClass & "','" & strInfo & "','" & _
            strDateA & "','" & strDateB & "','','" & strMen & "'"

    Set rsLog = gfBackRecordset(strSQL)

    lngPageCur = 1
    
    Call msShowValue
    Call msSetTable

End Sub
    
Private Sub Command2_Click()
    '''退出
    Unload Me
    
End Sub

Private Sub Command3_Click(Index As Integer)
    '''翻页
    Dim C As Long
    
    Select Case Index
        Case 1
            lngPageCur = 1
        Case 2
            lngPageCur = lngPageCur - 1
        Case 3
            lngPageCur = lngPageCur + 1
        Case 4
            lngPageCur = lngPageCount
        Case 0
            C = Val(Text2.Text)
            If C < 1 Then C = 1
            If C > lngPageCount Then C = lngPageCount
            lngPageCur = C
'            Text2.Text = CStr(c)
        Case Else
            Exit Sub
    End Select
    
    Call msShowValue
    
End Sub

Private Sub Form_Load()
    '主窗体加载
    Dim intAli As Integer
    Dim lngColor As Long

    With lngSize
        .frmWidth = 15800   '''初始化几个尺寸
        .frmHeight = 8900
        .vsWidth = 13800
        .vsHeight = 6000
        .rowHeight = 270
        .pageSize = 20
    End With
    
    intAli = 1
    lngPageSize = lngSize.pageSize
    lngColor = vbBlue
    
    Me.Icon = frmSysMDI.imgListCommandBars.ListImages("SysLog").Picture
    Me.Caption = frmSysMDI.cBS.Actions(gID.SysLog).Caption

    With Frame1 '查询条件
        .Move 120, 120, lngSize.vsWidth, 1200
        .Caption = "选择或输入搜索条件"
        .ForeColor = vbMagenta
        
        Label1.Item(0).Move 120, 300, 900, 255
        
        VSFlexGrid1.Move .Left, (.Top + .Height + 120), .Width, lngSize.vsHeight
        
    End With
    
    With Label1.Item(0)
        .Caption = "操作用户"
        .Alignment = intAli
        .ForeColor = lngColor
        
        Combo1.Move (.Left + .Width + 50), .Top - 30, .Width * 1.5
        
        Check1.Move (Combo1.Left + Combo1.Width + 500), .Top, .Width, .Height
        Check1.Caption = "时间段"
        Check1.Value = 1
        
        DTPicker1.Item(0).Move (Check1.Left + Check1.Width), .Top, 1300, .Height

        Label1.Item(2).Move .Left, (.Top + .Height + 200), .Width, .Height

    End With
    
    With DTPicker1.Item(0)
        .CustomFormat = "yyyy-MM-dd"
        .Format = dtpCustom
        .Value = Date
        
        Label1.Item(1).Caption = "--"
        Label1.Item(1).Move (.Left + .Width), .Top, 200, .Height
        
        DTPicker1.Item(1).Move (Label1(1).Left + Label1(1).Width), .Top, .Width, .Height
        DTPicker1.Item(1).CustomFormat = .CustomFormat
        DTPicker1.Item(1).Format = .Format
        DTPicker1.Item(1).Value = Date + 1
        
    End With
    
    With Label1.Item(2)
        .Caption = "操作类型"
        .Alignment = intAli
        .ForeColor = lngColor
        
        Combo2.Item(0).Move (Combo1.Left), .Top - 30, Combo1.Width
        Combo2.Item(1).Visible = False
        
        Label1.Item(3).Caption = "操作内容"
        .Alignment = intAli
        Label1.Item(3).Move Check1.Left, .Top, .Width, .Height
        Label1.Item(3).ForeColor = lngColor
        
        Text1.Text = ""
        Text1.Move DTPicker1(0).Left, .Top - 30, (DTPicker1(1).Left + DTPicker1(1).Width - DTPicker1(0).Left), .Height
        
    End With
    
    With Command1
        .Caption = "查询"
        .Height = 400
        .Move (Text1.Left + Text1.Width + 1000), (Text1.Top + Text1.Height - DTPicker1(1).Top - .Height) / 2 + DTPicker1(1).Top, 1000
        
        Command2.Caption = "退出"
        Command2.Move (.Left + .Width + 3000), .Top, .Width, .Height
    End With
    
    With VSFlexGrid1
        .Appearance = flexFlat
        .BorderStyle = flexBorderNone
        .GridLines = flexGridInset
        .FixedCols = 1
        .FixedRows = 1
        .Cols = 9
        .FormatString = "^序号|^操作用户|^操作时间|^操作类型|操作内容|窗口名称|^电脑IP|电脑名称|系统表名"
        .AllowUserResizing = flexResizeColumns
        .AllowUserFreezing = flexFreezeColumns
        .Editable = flexEDKbdMouse
        .BackColorBkg = Me.BackColor
        .BackColorFixed = RGB(121, 151, 219)
        .BackColorAlternate = RGB(250, 235, 215)
        
        Frame2.Top = (.Top + .Height)
        
    End With
        
    With Command3.Item(1)
        .Caption = "第一页"
        .Move 120, 120, 800, 375
        
        Command3.Item(2).Caption = "上一页"
        Command3.Item(2).Move (.Left + .Width), .Top, .Width, .Height
        
        Command3.Item(3).Caption = "下一页"
        Command3.Item(3).Move (.Left + .Width * 2), .Top, .Width, .Height
        
        Command3.Item(4).Caption = "最后页"
        Command3.Item(4).Move (.Left + .Width * 3), .Top, .Width, .Height
        
        Label1.Item(21).Caption = "共    页"
        Label1.Item(21).Move (.Left + .Width * 4), .Top + 100, .Width * 1.5, .Height

        
        Label1.Item(22).Caption = "跳至第         页"
        Label1.Item(22).Move (.Left + .Width * 5.5), Label1(21).Top, .Width * 2, .Height
        Label1.Item(22).ForeColor = vbMagenta
        
        Text2.Move (.Left + .Width * 6.17), Label1(21).Top - 30, .Width, 255
        Text2.Text = ""
        Text2.Alignment = 2
        
        Command3.Item(0).Caption = "跳转"
        Command3.Item(0).Move (.Left + .Width * 7.5), .Top, .Width, .Height
        
    End With

    With Frame2     '翻页键框架
        .Caption = ""
        .BorderStyle = 0
        .Width = Command3.Item(0).Left + Command3.Item(3).Width + 120
        .Height = Command3.Item(0).Top + Command3.Item(0).Height + 120
        .Left = VSFlexGrid1.Left + (VSFlexGrid1.Width - .Width) / 2
        lngSize.frameLeft = .Left
        lngSize.frameTop = .Top
    End With

    Me.Move 0, 0, lngSize.frmWidth, lngSize.frmHeight
    ctlMove.BorderStyle = 0
    ctlMove.Move 120, 120, 25000, 20000
    
    Call msSetTable
    Call msLoadMen
    

    For lngColor = udSelect To udUpdateBatch
        Combo2.Item(0).AddItem gfBackLogType(lngColor)
    Next
    
    Call gsLoadAuthority(Me, Command1)
    
End Sub

Private Sub Form_Resize()

    Dim lngW As Long
    Dim lngH As Long
    Dim lngVar As Long
    
    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 14400, 9000)
    
    If gMDI.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState = vbMinimized Then Exit Sub

    lngW = Me.Width
    lngH = Me.Height

    If lngW > lngSize.frmWidth Then     '当宽度变化
        lngVar = lngW - lngSize.frmWidth
    Else
        lngVar = 0
    End If
    Frame1.Width = lngSize.vsWidth + lngVar
    VSFlexGrid1.Width = Frame1.Width

    Frame2.Left = lngSize.frameLeft + lngVar / 2
    lngAddWith = lngVar

    If lngH > lngSize.frmHeight Then    '当高度变化
        lngVar = lngH - lngSize.frmHeight
    Else
        lngVar = 0
    End If
    VSFlexGrid1.Height = lngSize.vsHeight + lngVar
    Frame2.Top = lngSize.frameTop + lngVar
    lngPageSize = lngSize.pageSize + Int(lngVar / lngSize.rowHeight)
    VSFlexGrid1.Rows = VSFlexGrid1.FixedRows + lngPageSize
    
    If Len(VSFlexGrid1.TextMatrix(VSFlexGrid1.FixedRows, 0)) > 0 Then Call msShowValue '表格重新赋值
    
End Sub


Private Sub Hsb_Change()
    ctlMove.Left = -Hsb.Value
End Sub

Private Sub Hsb_Scroll()
    Call Hsb_Change    '也可不添加此Scroll事件，以下同。
End Sub

Private Sub Vsb_Change()
    ctlMove.Top = -Vsb.Value
End Sub

Private Sub Vsb_Scroll()
    Call Vsb_Change
End Sub

Private Sub msSetTable()
    '''设置表格格式
    With VSFlexGrid1
        .Redraw = flexRDNone
        
        .Rows = lngPageSize + .FixedRows
        .rowHeight(-1) = lngSize.rowHeight
        .rowHeight(0) = 400
        .ColWidth(0) = 650
        .ColWidth(1) = 1100
        .ColWidth(2) = 1900
        .ColWidth(3) = 1000
        .ColWidth(4) = 4800 + lngAddWith / 3
        .ColWidth(5) = 1200 + lngAddWith / 3
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
'        .ColWidth(8) = 2000 + lngAddWith / 3
        .ColWidth(8) = 0
        .Redraw = flexRDBuffered    'flexRDDirect,flexRDBuffered
    End With
    
End Sub

Private Sub msLoadMen()
    '''加载操作人列表
    
    Dim strSQL As String
    Dim rsL As ADODB.Recordset
    
    strSQL = "SELECT DISTINCT RIGHT(LogUserFullName,LEN(LogUserFullName)-" & _
            "CHARINDEX(',',LogUserFullName)) AS [LogUserFullName] FROM tb_Test_Sys_OperationLog"
    Set rsL = gfBackRecordset(strSQL)
    
    If rsL.State = adStateClosed Then Exit Sub
    If Not (rsL.BOF And rsL.EOF) Then
        With Combo1
            .Clear
            While Not rsL.EOF
                .AddItem rsL.Fields("LogUserFullName")
                rsL.MoveNext
            Wend
        End With
    End If
    Set rsL = Nothing
    
End Sub

Private Sub msShowValue()
    
    Dim I As Long
    Dim K As Long
    Dim N As Long
    Dim W As Long
    
    If rsLog.State = adStateClosed Then Exit Sub
    
    rsLog.pageSize = lngPageSize
    lngPageCount = rsLog.PageCount
        
    If rsLog.RecordCount = 0 Then
        MsgBox "没有符合条件的内容！", vbExclamation, "空值反馈"
        lngPageCur = 0
    Else
        If lngPageCur > lngPageCount Then lngPageCur = lngPageCount '''规范当前页码
        If lngPageCur < 1 Then lngPageCur = 1
        rsLog.AbsolutePage = lngPageCur
        
        N = lngPageSize * (lngPageCur - 1) + 1  '''第一条记录的序号
        
        With VSFlexGrid1
            For I = 1 To lngPageSize    '''将指定页的内容赋值到表格中
                If rsLog.EOF Then Exit For
                K = .FixedRows - 1 + I
                .TextMatrix(K, 0) = CStr(N)
                .TextMatrix(K, 1) = rsLog.Fields("LogUserFullName")
                .TextMatrix(K, 2) = rsLog.Fields("LogTime")
                .TextMatrix(K, 3) = rsLog.Fields("LogType")
                .TextMatrix(K, 4) = rsLog.Fields("LogContent")
                W = InStr(rsLog.Fields("LogFormName"), ",")
'                If W < 1 Then W = Len(rsLog.Fields("LogFormName"))
                .TextMatrix(K, 5) = Right(rsLog.Fields("LogFormName"), Len(rsLog.Fields("LogFormName")) - W)
                .TextMatrix(K, 6) = rsLog.Fields("LogPCIP")
                .TextMatrix(K, 7) = rsLog.Fields("LogPCName")
                .TextMatrix(K, 8) = rsLog.Fields("LogTable")
                N = N + 1
                rsLog.MoveNext
            Next
            
            If I < lngPageSize + 1 Then '''如果赋值内容不足一页，则清空表格中最后一条记录后面的内容
                For I = I To lngPageSize
                    K = .FixedRows - 1 + I
                    If Len(.TextMatrix(K, 0)) = 0 Then Exit For
                    For N = 0 To .Cols - 1
                        .TextMatrix(K, N) = ""
                    Next
                Next
            End If
            
        End With
    End If
    
    Command3.Item(1).Enabled = IIf(lngPageCur < 2, False, True)     '''设置4个翻页按钮的可用状态
    Command3.Item(2).Enabled = Command3.Item(1).Enabled
    Command3.Item(3).Enabled = IIf(lngPageCur = lngPageCount, False, True)
    Command3.Item(4).Enabled = Command3.Item(3).Enabled
    
    Label1.Item(21).Caption = "共 " & CStr(lngPageCount) & " 页"
    Text2.Text = CStr(lngPageCur)
    
End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal row As Long, ByVal col As Long)
    VSFlexGrid1.TextMatrix(row, col) = strLastTxt
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
    strLastTxt = VSFlexGrid1.TextMatrix(row, col)
End Sub


