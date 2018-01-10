VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Begin VB.Form frmMDB 
   Caption         =   "mdb"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   13635
   Begin VB.CommandButton Command1 
      Caption         =   "添加1000条"
      Height          =   500
      Left            =   12000
      TabIndex        =   6
      Top             =   4320
      Width           =   1300
   End
   Begin VB.CommandButton Command7 
      Caption         =   "创建表"
      Height          =   500
      Left            =   12000
      TabIndex        =   5
      Top             =   120
      Width           =   1300
   End
   Begin VB.CommandButton Command5 
      Caption         =   "查看记录"
      Height          =   500
      Left            =   12000
      TabIndex        =   4
      Top             =   3480
      Width           =   1300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "添加记录"
      Height          =   500
      Left            =   12000
      TabIndex        =   3
      Top             =   2760
      Width           =   1300
   End
   Begin VB.CommandButton Command4 
      Caption         =   "查看字段"
      Height          =   500
      Left            =   12000
      TabIndex        =   2
      Top             =   1800
      Width           =   1300
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   11535
      _cx             =   20346
      _cy             =   9340
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
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
   Begin VB.CommandButton Command2 
      Caption         =   "查看表名"
      Height          =   500
      Left            =   12000
      TabIndex        =   0
      Top             =   960
      Width           =   1300
   End
End
Attribute VB_Name = "frmMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    '添加1000条
    
    Dim rsAdd As ADODB.Recordset
    Dim cnAdd As ADODB.Connection
    Dim strSQL As String
    Dim I As Long, J As Long
    Dim sngTime As Single
    
    sngTime = Timer
    If VSFlexGrid1.Rows < 1000 Then Exit Sub
    
    Set cnAdd = New ADODB.Connection
    Set rsAdd = New ADODB.Recordset
    cnAdd.CursorLocation = adUseClient
    cnAdd.Open gVar.dbConn
    strSQL = "SELECT * FROM " & gVar.tbUser
    rsAdd.Open strSQL, cnAdd, adOpenStatic, adLockBatchOptimistic
    With VSFlexGrid1
        For I = 0 To 1000
            rsAdd.AddNew
            rsAdd.Fields(gVar.fdUserID) = .TextMatrix(I, 0)
            rsAdd.Fields(gVar.fdUserName) = .TextMatrix(I, 1)
            rsAdd.Fields(gVar.fdUserPassword) = .TextMatrix(I, 2)
            rsAdd.Fields(gVar.fdUserDeptID) = .TextMatrix(I, 3)
            rsAdd.Fields(gVar.fdUserFullName) = .TextMatrix(I, 4)
            rsAdd.Fields(gVar.fdUserSex) = .TextMatrix(I, 5)
            rsAdd.Fields(gVar.fdUserState) = .TextMatrix(I, 6)
            rsAdd.Fields(gVar.fdUserMemo) = Now
        Next
    End With
    rsAdd.UpdateBatch adAffectAll
    rsAdd.Close
    cnAdd.Close
    Set rsAdd = Nothing
    Set cnAdd = Nothing
    
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Clear
    sngTime = Format(Timer - sngTime, "0.000")
    
    MsgBox "OVER--" & sngTime & "s"
    
End Sub

Private Sub Command2_Click()
    '查看表名
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim K As Long
    Dim I As Long, C As Long
    Dim strTemp As String
    
    cn.CursorLocation = adUseClient
    cn.Open gVar.dbConn
    Set rs = cn.OpenSchema(adSchemaTables)
    
    If rs.State = adStateOpen Then
        rs.MoveFirst
        C = rs.Fields.Count - 1
        VSFlexGrid1.Clear
        VSFlexGrid1.Cols = rs.Fields.Count
        VSFlexGrid1.Rows = rs.RecordCount + 1
        For I = 0 To C
            VSFlexGrid1.TextMatrix(0, I) = rs.Fields(I).Name
        Next

        K = 1
        While Not rs.EOF
            For I = 0 To C
                VSFlexGrid1.TextMatrix(K, I) = rs.Fields(I).Value & ""
            Next
            rs.MoveNext
            K = K + 1
        Wend
        
    End If
    VSFlexGrid1.AutoSize 0, VSFlexGrid1.Cols - 1
    Set rs = Nothing
    Set cn = Nothing
    
End Sub

Private Sub Command3_Click()
    '添加一条记录
    
    Dim cnAdd As New ADODB.Connection
    Dim rsAdd As New ADODB.Recordset
    Dim strSQL As String

    cnAdd.Open gVar.dbConn
    rsAdd.CursorLocation = adUseClient
    strSQL = "SELECT * FROM " & gVar.tbUser

    
    rsAdd.Open strSQL, cnAdd, adOpenStatic, adLockOptimistic
    rsAdd.AddNew
    rsAdd.Fields(gVar.fdUserID) = "2001"
    rsAdd.Fields(gVar.fdUserName) = "xyz"
    rsAdd.Fields(gVar.fdUserPassword) = "123abc"
    rsAdd.Fields(gVar.fdUserFullName) = "小明"
    rsAdd.Fields(gVar.fdUserDeptID) = "1001"
    rsAdd.Fields(gVar.fdUserCreateMan) = "SystemAdmin"
    rsAdd.Fields(gVar.fdUserCreateTime) = Date
    
    rsAdd.Update
    
    Set rsAdd = Nothing
    Set cnAdd = Nothing
    
End Sub

Private Sub Command4_Click()
    '查看字段名
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim I As Long, C As Long, K As Long
    
    cn.CursorLocation = adUseClient
    cn.Open gVar.dbConn
    Set rs = cn.OpenSchema(adSchemaColumns)
    
    If rs.State = adStateOpen And rs.RecordCount > 0 Then
        rs.MoveFirst
        C = rs.Fields.Count - 1
        VSFlexGrid1.Clear
        VSFlexGrid1.Cols = rs.Fields.Count
        VSFlexGrid1.Rows = rs.RecordCount + 1
        For I = 0 To C
            VSFlexGrid1.TextMatrix(0, I) = rs.Fields(I).Name
        Next

        K = 1
        While Not rs.EOF
            For I = 0 To C
                VSFlexGrid1.TextMatrix(K, I) = rs.Fields(I).Value & ""
            Next
            rs.MoveNext
            K = K + 1
        Wend

        
    End If
    VSFlexGrid1.AutoSize 0, VSFlexGrid1.Cols - 1
    Set rs = Nothing
    Set cn = Nothing
End Sub

Private Sub Command5_Click()
    '查看记录
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim I As Long, C As Long, K As Long
    
    cn.CursorLocation = adUseClient
    cn.Open gVar.dbConn
    rs.Open "SELECT * FROM " & gVar.tbUser, cn, adOpenStatic, adLockReadOnly
    
    If rs.State = adStateOpen And rs.RecordCount > 0 Then
        rs.MoveFirst
        C = rs.Fields.Count - 1
        VSFlexGrid1.Clear
        VSFlexGrid1.Cols = rs.Fields.Count + 1
        VSFlexGrid1.Rows = rs.RecordCount + 1
        For I = 1 To C + 1
            VSFlexGrid1.TextMatrix(0, I) = rs.Fields(I - 1).Name
        Next

        K = 1
        While Not rs.EOF
            VSFlexGrid1.TextMatrix(K, 0) = K
            For I = 1 To C + 1
                VSFlexGrid1.TextMatrix(K, I) = rs.Fields(I - 1).Value & ""
            Next
            rs.MoveNext
            K = K + 1
        Wend
        
    End If
    
    VSFlexGrid1.AutoSize 0, VSFlexGrid1.Cols - 1
    
    Set rs = Nothing
    Set cn = Nothing
    
End Sub

Private Sub Command7_Click()
    '创建数据库与表
    
    Call gsRebuildDB
    MsgBox "OK"
    
End Sub

Private Sub Form_Load()

    Dim I As Long, J As Long
    With VSFlexGrid1
        .Editable = flexEDKbdMouse
        .Rows = 1001
        .Cols = 8
        J = 1001
        For I = 0 To 1000
            .TextMatrix(I, 0) = 2002 + I
            .TextMatrix(I, 1) = "xyz" & I
            .TextMatrix(I, 2) = "123abc" & I
            If I Mod 250 = 0 Then J = J + 1
            .TextMatrix(I, 3) = J
            .TextMatrix(I, 4) = "小明" & I
            .TextMatrix(I, 5) = IIf(I Mod 2 = 0, "男", "女")
            .TextMatrix(I, 6) = IIf(I Mod 2 = 0, "停用", "启用")
        Next
        .AutoSize 0, .Cols - 1
    End With
    
End Sub
