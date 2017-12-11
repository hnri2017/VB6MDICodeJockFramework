VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmForm4 
   Caption         =   "测试窗口4"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   9555
   Begin VB.VScrollBar Vsb 
      Height          =   6255
      Left            =   9000
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.HScrollBar Hsb 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6240
      Width           =   9375
   End
   Begin VB.Frame ctlMove 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.CommandButton Command3 
         Caption         =   "访问表"
         Height          =   495
         Left            =   1080
         TabIndex        =   12
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "FillGrid"
         Height          =   495
         Left            =   1080
         TabIndex        =   11
         Top             =   2880
         Width           =   1095
      End
      Begin FlexCell.Grid Grid1 
         Height          =   4935
         Left            =   2760
         TabIndex        =   10
         Top             =   1320
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   8705
         Cols            =   5
         GridColor       =   12632256
         Rows            =   30
      End
      Begin VB.Timer timeProgress 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1800
         Top             =   0
      End
      Begin VB.CommandButton Command8 
         Caption         =   "GridToExcel"
         Height          =   495
         Left            =   960
         TabIndex        =   7
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "进度条测试"
         Height          =   495
         Left            =   6600
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "加密"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "解密"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   390
         Left            =   2880
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   0
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim mstbPaneBar As XtremeCommandBars.StatusBarProgressPane
Dim mstbPaneText As XtremeCommandBars.StatusBarPane
Dim mlngMax As Long, mlngC As Long


Private Sub Command1_Click()
    '加密输入，生成密文
    
    Dim strText As String

    strText = Text1.Text
    If Len(strText) > 20 Then
        MsgBox "输入长度不能超过20个字符，超出部分已被删除！", vbExclamation, "长度警告"
        strText = Left(strText, 20)
        Text1.Text = strText
    End If
    Text2.Text = gfEncryptSimple(strText)
    
'    Dim strA As String
'    strA = Text1.Text
'    Text2.Text = gfAsciiAdd(strA)
Debug.Print Len(Text2), Text2

End Sub

Private Sub Command2_Click()
    '解密密文，还原成明文

    Text3.Text = gfDecryptSimple(Text2.Text)
        
'    Dim strB As String
'    strB = Text2.Text
'    Text1.Text = gfAsciiSub(strB)
End Sub

Private Sub Command3_Click()
    Dim strSQL As String, strPWD As String
    Dim rsT As ADODB.Recordset
    
    strSQL = "SELECT * FROM tb_Test_User"
    Set rsT = gfBackRecordset(strSQL)
    
    If rsT.State = adStateOpen Then
        If rsT.RecordCount > 0 Then
            While Not rsT.EOF
                
                
                rsT.MoveNext
            Wend
          
        End If
        rsT.Close
    End If
    Set rsT = Nothing
End Sub

Private Sub Command4_Click()
    
    Set mstbPaneBar = gMDI.cBS.StatusBar.FindPane(gID.StatusBarPaneProgress)
    Set mstbPaneText = gMDI.cBS.StatusBar.FindPane(gID.StatusBarPaneProgressText)
    mlngMax = 21474830
    
    With mstbPaneBar
        .Min = 0
        .Max = mlngMax
        .Value = 0
    End With
    
    
    timeProgress.Enabled = True
    Me.Enabled = False
    For mlngC = 0 To mlngMax
        DoEvents
    Next
    
End Sub

Private Sub Command7_Click()
    Dim I As Long
    With Grid1
        .Rows = 11
        .Cols = 6
        .Range(1, 1, .Rows - 1, .Cols - 1).ClearText
        .BackColor2 = vbCyan
        For I = 1 To .Rows - 1
            .Cell(I, 0).Text = I
        Next
        For I = 1 To .Cols - 1
            .Cell(0, I).Text = Chr(64 + I)
        Next
    End With
    
End Sub

Private Sub Command8_Click()
    Call gsGridToExcel(Grid1)
End Sub




Private Sub timeProgress_Timer()
    mstbPaneBar.Value = mlngC
    mstbPaneText.Text = CInt((mlngC / mlngMax) * 100) & "%"
    If mstbPaneBar.Value >= mstbPaneBar.Max Then
        timeProgress.Enabled = False
        Me.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Resize()
    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 12000, 9000)
End Sub
Private Sub Hsb_Change()
    ctlMove.Left = -Hsb.Value
End Sub

Private Sub Hsb_Scroll()
    Call Hsb_Change    '当滑动滚动条中的滑块时会同时更新对应内容，以下同
End Sub

Private Sub Vsb_Change()
    ctlMove.Top = -Vsb.Value
End Sub

Private Sub Vsb_Scroll()
    Call Vsb_Change
End Sub

