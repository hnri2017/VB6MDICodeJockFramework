VERSION 5.00
Begin VB.Form frmSysTestEx 
   Caption         =   "测试窗口二"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   9000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Form4"
      Height          =   615
      Left            =   4800
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Form3"
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin 工程1.ucLabelComboBox ucLabelComboBox1 
      Height          =   405
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   714
      Caption         =   "ucLabelComboBox1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin 工程1.ucTextComboBox ucTextComboBox1 
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "幼圆"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
      Text            =   "ucTextComboBox1"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "frmSysTestEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim rsCur As ADODB.Recordset
    Dim strCur As String
    
    strCur = "tb_test_user"
    Set rsCur = gfBackRecordset(strCur)
    
    If rsCur Is Nothing Then
        MsgBox "rsCur Is Nothing"
    Else
        MsgBox "rsCur Is Not Nothing"
    End If
    
    If rsCur.State = adStateOpen Then
        MsgBox "rsCur.State is adStateOpen"
    Else
        MsgBox "rsCur.State is adStateClosed"
    End If
    
    Set rsCur = Nothing
    
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form4.Show
End Sub

Private Sub Form_Load()
    With Me.ucTextComboBox1
        .AddItem "tb111"
        .AddItem "tb222"
        .AddItem "tbAAA"
    End With
    
    With Me.ucLabelComboBox1
        .AddItem "lcQQQ"
        .AddItem "lcWWW"
        .AddItem "lc3333"
    End With
End Sub

