VERSION 5.00
Begin VB.Form frmSysTest 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9405
   Begin VB.HScrollBar Hsb 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   6240
      Width           =   9375
   End
   Begin VB.VScrollBar Vsb 
      Height          =   6255
      Left            =   9000
      TabIndex        =   15
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame ctlMove 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   0
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   120
         Width           =   2175
      End
      Begin VB.ListBox List1 
         Height          =   1140
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   2760
         TabIndex        =   10
         Top             =   0
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   2520
         TabIndex        =   8
         Top             =   720
         Width           =   2175
         Begin VB.TextBox Text1 
            Height          =   735
            Left            =   240
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   180
         Left            =   5280
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   5520
         TabIndex        =   3
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Timer timeProgress 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1320
         Top             =   2160
      End
      Begin 工程1.ucTextComboBox ucTextComboBox1 
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   2880
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "ucTextComboBox1"
      End
      Begin 工程1.ucLabelComboBox ucLabelComboBox1 
         Height          =   300
         Left            =   0
         TabIndex        =   2
         Top             =   3360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         Caption         =   "ucLabelComboBox1"
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
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   5400
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmSysTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstbPaneBar As XtremeCommandBars.StatusBarProgressPane
Dim mstbPaneText As XtremeCommandBars.StatusBarPane
Dim mlngMax As Long, mlngC As Long

Private Type mtypTest
    Result As Boolean
    ErrNO As Long
End Type

Private Sub Command1_Click()
    
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

Private Sub Command2_Click()
    Dim strPath As String
    Dim blnBack As Boolean
    Dim testVar As gtypValueAndErr
    
'    strPath = "123456789"
'    strPath = App.Path & "\data\LogError.txt"
'    strPath = "\bi\flex.ocx"
'    strPath = "\\192.168.2.5\data"
'    strPath = "..\heart.ico"
    strPath = App.Path & "\a\b\"
'    blnBack = gfFileExist(strPath)
    
'    testVar = gfFileExistEx(strPath)
    
'    MsgBox blnBack & "," & testVar.Result & "," & testVar.ErrNum
'    MsgBox Left(strPath, InStrRev(strPath, 4) - 1)
    If gfFileRepair(strPath, True) Then
'    If gfFileRepair(strPath) Then
        MsgBox strPath & vbCrLf & "is created success"
    End If
    
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized

End Sub

Private Sub timeProgress_Timer()
    mstbPaneBar.Value = mlngC
    mstbPaneText.Text = CInt((mlngC / mlngMax) * 100) & "%"
    If mstbPaneBar.Value >= mstbPaneBar.Max Then
        timeProgress.Enabled = False
        Me.Enabled = True
    End If
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
