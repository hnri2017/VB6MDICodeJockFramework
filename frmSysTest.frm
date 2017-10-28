VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
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
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   495
         Left            =   4800
         TabIndex        =   23
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   495
         Left            =   4800
         TabIndex        =   22
         Top             =   4560
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   2775
         Left            =   0
         TabIndex        =   21
         Top             =   2760
         Width           =   4575
         _cx             =   8070
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
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   495
         Left            =   5760
         TabIndex        =   20
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   495
         Left            =   5760
         TabIndex        =   19
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   495
         Left            =   5760
         TabIndex        =   18
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   495
         Left            =   5760
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
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
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   5760
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame1"
         Height          =   1335
         Left            =   2280
         TabIndex        =   8
         Top             =   120
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
         Left            =   5760
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   375
         Left            =   4680
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   180
         Left            =   4680
         TabIndex        =   5
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Timer timeProgress 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3360
         Top             =   1800
      End
      Begin 工程1.ucTextComboBox ucTextComboBox1 
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   1800
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
         Top             =   2280
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
         Left            =   4800
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   4800
         TabIndex        =   13
         Top             =   720
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

Private Sub Command3_Click()
    Dim strVar As String
    
    strVar = "This is a test string!"
'    Open gID.FileLog For Output As #1
    Open gID.FileLog For Random As #1
    Put #1, , "put"
'    Print #1, ""
    Close #1
    
'    Call gfFileWrite(gID.FileLog, "", udOutput)
    Call gfFileWrite(gID.FileLog, strVar)
    Call gfFileWrite(gID.FileLog, strVar, , udWrite)
    Call gfFileWrite(gID.FileLog, strVar & vbTab & "Write")
    Call gfFileWrite(gID.FileLog, strVar & vbTab & "print", , udWrite)
    Call gfFileWrite(gID.FileLog, strVar, udAppend, udPut)
End Sub

Private Sub Command4_Click()
    Dim strFile As String
    
'    strFile = "\\192.168.2.5\data\err.log"
'    Call gfFileExist(strFile)
'    strFile = App.Path & "\data"
'    Call gfFileExist(strFile)
'    strFile = App.Path & "\data\record.log"
'    Call gfFileExist(strFile)
'    strFile = "."
'    Call gfFileExist(strFile)
'    strFile = ".."
'    Call gfFileExist(strFile)
    
    On Error Resume Next
    strFile = 1 / 0
    strFile = Err.Number & vbTab & Err.Description
    Call gfFileWrite(gID.FileLog, strFile, udOutput)
End Sub

Private Sub Command5_Click()
    Dim fsoU As FileSystemObject
    
    
End Sub

Private Sub Command6_Click()
    Dim strSQL As String
    Dim rsT As ADODB.Recordset
    
    strSQL = "SELECT * FROM tb_Test_User"
    Set rsT = gfBackRecordset(strSQL)
    If rsT.State = adStateOpen Then
        MsgBox "Connection Select OK"
        Debug.Print rsT.Fields.Count, rsT.Fields(0).Value, rsT.Fields(1).Value, rsT.Fields(2).Value
    End If
    
    Set rsT = Nothing
    
End Sub

Private Sub Command7_Click()
With VSFlexGrid1
        .Rows = 20
        .Cols = 20
        .ColWidth(-1) = 500
        .AllowUserResizing = flexResizeBoth
        .Cell(flexcpText, 0, 0, .Rows - 1, .Cols - 1) = "abc"
        .Cell(flexcpChecked, 1, 1, .Rows - 1) = True
        .Cell(flexcpTextStyle, 1, 3, .Rows - 1) = 2
        .BackColorAlternate = vbCyan
    End With
    
End Sub

Private Sub Command8_Click()
    Call gsGridToExcel(VSFlexGrid1)
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
