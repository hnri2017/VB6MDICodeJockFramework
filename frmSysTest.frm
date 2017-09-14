VERSION 5.00
Begin VB.Form frmSysTest 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   9285
   Begin VB.Timer timeProgress 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   4320
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   255
      Left            =   5640
      TabIndex        =   15
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   180
      Left            =   5640
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   2880
      TabIndex        =   7
      Top             =   1320
      Width           =   2295
      Begin VB.TextBox Text1 
         Height          =   3495
         Left            =   360
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5295
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   5880
      Width           =   7815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   5760
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5760
      TabIndex        =   9
      Top             =   600
      Width           =   1455
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
