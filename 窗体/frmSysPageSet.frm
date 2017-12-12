VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSysPageSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "页面设置"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6720
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1815
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "应用"
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6800
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "颜色"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "边框"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "其它"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSysPageSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub msApplyPageSet()
    '保存页面设置结果
    
    Dim gridSet As FlexCell.Grid '用 FlexCell.Grid 方便编程，编完后改为 Control
    
    Set gridSet = frmSysMDI.ActiveForm.ActiveControl
    If (Not gridSet Is Nothing) And (TypeOf gridSet Is FlexCell.Grid) Then
        With gridSet.PageSetup
            
        End With
    Else
        MsgBox "进行【" & Me.Caption & "】的窗口检测异常，请退出该窗口后重试！", vbExclamation
    End If
    
End Sub

Private Sub msLoadPageSet()
    '加载页面设置
    
    
    Dim gridSet As FlexCell.Grid '用 FlexCell.Grid 方便编程，编完后改为 Control
    
    Set gridSet = frmSysMDI.ActiveForm.ActiveControl
    If (Not gridSet Is Nothing) And (TypeOf gridSet Is FlexCell.Grid) Then
        With gridSet.PageSetup
            
        End With
    Else
        MsgBox "进行【" & Me.Caption & "】的窗口检测异常，请重试！", vbExclamation
        Unload Me
    End If
    
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
        Case 0
            Call msApplyPageSet
        Case 1
            Unload Me
        Case 2
            Call msApplyPageSet
            Unload Me
        Case Else
            MsgBox "【" & Command1.Item(Index).Caption & "】按钮 未定义！", vbExclamation
    End Select
End Sub

Private Sub Form_Load()
    
    Call msLoadPageSet
    
End Sub

Private Sub TabStrip1_Click()
    
End Sub
