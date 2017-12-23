VERSION 5.00
Begin VB.UserControl ProgressLabel 
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   ScaleHeight     =   1605
   ScaleWidth      =   4530
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   180
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Label2"
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "ProgressLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'
'
'属性定义


'控件自身事件
Private Sub Label1_Click()
'
End Sub

Private Sub Label2_Click()
'
End Sub

Private Sub Label3_Click()
'
End Sub

'UserControl事件
Private Sub UserControl_Initialize()
    Label1.Caption = ""
    Label2.Caption = ""
    Label3.Caption = "100%"
End Sub

Private Sub UserControl_Resize()
    Label1.Move 0, 0, UserControl.Width, UserControl.Height
    Label2.Move 0, 0, UserControl.Width, UserControl.Height
    With Label3
        .AutoSize = True
        .Move 0, (UserControl.Height - .Height) / 2, UserControl.Width
    End With
End Sub
