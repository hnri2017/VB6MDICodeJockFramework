VERSION 5.00
Begin VB.Form frmSysLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÏµÍ³µÇÂ½"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4005
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "µÇÂ½"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÃÜÂë"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ÕËºÅ"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   360
   End
End
Attribute VB_Name = "frmSysLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim frmNew As Form
    Dim I As Long
    For I = 1 To 15
        Set frmNew = New frmSysTest
        frmNew.Caption = "Form" & I
        frmNew.Command1.Caption = frmNew.Caption & "cmd1"
        frmNew.Show
    Next
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = gMDI.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not gMDI.Visible Then
        Unload gMDI
    End If
End Sub
