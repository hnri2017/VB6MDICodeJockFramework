VERSION 5.00
Begin VB.Form frmSysTestEx 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   9000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1080
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
    
    strCur = ""
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
