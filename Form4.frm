VERSION 5.00
Begin VB.Form frmForm4 
   Caption         =   "���Դ���4"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.TextBox Text3 
      Height          =   390
      Left            =   3120
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    '�������룬��������
    
    Dim strText As String

    strText = Text1.Text
    If Len(strText) > 20 Then
        MsgBox "���볤�Ȳ��ܳ���20���ַ������������ѱ�ɾ����", vbExclamation, "���Ⱦ���"
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
    '�������ģ���ԭ������

    Text3.Text = gfDecryptSimple(Text2.Text)
        
'    Dim strB As String
'    strB = Text2.Text
'    Text1.Text = gfAsciiSub(strB)
    
End Sub


