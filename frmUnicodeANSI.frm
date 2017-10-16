VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   9165
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   2400
      TabIndex        =   5
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "×¢²á¿Ø¼þ"
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UNICODEtoANSI"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ANSItoUnicode"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim bytS() As Byte
    Dim strS As String
    Dim I As Long
    
    strS = Text1.Text
    bytS = StrConv(strS, vbUnicode)
    
    strS = ""
    For I = 0 To UBound(bytS)
        strS = strS & bytS(I) & " "
    Next
    strS = Left(strS, Len(strS) - 1)
    Text2.Text = strS
    
End Sub

Private Sub Command2_Click()
    
    Dim bytS() As Byte
    Dim strS As String
    Dim I As Long
    Dim arrS() As String
    
    strS = Text2.Text
    arrS = Split(strS, " ")
    ReDim bytS(UBound(arrS))
    For I = 0 To UBound(arrS)
        bytS(I) = arrS(I)
    Next
    strS = StrConv(bytS, vbFromUnicode)
    Text1.Text = strS
    
End Sub

Private Sub Command3_Click()
    Dim strPath As String
    Dim dblValue As Double
    
    If List1.ListCount = 0 Then Exit Sub
    If List1.ListIndex = -1 Then Exit Sub
    
    strPath = App.Path & "\bin\" & List1.Text
    dblValue = Shell("c:\windows\system32\regsvr32 " & strPath)
    MsgBox dblValue
End Sub

Private Sub Form_Load()
    Dim strPath As String
    Dim intLoc As Integer
    
    List1.Clear
    strPath = Dir(App.Path & "\bin\*.ocx")
    While Len(strPath) > 0
        intLoc = InStrRev(strPath, "\")
        List1.AddItem Right(strPath, Len(strPath) - intLoc)
        strPath = Dir
    Wend
    
End Sub
