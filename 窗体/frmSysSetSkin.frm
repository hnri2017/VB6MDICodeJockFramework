VERSION 5.00
Begin VB.Form frmSysSetSkin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "窗体主题设置"
   ClientHeight    =   3570
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "默认主题"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   160
      Width           =   3735
   End
   Begin VB.ListBox List2 
      Height          =   1140
      Left            =   2880
      TabIndex        =   2
      Top             =   960
      Width           =   2400
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2400
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "退出"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "应用"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主题文件路径："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第二选择："
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主题选择："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frmSysSetSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
    If MsgBox("确定要将当前窗口主题恢复成初始值吗？", vbQuestion + vbOKCancel, "更改确认") = vbCancel Then
        Exit Sub
    End If
    
    List1.ListIndex = -1
    List2.ListIndex = -1
    gID.SkinPath = ""
    gID.SkinIni = ""
    Call gMDI.gmsThemeSkinSet(gID.SkinPath, gID.SkinIni)
    
End Sub

Private Sub Form_Load()
    
    Me.Icon = gMDI.imgListCommandBars.ListImages("themeSet").Picture
    Text1.Text = gID.FolderStyles  '显示主题文件所在路径
    
    Dim strFile As String
    '加载相应文件夹中的主题文件名称
'    '方式一
'    strFile = Dir(gID.FolderStyles & "*styles", vbHidden + vbNormal + vbReadOnly + vbSystem)
'    While Len(strFile) > 0
'        List1.AddItem strFile
'        strFile = Dir
'    Wend

    '方式二。好处是控件在枚举样式文件时会排除掉所有非样式文件
    Dim dES As SkinDescriptions
    Set dES = gMDI.skinFW.EnumerateSkinDirectory(gID.FolderStyles, False)
    If dES Is Nothing Then
        Exit Sub
    End If
    Dim skinDes As SkinDescription
    For Each skinDes In dES
        strFile = skinDes.Path
        List1.AddItem Right(strFile, Len(strFile) - InStrRev(strFile, "\"))
    Next
    
    '定位当前窗口已设置的主题
    If List1.ListCount > 0 Then
        If Len(gID.SkinPath) > 0 Then
            Dim I As Long, J As Long
            
            For I = 0 To List1.ListCount - 1
                
                If LCase(List1.List(I)) = LCase(gID.SkinPath) Then
                    
                    List1.ListIndex = I
                    
                    If Len(gID.SkinIni) > 0 Then
                        For J = 0 To List2.ListCount - 1
                            If LCase(List2.List(J)) = LCase(gID.SkinIni) Then
                                List2.ListIndex = J
                                Exit For
                            End If
                        Next
                    Else
                        List2.ListIndex = -1
                    End If
                    
                    Exit For
                End If
                
            Next
            
        Else
            List1.ListIndex = -1
        End If
    End If
    
End Sub

Private Sub List1_Click()
    '主题选择后自动加载对应的Ini文件列表

    If List1.ListCount = 0 Then Exit Sub

    If List1.ListIndex > -1 Then
        Dim strFile As String
        
        strFile = List1.Text
        List2.Clear
        List2.Text = ""

        If Len(strFile) > 0 Then
            Dim skinDes As SkinDescription
            Dim SkinIni As SkinIniFile
            
            Set skinDes = gMDI.skinFW.EnumerateSkinFile(gID.FolderStyles & strFile)
            
            If skinDes Is Nothing Then  '如果文件出错会是Nothing
                List1.RemoveItem List1.ListIndex
                Exit Sub
            End If
            
            If skinDes.Count > 0 Then
                For Each SkinIni In skinDes
                    List2.AddItem SkinIni.IniFileName
                Next
                
                List2.ListIndex = 0
    
            End If
        End If
    End If
    
End Sub

Private Sub OKButton_Click()
    
    Dim strPro As String
    
    strPro = Replace(Label1.Item(0).Caption, "：", "")
    If List1.ListIndex < 0 Then
        MsgBox strPro & " 中先选一个！", vbExclamation, strPro & "未选提示"
        Exit Sub
    End If
    
    strPro = Replace(Label1.Item(1).Caption, "：", "")
    If List2.ListIndex < 0 Then
        MsgBox strPro & " 中先选一个！", vbExclamation, strPro & "未选提示"
        Exit Sub
    End If
    
    If MsgBox("确定要更改当前窗口主题吗？", vbQuestion + vbOKCancel, "更改确认") = vbCancel Then
        Exit Sub
    End If
    
    gID.SkinPath = List1.Text
    gID.SkinIni = List2.Text
    Call gMDI.gmsThemeSkinSet(gID.SkinPath, gID.SkinIni)
    
End Sub
