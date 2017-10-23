VERSION 5.00
Begin VB.Form frmSysLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ϵͳ��½"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4005
   StartUpPosition =   2  '��Ļ����
   Begin ����1.ucTextComboBox ucTC 
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   12
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��½"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   580
      Width           =   360
   End
End
Attribute VB_Name = "frmSysLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const mconDot As String = ","


Private Sub msLoadUserList()
    '����������½�����û����б�
    
    Dim strReg As String
    Dim strList() As String
    Dim strName As String
    Dim I As Long
    
    strReg = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserList, "")
    If Len(strReg) > 0 Then
        strList = Split(strReg, mconDot)
        For I = 0 To UBound(strList)
            strName = Trim(strList(I))  '��ֹע�������Ϣ����Ϊ���߼���ո�
            If Len(strName) > 0 Then ucTC.AddItem strName  '����ո�
        Next
    End If
    
End Sub

Private Sub msSaveUserList()
    '����ǰ�û�����������½�����û����б���
    '���������б�ĵ�һλ������ʾԽ����½�����û���Խ�����б�ǰ��
    
    Dim strName As String
    Dim strReg As String
    Dim strSave As String
    Dim strList() As String
    Dim I As Long
    
    strName = Trim(ucTC.Text)
    If Len(strName) = 0 Then Exit Sub
    strReg = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserList, "")
        
    If Len(strReg) > 0 Then
        strList = Split(strReg, mconDot)
        If LCase(strName) = LCase(Trim(strList(0))) Then Exit Sub
        
        For I = 0 To UBound(strList)    '��ѭ������һ��ԭ���û������˵Ŀո񣬲�ȥ���˴ε�½����
            If LCase(strName) = LCase(Trim(strList(I))) Then
                strList(I) = ""
            Else
                strList(I) = Trim(strList(I))
            End If
        Next
        
        strSave = strName & mconDot '��ǰ��½���û�����������ǰ��
        For I = 0 To UBound(strList)    '����Ч���û���ƴ������
            If Len(strList(I)) > 0 Then strSave = strSave & strList(I) & mconDot
        Next
        
        strSave = Left(strSave, Len(strSave) - 1)   'ȥ�����ұߵ�mconDot
    Else
        strSave = strName
    End If
    
    SaveSetting gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserList, strSave
    
End Sub

Private Function mfVersionCheck() As Boolean
    '�汾���

    mfVersionCheck = True
    
End Function


Private Sub Command1_Click()
    Dim strName As String
    Dim frmNew As Form
    Dim I As Long
    
    strName = Trim(ucTC.Text)
    
    For I = 1 To 1
        Set frmNew = New frmSysTest
        frmNew.Caption = "Form" & I
        frmNew.Command1.Caption = frmNew.Caption & "cmd1"
        frmNew.Show
    Next
    
    SaveSetting gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserLast, strName
    Call msSaveUserList
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    If Not mfVersionCheck Then
        Unload Me
    End If
    
    Set Me.Icon = gMDI.Icon
    
    ucTC.Text = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserLast, "")
    Call msLoadUserList
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not gMDI.Visible Then
        Unload gMDI
    End If
End Sub
