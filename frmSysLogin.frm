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


Private Function mfVersionCheck() As Boolean
    '�汾���
    
    Dim fsoVer As FileSystemObject
    Dim arrNet() As String
    Dim arrLoc() As String
    Dim I As Long
    Dim blnNew As Boolean
    Dim strOut As String
    
    If Not gfFileExist(gID.FileAppNet) Then Exit Function   '�����ϵ��ļ��Ƿ����
    
    On Error GoTo LineErr
    
    If GetAttr(gID.FileAppNet) <> vbNormal Then SetAttr gID.FileAppNet, vbNormal    '�޸ĳ���������
    If GetAttr(gID.FileAppLoc) <> vbNormal Then SetAttr gID.FileAppLoc, vbNormal    '
    
    Set fsoVer = New FileSystemObject
    arrNet = Split(fsoVer.GetFileVersion(gID.FileAppNet), ".")
    arrLoc = Split(fsoVer.GetFileVersion(gID.FileAppLoc), ".")
    For I = 0 To UBound(arrNet)
        If Val(arrNet(I)) > Val(arrLoc(I)) Then
            blnNew = True
            Exit For
        End If
    Next
    
    If blnNew Then
        If Not gfFileExist(gID.FileSetupNet) Then
            MsgBox "���³����쳣��", vbCritical
            Exit Function
        End If
        
        If GetAttr(gID.FileSetupNet) <> vbNormal Then SetAttr gID.FileSetupNet, vbNormal
        FileCopy gID.FileSetupNet, gID.FileSetupLoc
        Shell gID.FileSetupLoc
        
        Set gMDI = Nothing
        End
        
    End If
    
    Set fsoVer = Nothing
    mfVersionCheck = True
    
    Exit Function
    
LineErr:
    Call gsAlarmAndLog("�汾����쳣")
    
End Function


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


Private Sub Command1_Click()
    Dim strName As String
    Dim frmNew As Form
    Dim I As Long
    
    strName = Trim(ucTC.Text)
        
    SaveSetting gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserLast, strName
    Call msSaveUserList
    
    gMDI.Show
    
    Unload Me
    
End Sub

Private Sub Form_Load()
   
    If App.PrevInstance Then
        If MsgBox("�ó����ڽ����д��ڡ��Ѿ����򿪣�" & vbCrLf & vbCrLf _
            & "�����鿪�������ͬ����ˣ��Ƿ���Ҫ������", vbExclamation + vbYesNo) = vbNo Then
            Set gMDI = Nothing
            End
        End If
    End If
        
    If Not mfVersionCheck Then
        If MsgBox("����汾���ʧ�ܣ��Ƿ������½��", vbExclamation + vbYesNo) = vbNo Then
            Set gMDI = Nothing
            End
        End If
    End If
    
    Set Me.Icon = gMDI.Icon
    
'''    ucTC.Text = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveUserLast, "")
    Call msLoadUserList
    If ucTC.ListCount > 0 Then ucTC.Text = ucTC.List(0)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not gMDI.Visible Then
        Unload gMDI
    End If
End Sub
