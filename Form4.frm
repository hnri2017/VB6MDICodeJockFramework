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
End
Attribute VB_Name = "frmForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const iAscAdd As Integer = 3    '�ַ�ת����ASCII��ֵ����



Private Sub Command1_Click()
    '�������룬��������
    
    Dim strText As String
    
    strText = Text1.Item(0).Text
    If Len(strText) > 20 Then
        MsgBox "���볤�Ȳ��ܳ���20���ַ������������ѱ�ɾ����", vbExclamation, "���Ⱦ���"
        strText = Left(strText, 20)
        Text1.Item(0).Text = strText
    End If
    Text1.Item(1).Text = gEncryptSimple(strText)
    
End Sub

Private Sub Command2_Click()
    '�������ģ���ԭ������

    Text1.Item(2).Text = gDecryptSimple(Text1.Item(1).Text)
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With Text1
        .Item(0).Text = ""
        .Item(1).Text = ""
        .Item(2).Text = ""
    End With
    
    With Label1
        .Item(0).Caption = "�������룺"
        .Item(1).Caption = "����:"
        .Item(2).Caption = "����:"
    End With
    
    Command1.Caption = "����"
    Command2.Caption = "����"
    Command3.Caption = "�˳�"
    
End Sub


Public Function gBackRandomString() As String
    'EDIT BY XMH,At 2016.10.11
    '����һ�������ĸ(���д��Сд)������
    '��ע����������ĸ��ASCII�뷶Χ�ǣ�0-9Ϊ48-57������A-ZΪ65-90������a-zΪ97-122��
    
    Dim intRandom As Integer    '��ȡ���������
    
    Randomize   '��ʼ�������������
    Do
        intRandom = CInt(Rnd * 74) + 48     '����һ����48��122֮��������ASCIIֵ
        Select Case intRandom
            Case 48 To 57, 65 To 90, 97 To 122  '������ɵ��������ASCIIֵ���������ֻ���ĸ�����ѭ���������������ɡ�
                Exit Do
        End Select
    Loop
    
    gBackRandomString = Chr(intRandom)  '�������ASCIIֵ��Ӧ�����ֻ���ĸ

End Function

Public Function gAsciiAdd(ByVal strIn) As String
    'EDIT BY XMH ,AT 2016.10.13��
    '���ش����ַ���Ascii��ֵ��N�� ��Ӧ���ַ���
    '��gAsciiSub���̻���
    'ע��1����ʱ�趨֧����ĸ�����֡�
    'ע��2��������ַ���Ӧ��ASCIIֵ���ܳ���122��Сд��ĸz��
    'ע��3���ַ���������0�Ҳ��ܳ���5��
    
    Dim intASC As Integer
    
    If Len(strIn) = 0 Then Exit Function
    If iAscAdd > 5 Or iAscAdd = 0 Then
        MsgBox "�ַ���������0�Ҳ��ܳ���5��", vbExclamation, "�ַ�ת����������"
        Exit Function
    End If
    
    intASC = Asc(strIn)
    Select Case intASC
        Case 48 To 57, 65 To 90, 97 To 122
            intASC = intASC + iAscAdd
            Select Case intASC
                Case 48 To 57, 65 To 90, 97 To 122
                    '��Щ�����ʾ����ת��
                Case 58 To 64
                    intASC = intASC + 7     '7= - 57 + 64
                Case 91 To 96
                    intASC = intASC + 6     '6= - 90 + 96
                Case 123 To 127
                    intASC = intASC - 75    '-75= - 122 + 47
            End Select
            gAsciiAdd = Chr(intASC)
            
        Case Else
            MsgBox "�Ƿ��ַ�ת����" & strIn & "����" & vbCrLf & "�ݲ�֧�����ֺ���ĸ������ַ���", vbExclamation, "��֧���ַ�����"
    End Select
    
End Function

Public Function gAsciiSub(ByVal strIn As String) As String
    'Edit By XMH ,At 2016.10.13��
    '���ش����ַ���Ascii��ֵ��N�� ��Ӧ���ַ���
    '��gAsciiAdd���̻���
    'ע��1����ʱ�趨ֻ֧����ĸ�����֡�
    'ע��2��������ַ���Ӧ��ASCIIֵ���ܳ���127��
    'ע��3���ַ���������0�Ҳ��ܳ���5��
    
    Dim intSub As Integer
    
    If Len(strIn) = 0 Then Exit Function
    If iAscAdd > 5 Or iAscAdd = 0 Then
        MsgBox "�ַ���������0�Ҳ��ܳ���5��", vbExclamation, "�ַ�ת����������"
        Exit Function
    End If
    
    intSub = Asc(strIn)
    Select Case intSub
        Case 48 To 57, 65 To 90, 97 To 122
            intSub = intSub - iAscAdd
            Select Case intSub
                Case 48 To 57, 65 To 90, 97 To 122
                    '��Щ�����ʾ����ת��
                Case 43 To 47
                    intSub = intSub + 75    '=122-(47-intSub)
                Case 58 To 64
                    intSub = intSub - 7     '=57-(64-intSub)
                Case 91 To 96
                    intSub = intSub - 6     '=90-(96-intSub)
            End Select
            gAsciiSub = Chr(intSub)
            
        Case Else
            MsgBox "�Ƿ��ַ�ת����" & strIn & "����" & vbCrLf & "�ݲ�֧�����ֺ���ĸ������ַ���", vbExclamation, "��֧���ַ�����"
    End Select
    
End Function


Public Function gEncryptSimple(ByVal strIn As String) As String
    'EDIT BY XMH,At 2016.10.11
    '��������ַ������м򵥼��ܣ��������Ĳ����ظ�������
    
    Dim strEt As String     '����
    Dim strMid As String    '��ȡ�����ַ����е�ÿһ���ַ�
    Dim K As Long, C As Long        'ѭ������
    Dim intAddLenEnd As Integer     '���������ַ�����
    Const intAddLenStart As Integer = 5     '���ڿ�ʼ���ַ�����
    Const intSumLen As Integer = 30         '���ĵ����ַ���
    
    C = Len(strIn)
    If C = 0 Then Exit Function
    
    'һ�����ַ����е�ÿ���ַ���ASCIIֵǰ��Nλ�õ�һ���ַ�����NֵΪ�����ж���ĳ���iAscAdd��
    For K = 1 To C
        strEt = strEt & gAsciiAdd(Mid(strIn, K, 1))
    Next
    
    '������ת������ַ���strEtǰ������������ַ�
    For K = 1 To intAddLenStart
        strEt = gBackRandomString & strEt
    Next
    
    '�����ڶ�֮ǰ���������ַ�����һ����������ַ�Ϊ���ֱ�ʾintAddLenEndֵ���ڶ���Ϊ����ַ�
    intAddLenEnd = intSumLen - intAddLenStart - C - 3
    If intAddLenEnd = 0 Then
        strEt = "0" & gBackRandomString & "0" & strEt
    ElseIf intAddLenEnd < 10 Then
        strEt = "0" & gBackRandomString & CStr(intAddLenEnd) & strEt
    Else
        strEt = Mid(CStr(intAddLenEnd), 1, 1) & gBackRandomString & Mid(CStr(intAddLenEnd), 2, 1) & strEt
    End If
    
    '�ġ��������м������intAddLenEndֵ����strEt��׷��intAddLenEnd������ַ����������յ�����
    If intAddLenEnd > 0 Then
        For K = 1 To intAddLenEnd
            strEt = strEt & gBackRandomString
        Next
    End If
    
    gEncryptSimple = strEt  '���strEt���������ķ���ֵ
    
End Function

Public Function gDecryptSimple(ByVal strIn As String) As String
    'Edit By XMH ,At 2016.10.13��
    '����������ַ���Ϊ����
    
    Dim strVar As String    '�м�����
    Dim strPt As String     '����
    Dim strMid As String    '��ȡ�����ַ����е�ÿһ���ַ�
    Dim K As Long, C As Long        'ѭ������
    Dim intAddLenEnd As Integer     '���������ַ�����
    Const intAddLenStart As Integer = 5     '���ڿ�ʼ���ַ�����
    Const intSumLen As Integer = 30         '���ĵ����ַ���
    
    C = Len(strIn)
    If C < (intAddLenStart + 3) Then GoTo LineBreak
    
    'һ����ȡ���������ַ�����
    strVar = Left(strIn, 1)
    K = Asc(strVar)
    If K < 48 Or K > 57 Then GoTo LineBreak
    strMid = Mid(strIn, 3, 1)
    K = Asc(strMid)
    If K < 48 Or K > 57 Then GoTo LineBreak
    intAddLenEnd = CInt(strVar & strMid)
    
    '����ɾ��������������intAddLenEnd���ַ�
    strVar = Left(strIn, 30 - intAddLenEnd)
    
    '����ȥ������ǰ��İ˸��ַ�
    strVar = Mid(strVar, 9, C)
    
    '�ġ�����ʣ�µ�strVar�ַ�
    C = Len(strVar)
    If C > 0 Then
        For K = 1 To C
            strPt = strPt & gAsciiSub(Mid(strVar, K, 1))
        Next
    End If
    
    gDecryptSimple = strPt  '�����ܺõ����ķ��ظ������ĵ�����
    
    Exit Function
    
LineBreak:
    MsgBox "����Ĳ��Ǳ�׼���ģ��޷����ܣ�", vbExclamation, "���ľ���"
    
End Function

