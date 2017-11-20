VERSION 5.00
Begin VB.Form frmForm4 
   Caption         =   "测试窗口4"
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

Private Const iAscAdd As Integer = 3    '字符转化中ASCII码值增量



Private Sub Command1_Click()
    '加密输入，生成密文
    
    Dim strText As String
    
    strText = Text1.Item(0).Text
    If Len(strText) > 20 Then
        MsgBox "输入长度不能超过20个字符，超出部分已被删除！", vbExclamation, "长度警告"
        strText = Left(strText, 20)
        Text1.Item(0).Text = strText
    End If
    Text1.Item(1).Text = gEncryptSimple(strText)
    
End Sub

Private Sub Command2_Click()
    '解密密文，还原成明文

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
        .Item(0).Caption = "输入密码："
        .Item(1).Caption = "密文:"
        .Item(2).Caption = "明文:"
    End With
    
    Command1.Caption = "加密"
    Command2.Caption = "解密"
    Command3.Caption = "退出"
    
End Sub


Public Function gBackRandomString() As String
    'EDIT BY XMH,At 2016.10.11
    '返回一个随机字母(或大写或小写)或数字
    '备注：数字与字母的ASCII码范围是（0-9为48-57）、（A-Z为65-90）、（a-z为97-122）
    
    Dim intRandom As Integer    '获取的随机数字
    
    Randomize   '初始化随机数生成器
    Do
        intRandom = CInt(Rnd * 74) + 48     '生成一个在48至122之间的随机数ASCII值
        Select Case intRandom
            Case 48 To 57, 65 To 90, 97 To 122  '如果生成的随机数即ASCII值正好是数字或字母则结束循环，否则重新生成。
                Exit Do
        End Select
    Loop
    
    gBackRandomString = Chr(intRandom)  '返回随机ASCII值对应的数字或字母

End Function

Public Function gAsciiAdd(ByVal strIn) As String
    'EDIT BY XMH ,AT 2016.10.13。
    '返回传入字符的Ascii码值加N后 对应的字符。
    '与gAsciiSub过程互逆
    '注意1：暂时设定支持字母和数字。
    '注意2：输入的字符对应的ASCII值不能超过122即小写字母z。
    '注意3：字符增量大于0且不能超过5。
    
    Dim intASC As Integer
    
    If Len(strIn) = 0 Then Exit Function
    If iAscAdd > 5 Or iAscAdd = 0 Then
        MsgBox "字符增量大于0且不能超过5！", vbExclamation, "字符转化增量警告"
        Exit Function
    End If
    
    intASC = Asc(strIn)
    Select Case intASC
        Case 48 To 57, 65 To 90, 97 To 122
            intASC = intASC + iAscAdd
            Select Case intASC
                Case 48 To 57, 65 To 90, 97 To 122
                    '在些区间表示正常转化
                Case 58 To 64
                    intASC = intASC + 7     '7= - 57 + 64
                Case 91 To 96
                    intASC = intASC + 6     '6= - 90 + 96
                Case 123 To 127
                    intASC = intASC - 75    '-75= - 122 + 47
            End Select
            gAsciiAdd = Chr(intASC)
            
        Case Else
            MsgBox "非法字符转化【" & strIn & "】！" & vbCrLf & "暂不支持数字和字母以外的字符！", vbExclamation, "不支持字符警告"
    End Select
    
End Function

Public Function gAsciiSub(ByVal strIn As String) As String
    'Edit By XMH ,At 2016.10.13。
    '返回传入字符的Ascii码值减N后 对应的字符。
    '与gAsciiAdd过程互逆
    '注意1：暂时设定只支持字母和数字。
    '注意2：输入的字符对应的ASCII值不能超过127。
    '注意3：字符增量大于0且不能超过5。
    
    Dim intSub As Integer
    
    If Len(strIn) = 0 Then Exit Function
    If iAscAdd > 5 Or iAscAdd = 0 Then
        MsgBox "字符增量大于0且不能超过5！", vbExclamation, "字符转化增量警告"
        Exit Function
    End If
    
    intSub = Asc(strIn)
    Select Case intSub
        Case 48 To 57, 65 To 90, 97 To 122
            intSub = intSub - iAscAdd
            Select Case intSub
                Case 48 To 57, 65 To 90, 97 To 122
                    '在些区间表示正常转化
                Case 43 To 47
                    intSub = intSub + 75    '=122-(47-intSub)
                Case 58 To 64
                    intSub = intSub - 7     '=57-(64-intSub)
                Case 91 To 96
                    intSub = intSub - 6     '=90-(96-intSub)
            End Select
            gAsciiSub = Chr(intSub)
            
        Case Else
            MsgBox "非法字符转化【" & strIn & "】！" & vbCrLf & "暂不支持数字和字母以外的字符！", vbExclamation, "不支持字符警告"
    End Select
    
End Function


Public Function gEncryptSimple(ByVal strIn As String) As String
    'EDIT BY XMH,At 2016.10.11
    '将传入的字符串进行简单加密，生成密文并返回给调用者
    
    Dim strEt As String     '密文
    Dim strMid As String    '截取输入字符串中的每一个字符
    Dim K As Long, C As Long        '循环变量
    Dim intAddLenEnd As Integer     '加在最后的字符数量
    Const intAddLenStart As Integer = 5     '加在开始的字符数量
    Const intSumLen As Integer = 30         '密文的总字符数
    
    C = Len(strIn)
    If C = 0 Then Exit Function
    
    '一、将字符串中的每个字符的ASCII值前进N位得到一新字符串，N值为窗体中定义的常量iAscAdd。
    For K = 1 To C
        strEt = strEt & gAsciiAdd(Mid(strIn, K, 1))
    Next
    
    '二、在转化后的字符串strEt前面加入五个随机字符
    For K = 1 To intAddLenStart
        strEt = gBackRandomString & strEt
    Next
    
    '三、在二之前加入三个字符，第一个与第三个字符为数字表示intAddLenEnd值，第二个为随机字符
    intAddLenEnd = intSumLen - intAddLenStart - C - 3
    If intAddLenEnd = 0 Then
        strEt = "0" & gBackRandomString & "0" & strEt
    ElseIf intAddLenEnd < 10 Then
        strEt = "0" & gBackRandomString & CStr(intAddLenEnd) & strEt
    Else
        strEt = Mid(CStr(intAddLenEnd), 1, 1) & gBackRandomString & Mid(CStr(intAddLenEnd), 2, 1) & strEt
    End If
    
    '四、根据三中计算出的intAddLenEnd值，在strEt后追加intAddLenEnd个随机字符，生成最终的密文
    If intAddLenEnd > 0 Then
        For K = 1 To intAddLenEnd
            strEt = strEt & gBackRandomString
        Next
    End If
    
    gEncryptSimple = strEt  '最后将strEt赋给函数的返回值
    
End Function

Public Function gDecryptSimple(ByVal strIn As String) As String
    'Edit By XMH ,At 2016.10.13。
    '解密输入的字符串为明文
    
    Dim strVar As String    '中间流量
    Dim strPt As String     '明文
    Dim strMid As String    '截取输入字符串中的每一个字符
    Dim K As Long, C As Long        '循环变量
    Dim intAddLenEnd As Integer     '加在最后的字符数量
    Const intAddLenStart As Integer = 5     '加在开始的字符数量
    Const intSumLen As Integer = 30         '密文的总字符数
    
    C = Len(strIn)
    If C < (intAddLenStart + 3) Then GoTo LineBreak
    
    '一、获取加在最后的字符个数
    strVar = Left(strIn, 1)
    K = Asc(strVar)
    If K < 48 Or K > 57 Then GoTo LineBreak
    strMid = Mid(strIn, 3, 1)
    K = Asc(strMid)
    If K < 48 Or K > 57 Then GoTo LineBreak
    intAddLenEnd = CInt(strVar & strMid)
    
    '二、删除加在密文最后的intAddLenEnd个字符
    strVar = Left(strIn, 30 - intAddLenEnd)
    
    '三、去除加在前面的八个字符
    strVar = Mid(strVar, 9, C)
    
    '四、解密剩下的strVar字符
    C = Len(strVar)
    If C > 0 Then
        For K = 1 To C
            strPt = strPt & gAsciiSub(Mid(strVar, K, 1))
        Next
    End If
    
    gDecryptSimple = strPt  '将解密好的密文返回给函数的调用者
    
    Exit Function
    
LineBreak:
    MsgBox "传入的不是标准密文，无法解密！", vbExclamation, "密文警告"
    
End Function

