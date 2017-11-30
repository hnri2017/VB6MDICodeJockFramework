Attribute VB_Name = "modFunction"
Option Explicit


Public Function gfAsciiAdd(ByVal strIn As String) As String
    '���ش����ַ���Ascii��ֵ��N�� ��Ӧ���ַ���
    '��gAsciiSub���̻���
    'ע��1����ʱ�趨֧����ĸ�����֡�
    'ע��2��������ַ���Ӧ��ASCIIֵ���ܳ���122��Сд��ĸz��
    'ע��3���ַ�����Nֵ����0�Ҳ��ܳ���5��
    
    Dim intASC As Integer
    
    If Len(strIn) = 0 Then Exit Function
    
    If gconAscAdd > 5 Or gconAscAdd = 0 Then
        MsgBox "�ַ���������0�Ҳ��ܳ���5��", vbExclamation, "�ַ�ת����������"
        Exit Function
    End If
    
    intASC = Asc(Left(strIn, 1))
    Select Case intASC
        Case 48 To 57, 65 To 90, 97 To 122
            
            intASC = intASC + gconAscAdd
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
            gfAsciiAdd = Chr(intASC)
            
        Case Else
            MsgBox "�Ƿ��ַ�ת����" & strIn & "����" & vbCrLf & "�ݲ�֧�����ֺ���ĸ������ַ���", vbExclamation, "��֧���ַ�����"
    End Select
    
End Function

Public Function gfAsciiSub(ByVal strIn As String) As String
    '���ش����ַ���Ascii��ֵ��N�� ��Ӧ���ַ���
    '��gAsciiAdd���̻���
    'ע��1����ʱ�趨ֻ֧����ĸ�����֡�
    'ע��2��������ַ���Ӧ��ASCIIֵ���ܳ���127��
    'ע��3���ַ�����N����0�Ҳ��ܳ���5��
    
    Dim intSub As Integer
    
    If Len(strIn) = 0 Then Exit Function
    
    If gconAscAdd > 5 Or gconAscAdd = 0 Then
        MsgBox "�ַ���������0�Ҳ��ܳ���5��", vbExclamation, "�ַ�ת����������"
        Exit Function
    End If
    
    intSub = Asc(Left(strIn, 1))
    Select Case intSub
        Case 48 To 57, 65 To 90, 97 To 122
            
            intSub = intSub - gconAscAdd
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
            gfAsciiSub = Chr(intSub)
            
        Case Else
            MsgBox "�Ƿ��ַ�ת����" & strIn & "����" & vbCrLf & "�ݲ�֧�����ֺ���ĸ������ַ���", vbExclamation, "��֧���ַ�����"
    End Select
    
End Function


Public Function gfBackConnection(ByVal strCon As String, _
        Optional ByVal CursorLocation As CursorLocationEnum = adUseClient) As ADODB.Connection
    '�������ݿ�����
       
    On Error GoTo LineErr
    
    Set gfBackConnection = New ADODB.Connection
    gfBackConnection.CursorLocation = CursorLocation
    gfBackConnection.ConnectionString = gID.CnString
    gfBackConnection.CommandTimeout = 5
    gfBackConnection.Open
    
    Exit Function
    
LineErr:
    Call gsAlarmAndLog("���ݿ������쳣")
    
End Function


Public Function gfBackRecordset(ByVal cnSQL As String, _
                Optional ByVal cnCursorType As CursorTypeEnum = adOpenStatic, _
                Optional ByVal cnLockType As LockTypeEnum = adLockReadOnly, _
                Optional ByVal CursorLocation As CursorLocationEnum = adUseClient) As ADODB.Recordset
    '����ָ��SQL��ѯ���ļ�¼��
    
    Dim cnBack As ADODB.Connection
    
    On Error GoTo LineErr

    Set gfBackRecordset = New ADODB.Recordset
    Set cnBack = gfBackConnection(gID.CnString, CursorLocation)
    If cnBack.State = adStateClosed Then Exit Function
    gfBackRecordset.CursorLocation = CursorLocation
    gfBackRecordset.Open cnSQL, cnBack, cnCursorType, cnLockType
    
    Exit Function

LineErr:
    Call gsAlarmAndLog("���ؼ�¼���쳣")

End Function


Public Function gfBackLogType(Optional ByVal strType As genmLogType = udSelect) As String
    '������־��������
    Select Case strType
        Case udDelete
            gfBackLogType = "Delete"
        Case udDeleteBatch
            gfBackLogType = "DeleteBatch"
        Case udInsert
            gfBackLogType = "Insert"
        Case udInsertBatch
            gfBackLogType = "InsertBatch"
        Case udSelectBatch
            gfBackLogType = "SelectBatch"
        Case udUpdate
            gfBackLogType = "Update"
        Case udUpdateBatch
            gfBackLogType = "UpdateBatch"
        Case Else
            gfBackLogType = "Select"
    End Select
End Function


Public Function gfBackOneChar(Optional ByVal CharType As genmCharType = udUpperLowerNum) As String
    '�������һ���ַ�����ĸ�����֣�
    '48-57:0-9
    '65-90:A-Z
    '97-122:a-z
    
    Dim intRd  As Integer

    If (CharType > udUpperLowerNum) Or (CharType < udLowerCase) Then CharType = udUpperLowerNum
    
    Randomize
    Do
        intRd = CInt((74 * Rnd) + 48)
        If (CharType Or udNumber) = CharType Then
            If (intRd > 47 And intRd < 58) Then Exit Do
        End If
        If (CharType Or udUpperCase) = CharType Then
            If (intRd > 64 And intRd < 91) Then Exit Do
        End If
        If (CharType Or udLowerCase) = CharType Then
            If (intRd > 96 And intRd < 123) Then Exit Do
        End If
    Loop
    
    gfBackOneChar = Chr(intRd)
    
End Function


Public Function gfDecryptSimple(ByVal strIn As String) As String
    '����������ַ�������Ϊ����
    '���ĳ����޶�ΪgconSumLenλ
    
    Dim strVar As String    '�м����
    Dim strPt As String     '����
    Dim strMid As String    '��ȡ�����ַ����е�ÿһ���ַ�
    Dim intMid As Integer, K As Integer, c As Integer, R As Integer   '����
    
    strIn = Trim(strIn) 'ȥ�ո�
    c = Len(strIn)
    If c <> gconSumLen Then GoTo LineBreak
    
    'һ����ȡ���������������ַ����������ĵĳ���
    R = Val(Mid(strIn, 2, 1))       '��ȡ���ĵĵڶ�λ����ֵ�����ĵ�gconAddLenStart+1λ�������������������
    If R < 1 Then GoTo LineBreak
    
    intMid = Val(Left(strIn, 1))    '��ȡ���ĵĵ�һλ����������ַ�������ֵ�� ��λ�ϵ�����
    c = IIf(intMid < (gconAddLenStart - 2), intMid, gconAddLenStart - 2)  'ͨ����һλ����ֵ����������ֵ��ʮλ�ϵ���������λ��
    K = Val(Mid(strIn, c + 2 + 1, 1))   '��ȡ�����ֵ��ʮλ�ϵ�����
    c = Val(CStr(K) & CStr(intMid))     '�ó������� ����ַ� ����ֵ
    If (c < (gconSumLen - gconMaxPWD)) Or (c > (gconSumLen - 1)) Then GoTo LineBreak
    
    c = gconSumLen - c  '�ó����ĵĳ���
    c = c * 2           '��Ϊ�����в�������ͬ����������ַ�
    
    '����ɾ����������ǰ���gconAddLenStart+ 1 + R ���ַ� �� �������������ַ�
    strVar = Mid(strIn, gconAddLenStart + 1 + R + 1, c)
    If Len(strVar) <> c Then GoTo LineBreak
    
    '��������ʣ�µ�strVar�ַ�
    For K = 1 To c Step 2
        strPt = strPt & gfAsciiSub(Mid(strVar, K, 1))
    Next
    If Len(strPt) <> c / 2 Then GoTo LineBreak
    
    gfDecryptSimple = strPt  '�����ܺõ����ķ��ظ������ĵ�����
    
    Exit Function
    
LineBreak:
    MsgBox "���ı��ƻ����޷����ܣ�", vbExclamation, "���ľ���"
    
End Function

Public Function gfEncryptSimple(ByVal strIn As String) As String
    '��������ַ���(����)���м򵥼��ܣ��������Ĳ����ظ�������
    '���ĳ���<=20���ַ�����ֻ���Ǵ�д��Сд��ĸ�����֣�����ת��ʱ�ᱨ��
    
    Dim strEt As String     '����
    Dim strMid As String    '��ȡ�����ַ����е�ÿһ���ַ�
    Dim strTen As String    '���ĵ�ǰ10���ַ�
    Dim K As Integer, J As Integer, R As Integer  '����
    Dim c As Integer        '���ĵ��ַ�����
    Dim intFill As Integer  '����ַ���
    Dim intRightNum As Integer      'strFill ��λ�ϵ�����
    Dim intAddLenEnd As Integer     '���������ַ�����

    c = Len(Trim(strIn))
    If c = 0 Then
        MsgBox "�����ַ�����Ϊ���ַ����Ҳ����пո�", vbCritical, "���ַ�����"
        Exit Function
    End If
    strIn = Left(strIn, gconMaxPWD) '��ȡǰgconMaxPWD(20)�ַ�
    c = Len(strIn)  '���»�ȡ�ַ���������Ҫ��
    
    'һ�����ַ����е�ÿ���ַ���ASCIIֵǰ��Nλ������һ������ַ��õ�һ���ַ���
    For K = 1 To c
        strEt = strEt & gfAsciiAdd(Mid(strIn, K, 1)) & gfBackOneChar(udUpperLowerNum)
    Next
    If Len(strEt) <> (c * 2) Then
        MsgBox "�����ַ����淶��ֻ�������ֻ���ĸ��", vbCritical, "�ַ�����"
        Exit Function
    End If
    
    '������ת������ַ���strEtǰ�����Ǽ���gconAddLenStart���ַ�
    '   ����gconAddLenStart���ַ��а������ĵĳ�����ϢgconSumLen-C
    '   Ȼ��gconSumLen-C��ֵ�� ��λ��ʮλ����λ��
    '   Ȼ����strTen�ĵڶ�λ����ԭstrTen��Ӧ����������ָ���
    intFill = gconSumLen - c        '����ȥ�����ĸ�����Ҫ�������ַ�����
    intRightNum = intFill Mod 10    '��ȡ��λ�ϵ�����
    strTen = CStr(intRightNum)      '����λ�ϵ����ַ���strTen�ĵ�һλ,Ҳ�����ĵĵ�һλ
    
    '����strTen�ĵ�һλ��ֵ�������������������ֵĸ���
    J = IIf(intRightNum < (gconAddLenStart - 2), intRightNum, gconAddLenStart - 2)
    For K = 1 To J
        strTen = strTen & gfBackOneChar(udNumber)
    Next
    strTen = strTen & CStr(Int(intFill / 10))   '����intFill��ʮλ�ϵ�����
    
    Do
        R = gfBackOneChar(udNumber)     '��ȡһ��1~9�е��������
        If R > 0 Then Exit Do
    Loop
    strTen = Left(strTen, 1) & CStr(R) & Right(strTen, Len(strTen) - 1)
    
    '��strTen�ĳ��Ȳ���gconAddLenStartλ��������������,����strTen���沢�����R������
    J = (gconAddLenStart - 2 - J) + R
    For K = 1 To J
        strTen = strTen & gfBackOneChar(udNumber)
    Next
    strEt = strTen & strEt
    
    '������strEt��׷��intAddLenEnd������ַ��ճ�gconSumLen���ַ�����������
    intAddLenEnd = gconSumLen - (c * 2) - gconAddLenStart - R - 1
    If intAddLenEnd > 0 Then
        For K = 1 To intAddLenEnd
            strEt = strEt & gfBackOneChar(udUpperLowerNum)
        Next
    End If
    
    gfEncryptSimple = strEt  '���strEt���������ķ���ֵ
    
End Function


Public Function gfFileExist(ByVal strPath As String) As Boolean
    '�ж��ļ����ļ�Ŀ¼ �Ƿ����

    Dim strBack As String
        
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '���ַ�������
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then gfFileExist = True
    End If
  
    Exit Function
    
LineErr:
    Call gsAlarmAndLog("�ж��ļ��쳣")
    
End Function


Public Function gfFileExistEx(ByVal strPath As String) As gtypValueAndErr
    '��һ�ַ���ֵ��ʽ�����ж��ļ����ļ�Ŀ¼ �Ƿ����
    'ר������Ĺ���gfFileRepair����
    
    Dim strBack As String
    
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '���ַ�������
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then
            gfFileExistEx.Result = True
        Else
            gfFileExistEx.ErrNum = -1   '�����ڣ�Ҳû�쳣
        End If
    End If
    
    Exit Function
    
LineErr:
    gfFileExistEx.ErrNum = Err.Number   '�쳣�ˣ�Ҳ������������
    Call gsAlarmAndLog("�ļ��жϷ����쳣")
    
End Function


Public Function gfFileOpen(ByVal strFilePath As String) As gtypValueAndErr
    '��ָ��ȫ·�����ļ�
    
    Dim lngRet As Long
    Dim strDir As String
    
    On Error GoTo LineErr
    
    If gfFileExist(strFilePath) Then
        
        lngRet = ShellExecute(GetDesktopWindow, "open", strFilePath, vbNullString, vbNullString, vbNormalFocus)
        If lngRet = SE_ERR_NOASSOC Then     'û�й����ĳ���
             strDir = Space(260)
             lngRet = GetSystemDirectory(strDir, Len(strDir))
             strDir = Left(strDir, lngRet)
             
            '��ʾ�򿪷�ʽ����
            Call ShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & strFilePath, strDir, vbNormalFocus)
            gfFileOpen.ErrNum = -1   '���ɹ���Ҳû�쳣
        Else
            gfFileOpen.Result = True
        End If
        
    End If
    
    Exit Function
    
LineErr:
    gfFileOpen.ErrNum = Err.Number
    Call gsAlarmAndLog("�ļ����쳣")
    
End Function


Public Function gfFileRepair(ByVal strFile As String, Optional ByVal blnFolder As Boolean) As Boolean
    '��� �ļ�/�ļ��� ������ �򴴽�
    'ǰ����·�����ϲ�Ŀ¼�ɷ���
    '����blnFolderָ�������·��strFile���ļ�����ΪTrue��Ĭ�����ļ�False
    
    Dim strTemp As String
    Dim typBack As gtypValueAndErr
    Dim lngLoc As Long
    
    If Right(strFile, 1) = "\" Then
        strFile = Left(strFile, Len(strFile) - 1)   'ȥ����ĩ��"\"
    End If
    strTemp = strFile
    If Len(strTemp) = 0 Then Exit Function          '��ֹ������ַ���
    
    On Error GoTo LineErr

    typBack = gfFileExistEx(strTemp)    '�ж��Ƿ����
    If Not typBack.Result Then          '�ļ�������
        If typBack.ErrNum = -1 Then     '�����쳣
            
            lngLoc = InStrRev(strTemp, "\") '�ж��Ƿ����ϲ�Ŀ¼
            If lngLoc > 0 Then              '���ϲ�Ŀ¼��ݹ�
                strTemp = Left(strTemp, lngLoc - 1) '�ó��ϲ�Ŀ¼�ľ���·��
                Call gfFileRepair(strTemp, True)    '�ݹ���������Ա�֤�ϲ�Ŀ¼����
            End If

            If blnFolder Then                   '����������ļ���
                MkDir strFile                   '�򴴽��ļ���
            Else                                '����������ļ�
                Close                           '�򴴽��ļ�
                Open strFile For Random As #1
                Close
            End If
            
            gfFileRepair = True '�����ɹ�����True
            
        End If
        
    Else
        gfFileRepair = True '·������ֱ�ӷ���True
    End If

LineErr:
    
End Function


Public Function gfFormLoad(ByVal strFormName As String) As Boolean
    '�ж�ָ�������Ƿ񱻼�����
    
    Dim frmLoad As Form
    
    strFormName = LCase(strFormName)
    For Each frmLoad In Forms
        If LCase(frmLoad.Name) = strFormName Then
            gfFormLoad = True
            Exit Function
        End If
    Next
    
End Function


Public Function gfStringCheck(ByVal strIn As String) As String
    '''�����ַ����
    
    Dim arrStr As Variant
    Dim I As Long
    
    arrStr = Array(";", "--", "'", "//", "/*", "*/", "select", "update", _
                   "delete", "insert", "alter", "drop", "create")
    strIn = LCase(strIn)
    For I = LBound(arrStr) To UBound(arrStr)
        If InStr(strIn, arrStr(I)) > 0 Then
            gfStringCheck = arrStr(I)
            Exit Function
        End If
    Next

End Function
