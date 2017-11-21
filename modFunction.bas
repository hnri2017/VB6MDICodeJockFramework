Attribute VB_Name = "modFunction"
Option Explicit


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


Public Function gfBackOneChar(Optional ByVal CharType As genmCharType = udUpperLowerNum) As String
    '�������һ���ַ�����ĸ�����֣�
    '48-57:0-9
    '65-90:A-Z
    '97-122:a-z
    
    Dim intRd  As Integer

    If (CharType > udUpperLowerNum) Or (CharType < udLowerCase) Then CharType = udUpperLowerNum
    
    Randomize
    Do
        intRd = CInt((72 * Rnd) + 48)
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
