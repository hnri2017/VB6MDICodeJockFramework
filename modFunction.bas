Attribute VB_Name = "modFunction"
Option Explicit



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
    MsgBox "�쳣���ţ�" & Err.Number & vbCrLf & "�쳣������" & Err.Description, vbCritical
    
End Function

Public Function gfFileExistEx(ByVal strPath As String) As gtypValueAndErr
    '��һ�ַ���ֵ��ʽ�����ж��ļ����ļ�Ŀ¼ �Ƿ����
    
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

Public Function gfFileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As Long = 0, Optional ByVal WriteMode As Long = 0) As Boolean
    '��ָ��������ָ���ķ�ʽд��ָ���ļ���
    
    
End Function
