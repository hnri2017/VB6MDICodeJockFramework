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
    If Not typBack.Result Then
        If typBack.ErrNum = -1 Then     '�����ڣ�Ҳû�쳣
            lngLoc = InStrRev(strTemp, "\")
            If lngLoc > 0 Then  '�ж��Ƿ����ϲ�Ŀ¼������ݹ�
                If gfFileExistEx(strTemp).Result Then
                    If blnFolder Then
                        MkDir strTemp                   '�����ļ���
                    Else
                        Close
                        Open strTemp For Random As #1   '�����ļ�
                        Close
                    End If
                Else
                    strTemp = Left(strTemp, lngLoc - 1)
                    If Not gfFileRepair(strTemp, True) Then Exit Function
                End If
            End If
            
            '���ϲ�Ŀ¼ֱ�Ӵ���
            If blnFolder Then
                MkDir strFile                   '�����ļ���
            Else
                Close
                Open strFile For Random As #1   '�����ļ�
            End If
            
            gfFileRepair = True '����ִ�гɹ�
            Close

        End If
    End If

LineErr:
    
End Function

Public Function gfFileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As Long = 0, Optional ByVal WriteMode As Long = 0) As Boolean
    '��ָ��������ָ���ķ�ʽд��ָ���ļ���
    
    
End Function
