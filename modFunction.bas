Attribute VB_Name = "modFunction"
Option Explicit



Public Function gfFileExist(ByVal strPath As String) As Boolean
    '判断文件、文件目录 是否存在
    
    Dim strBack As String
    
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '空字符串不算
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then gfFileExist = True
    End If
    
    Exit Function
    
LineErr:
    MsgBox "异常代号：" & Err.Number & vbCrLf & "异常描述：" & Err.Description, vbCritical
    
End Function

Public Function gfFileExistEx(ByVal strPath As String) As gtypValueAndErr
    '另一种返回值方式：来判断文件、文件目录 是否存在
    
    Dim strBack As String
    
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '空字符串不算
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then
            gfFileExistEx.Result = True
        Else
            gfFileExistEx.ErrNum = -1   '不存在，也没异常
        End If
    End If
    
    Exit Function
    
LineErr:
    gfFileExistEx.ErrNum = Err.Number   '异常了，也当作不存在了
    
End Function

Public Function gfFileRepair(ByVal strFile As String, Optional ByVal blnFolder As Boolean) As Boolean
    '如果 文件/文件夹 不存在 则创建
    '前提是路径的上层目录可访问
    '参数blnFolder指明传入的路径strFile是文件夹则为True，默认是文件False
    
    Dim strTemp As String
    Dim typBack As gtypValueAndErr
    Dim lngLoc As Long
    
    If Right(strFile, 1) = "\" Then
        strFile = Left(strFile, Len(strFile) - 1)   '去掉最末的"\"
    End If
    strTemp = strFile
    If Len(strTemp) = 0 Then Exit Function          '防止传入空字符串
    
    On Error GoTo LineErr

    typBack = gfFileExistEx(strTemp)    '判断是否存在
    If Not typBack.Result Then          '文件不存在
        If typBack.ErrNum = -1 Then     '且无异常
            
            lngLoc = InStrRev(strTemp, "\") '判断是否有上层目录
            If lngLoc > 0 Then              '有上层目录则递归
                strTemp = Left(strTemp, lngLoc - 1) '得出上层目录的具体路径
                Call gfFileRepair(strTemp, True)    '递归调用自身，以保证上层目录存在
            End If

            If blnFolder Then                   '传入参数是文件夹
                MkDir strFile                   '则创建文件夹
            Else                                '传入参数是文件
                Close                           '则创建文件
                Open strFile For Random As #1
                Close
            End If
            
            gfFileRepair = True '创建成功返回True
            
        End If
        
    Else
        gfFileRepair = True '路径完整直接返回True
    End If

LineErr:
    
End Function

Public Function gfFileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As Long = 0, Optional ByVal WriteMode As Long = 0) As Boolean
    '将指定内容以指定的方式写入指定文件中
    
    
End Function
