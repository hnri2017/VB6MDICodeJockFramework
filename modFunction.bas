Attribute VB_Name = "modFunction"
Option Explicit


Public Function gfBackConnection(ByVal strCon As String, _
        Optional ByVal CursorLocation As CursorLocationEnum = adUseClient) As ADODB.Connection
    '返回数据库连接
       
    On Error GoTo LineErr
    
    Set gfBackConnection = New ADODB.Connection
    gfBackConnection.CursorLocation = CursorLocation
    gfBackConnection.ConnectionString = gID.CnString
    gfBackConnection.CommandTimeout = 5
    gfBackConnection.Open
    
    Exit Function
    
LineErr:
    Call gsAlarmAndLog("数据库连接异常")
    
End Function


Public Function gfBackRecordset(ByVal cnSQL As String, _
                Optional ByVal cnCursorType As CursorTypeEnum = adOpenStatic, _
                Optional ByVal cnLockType As LockTypeEnum = adLockReadOnly, _
                Optional ByVal CursorLocation As CursorLocationEnum = adUseClient) As ADODB.Recordset
    '返回指定SQL查询语句的记录集
    
    Dim cnBack As ADODB.Connection
    
    On Error GoTo LineErr

    Set gfBackRecordset = New ADODB.Recordset
    Set cnBack = gfBackConnection(gID.CnString, CursorLocation)
    If cnBack.State = adStateClosed Then Exit Function
    gfBackRecordset.CursorLocation = CursorLocation
    gfBackRecordset.Open cnSQL, cnBack, cnCursorType, cnLockType
    
    Exit Function

LineErr:
    Call gsAlarmAndLog("返回记录集异常")

End Function


Public Function gfBackOneChar(Optional ByVal CharType As genmCharType = udUpperLowerNum) As String
    '随机返回一个字符（字母或数字）
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
    '判断文件、文件目录 是否存在

    Dim strBack As String
        
    On Error GoTo LineErr
    
    If Len(strPath) > 0 Then    '空字符串不算
        strBack = Dir(strPath, vbDirectory + vbHidden + vbReadOnly + vbSystem)
        If Len(strBack) > 0 Then gfFileExist = True
    End If
  
    Exit Function
    
LineErr:
    Call gsAlarmAndLog("判断文件异常")
    
End Function


Public Function gfFileExistEx(ByVal strPath As String) As gtypValueAndErr
    '另一种返回值方式：来判断文件、文件目录 是否存在
    '专供后面的过程gfFileRepair调用
    
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
    Call gsAlarmAndLog("文件判断返回异常")
    
End Function


Public Function gfFileOpen(ByVal strFilePath As String) As gtypValueAndErr
    '打开指定全路径的文件
    
    Dim lngRet As Long
    Dim strDir As String
    
    On Error GoTo LineErr
    
    If gfFileExist(strFilePath) Then
        
        lngRet = ShellExecute(GetDesktopWindow, "open", strFilePath, vbNullString, vbNullString, vbNormalFocus)
        If lngRet = SE_ERR_NOASSOC Then     '没有关联的程序
             strDir = Space(260)
             lngRet = GetSystemDirectory(strDir, Len(strDir))
             strDir = Left(strDir, lngRet)
             
            '显示打开方式窗口
            Call ShellExecute(GetDesktopWindow, vbNullString, "RUNDLL32.EXE", "shell32.dll,OpenAs_RunDLL " & strFilePath, strDir, vbNormalFocus)
            gfFileOpen.ErrNum = -1   '不成功，也没异常
        Else
            gfFileOpen.Result = True
        End If
        
    End If
    
    Exit Function
    
LineErr:
    gfFileOpen.ErrNum = Err.Number
    Call gsAlarmAndLog("文件打开异常")
    
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


Public Function gfFormLoad(ByVal strFormName As String) As Boolean
    '判断指定窗口是否被加载了
    
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
    '''敏感字符检测
    
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
