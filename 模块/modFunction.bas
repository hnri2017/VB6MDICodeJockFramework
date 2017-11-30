Attribute VB_Name = "modFunction"
Option Explicit


Public Function gfAsciiAdd(ByVal strIn As String) As String
    '返回传入字符的Ascii码值加N后 对应的字符。
    '与gAsciiSub过程互逆
    '注意1：暂时设定支持字母和数字。
    '注意2：输入的字符对应的ASCII值不能超过122即小写字母z。
    '注意3：字符增量N值大于0且不能超过5。
    
    Dim intASC As Integer
    
    If Len(strIn) = 0 Then Exit Function
    
    If gconAscAdd > 5 Or gconAscAdd = 0 Then
        MsgBox "字符增量大于0且不能超过5！", vbExclamation, "字符转化增量警告"
        Exit Function
    End If
    
    intASC = Asc(Left(strIn, 1))
    Select Case intASC
        Case 48 To 57, 65 To 90, 97 To 122
            
            intASC = intASC + gconAscAdd
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
            gfAsciiAdd = Chr(intASC)
            
        Case Else
            MsgBox "非法字符转化【" & strIn & "】！" & vbCrLf & "暂不支持数字和字母以外的字符！", vbExclamation, "不支持字符警告"
    End Select
    
End Function

Public Function gfAsciiSub(ByVal strIn As String) As String
    '返回传入字符的Ascii码值减N后 对应的字符。
    '与gAsciiAdd过程互逆
    '注意1：暂时设定只支持字母和数字。
    '注意2：输入的字符对应的ASCII值不能超过127。
    '注意3：字符增量N大于0且不能超过5。
    
    Dim intSub As Integer
    
    If Len(strIn) = 0 Then Exit Function
    
    If gconAscAdd > 5 Or gconAscAdd = 0 Then
        MsgBox "字符增量大于0且不能超过5！", vbExclamation, "字符转化增量警告"
        Exit Function
    End If
    
    intSub = Asc(Left(strIn, 1))
    Select Case intSub
        Case 48 To 57, 65 To 90, 97 To 122
            
            intSub = intSub - gconAscAdd
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
            gfAsciiSub = Chr(intSub)
            
        Case Else
            MsgBox "非法字符转化【" & strIn & "】！" & vbCrLf & "暂不支持数字和字母以外的字符！", vbExclamation, "不支持字符警告"
    End Select
    
End Function


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


Public Function gfBackLogType(Optional ByVal strType As genmLogType = udSelect) As String
    '返回日志操作类型
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
    '随机返回一个字符（字母或数字）
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
    '解密输入的字符串密文为明文
    '密文长度限定为gconSumLen位
    
    Dim strVar As String    '中间变量
    Dim strPt As String     '明文
    Dim strMid As String    '截取输入字符串中的每一个字符
    Dim intMid As Integer, K As Integer, c As Integer, R As Integer   '变量
    
    strIn = Trim(strIn) '去空格
    c = Len(strIn)
    If c <> gconSumLen Then GoTo LineBreak
    
    '一、获取密文中填充的无用字符个数、明文的长度
    R = Val(Mid(strIn, 2, 1))       '截取密文的第二位，其值即密文第gconAddLenStart+1位后填充的无用随机数个数
    If R < 1 Then GoTo LineBreak
    
    intMid = Val(Left(strIn, 1))    '截取密文的第一位，计算填充字符个数的值的 个位上的数字
    c = IIf(intMid < (gconAddLenStart - 2), intMid, gconAddLenStart - 2)  '通过第一位的数值计算出填充数值的十位上的数字所在位置
    K = Val(Mid(strIn, c + 2 + 1, 1))   '截取填充数值的十位上的数字
    c = Val(CStr(K) & CStr(intMid))     '得出真正的 填充字符 总数值
    If (c < (gconSumLen - gconMaxPWD)) Or (c > (gconSumLen - 1)) Then GoTo LineBreak
    
    c = gconSumLen - c  '得出明文的长度
    c = c * 2           '因为明文中插入了相同个数的随机字符
    
    '二、删除加在密文前面的gconAddLenStart+ 1 + R 个字符 和 加在密文最后的字符
    strVar = Mid(strIn, gconAddLenStart + 1 + R + 1, c)
    If Len(strVar) <> c Then GoTo LineBreak
    
    '三、解密剩下的strVar字符
    For K = 1 To c Step 2
        strPt = strPt & gfAsciiSub(Mid(strVar, K, 1))
    Next
    If Len(strPt) <> c / 2 Then GoTo LineBreak
    
    gfDecryptSimple = strPt  '将解密好的密文返回给函数的调用者
    
    Exit Function
    
LineBreak:
    MsgBox "密文被破坏，无法解密！", vbExclamation, "密文警告"
    
End Function

Public Function gfEncryptSimple(ByVal strIn As String) As String
    '将传入的字符串(明文)进行简单加密，生成密文并返回给调用者
    '明文长度<=20个字符，且只能是大写或小写字母、数字，否则转化时会报错
    
    Dim strEt As String     '密文
    Dim strMid As String    '截取输入字符串中的每一个字符
    Dim strTen As String    '密文的前10个字符
    Dim K As Integer, J As Integer, R As Integer  '变量
    Dim c As Integer        '明文的字符个数
    Dim intFill As Integer  '填充字符数
    Dim intRightNum As Integer      'strFill 个位上的数字
    Dim intAddLenEnd As Integer     '加在最后的字符数量

    c = Len(Trim(strIn))
    If c = 0 Then
        MsgBox "传入字符不能为空字符，且不能有空格！", vbCritical, "空字符警报"
        Exit Function
    End If
    strIn = Left(strIn, gconMaxPWD) '截取前gconMaxPWD(20)字符
    c = Len(strIn)  '重新获取字符个数。重要！
    
    '一、将字符串中的每个字符的ASCII值前进N位并插入一个随机字符得到一新字符串
    For K = 1 To c
        strEt = strEt & gfAsciiAdd(Mid(strIn, K, 1)) & gfBackOneChar(udUpperLowerNum)
    Next
    If Len(strEt) <> (c * 2) Then
        MsgBox "输入字符不规范，只能是数字或字母！", vbCritical, "字符警报"
        Exit Function
    End If
    
    '二、在转化后的字符串strEt前面总是加入gconAddLenStart个字符
    '   在这gconAddLenStart个字符中包含明文的长度信息gconSumLen-C
    '   然后将gconSumLen-C的值的 个位与十位调换位置
    '   然后在strTen的第二位插入原strTen后应填充的随机数字个数
    intFill = gconSumLen - c        '计算去除明文个数后要填充的总字符个数
    intRightNum = intFill Mod 10    '获取个位上的数字
    strTen = CStr(intRightNum)      '将个位上的数字放在strTen的第一位,也即密文的第一位
    
    '根据strTen的第一位的值计算在其后插入的随机数字的个数
    J = IIf(intRightNum < (gconAddLenStart - 2), intRightNum, gconAddLenStart - 2)
    For K = 1 To J
        strTen = strTen & gfBackOneChar(udNumber)
    Next
    strTen = strTen & CStr(Int(intFill / 10))   '并上intFill的十位上的数字
    
    Do
        R = gfBackOneChar(udNumber)     '获取一个1~9中的随机数字
        If R > 0 Then Exit Do
    Loop
    strTen = Left(strTen, 1) & CStr(R) & Right(strTen, Len(strTen) - 1)
    
    '若strTen的长度不够gconAddLenStart位，则填充随机数字,再在strTen后面并上随机R个数字
    J = (gconAddLenStart - 2 - J) + R
    For K = 1 To J
        strTen = strTen & gfBackOneChar(udNumber)
    Next
    strEt = strTen & strEt
    
    '三、在strEt后追加intAddLenEnd个随机字符凑成gconSumLen个字符的最终密文
    intAddLenEnd = gconSumLen - (c * 2) - gconAddLenStart - R - 1
    If intAddLenEnd > 0 Then
        For K = 1 To intAddLenEnd
            strEt = strEt & gfBackOneChar(udUpperLowerNum)
        Next
    End If
    
    gfEncryptSimple = strEt  '最后将strEt赋给函数的返回值
    
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
