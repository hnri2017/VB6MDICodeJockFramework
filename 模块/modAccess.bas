Attribute VB_Name = "modAccess"
Option Explicit



Public Type typeMDBCommonVariant
    dbName As String
    dbPassword As String
    dbPath As String
    dbConn As String
    
    tbUser As String
    tbDept As String
    tbRole As String
    tbFunc As String
    tbUserRole As String
    tbRoleFunc As String
    tbOperation As String
    
    fdUserID As String
    fdUserName As String
    fdUserPassword As String
    fdUserFullName As String
    fdUserSex As String
    fdUserState As String
    fdUserDeptID As String
    fdUserMemo As String
    fdUserCreateTime As String
    fdUserCreateMan As String
    
    fdDeptID As String
    fdDeptName As String
    fdDeptParentID As String
    
    fdRoleID As String
    fdRoleName As String
    fdRoleDeptID As String
    
    fdFuncID As String
    fdFuncName As String
    fdFuncTitle As String
    fdFuncType As String
    fdFuncParentID As String
    
    fdLogID As String
    fdLogType As String
    fdLogContent As String
    fdLogTime As String
    fdLogTable As String
    fdLogFormName As String
    fdLogUserFullName As String
    fdLogPCIP As String
    fdLogPCName As String
    
    fdRFRoleID As String
    fdRFFuncID As String
    
    fdURUserID As String
    fdURRoleID As String
    
End Type

Public gVar As typeMDBCommonVariant




Public Sub gsMDBInitialize()
    With gVar
        .dbName = "DBCORE.mdb" ' "DBCORE.mdb"
        .dbPassword = "dbcoremdb"
        .dbPath = App.Path & "\" & .dbName
        .dbConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=" & .dbPath & ";" & _
                  "Jet OLEDB:Database Locking Mode=1;" & _
                  "Jet OLEDB:Database Password=" & .dbPassword & ";"
        
        .tbUser = "tb_Test_Sys_User"
        .tbDept = "tb_Test_Sys_Department"
        .tbRole = "tb_Test_Sys_Role"
        .tbFunc = "tb_Test_Sys_Func"
        .tbRoleFunc = "tb_Test_Sys_RoleFunc"
        .tbUserRole = "tb_Test_Sys_UserRole"
        .tbOperation = "tb_Test_Sys_OperationLog"
        
        .fdUserID = "UserAutoID"
        .fdUserName = "UserLoginName"
        .fdUserPassword = "UserPassword"
        .fdUserFullName = "UserFullName"
        .fdUserSex = "UserSex"
        .fdUserState = "UserState"
        .fdUserDeptID = "UserDeptID"
        .fdUserMemo = "UserMemo"
        
        .fdDeptID = "DeptID"
        .fdDeptName = "DeptName"
        .fdDeptParentID = "ParentID"
        
        .fdRoleDeptID = "DeptID"
        .fdRoleID = "RoleAutoID"
        .fdRoleName = "RoleName"
        
        .fdFuncID = "FuncAutoID"
        .fdFuncName = "FuncName"
        .fdFuncParentID = "FuncParentID"
        .fdFuncTitle = "FuncCaption"
        .fdFuncType = "FuncType"
        
        .fdLogID = "LogID"
        .fdLogType = "LogType"
        .fdLogContent = "LogContent"
        .fdLogTime = "LogTime"
        .fdLogTable = "LogTable"
        .fdLogFormName = "LogFormName"
        .fdLogUserFullName = "LogUserFullName"
        .fdLogPCIP = "LogPCIP"
        .fdLogPCName = "LogPCName"
        
        .fdRFRoleID = "RoleAutoID"
        .fdRFFuncID = "FuncAutoID"
        
        .fdURUserID = "UserAutoID"
        .fdURRoleID = "RoleAutoID"
        
    End With

End Sub

Public Sub gsRebuildDB()
    '重建数据库与数据表，记得先引用ADOEx与ADO
    'msado25.tlb--Microsoft ActiveX Data Objects 2.5 Library
    'msadox.dll--Microsoft ADO Ext. 6.0 for DDL and Security
    
    Dim strDir As String
    Dim dbNew As ADOX.Catalog
    Dim tbNew As ADOX.Table
    
    On Error GoTo LineErr
    
    strDir = Dir(gVar.dbPath)
    If Len(strDir) > 0 Then
        If InStr(strDir, gVar.dbName) > 0 Then
            Kill gVar.dbPath
        End If
    End If
    
    Set dbNew = New ADOX.Catalog
    dbNew.Create gVar.dbConn    '新建数据库文件，并打开该连接
    
    '添加表gVar.tbUser
    Set tbNew = New ADOX.Table
    With tbNew   '添加字段
        .ParentCatalog = dbNew
        .Name = gVar.tbUser
        .Columns.Append gVar.fdUserID, adInteger
        .Columns(gVar.fdUserID).Properties("AutoIncrement") = True  '默认与增量皆为1。似乎不能设置这两个变量。
        .Columns.Append gVar.fdUserName, adVarWChar, 50
        .Columns.Append gVar.fdUserPassword, adVarWChar, 60
        .Columns.Append gVar.fdUserFullName, adVarWChar, 50
        .Columns.Append gVar.fdUserSex, adVarWChar, 2
        .Columns.Append gVar.fdUserState, adVarWChar, 50
        .Columns.Append gVar.fdUserDeptID, adInteger
        .Columns.Append gVar.fdUserMemo, adVarWChar, 200
    End With
    dbNew.Tables.Append tbNew   '创建表。注意参数是表对象，不是表名称
    
    
'Dim colNew As ADOX.Column
'Set colNew = tbNew.Columns(gVar.fdUserID)
'Debug.Print "AutoIncrement=【" & colNew.Properties("AutoIncrement") & "】"
'Debug.Print "Default=【" & colNew.Properties("Default") & "】"
'Debug.Print "Description=【" & colNew.Properties("Description") & "】"
'Debug.Print "Nullable=【" & colNew.Properties("Nullable") & "】"
'Debug.Print "Fixed Length=【" & colNew.Properties("Fixed Length") & "】"
'Debug.Print "Seed=【" & colNew.Properties("Seed") & "】"
'Debug.Print "Increment=【" & colNew.Properties("Increment") & "】"
'Debug.Print "Jet OLEDB:Column Validation Text=【" & colNew.Properties("Jet OLEDB:Column Validation Text") & "】"
'Debug.Print "Jet OLEDB:Column Validation Rule=【" & colNew.Properties("Jet OLEDB:Column Validation Rule") & "】"
'Debug.Print "Jet OLEDB:IISAM Not Last Column=【" & colNew.Properties("Jet OLEDB:IISAM Not Last Column") & "】"
'Debug.Print "Jet OLEDB:AutoGenerate=【" & colNew.Properties("Jet OLEDB:AutoGenerate") & "】"
'Debug.Print "Jet OLEDB:One BLOB per Page=【" & colNew.Properties("Jet OLEDB:One BLOB per Page") & "】"
'Debug.Print "Jet OLEDB:Compressed UNICODE Strings=【" & colNew.Properties("Jet OLEDB:Compressed UNICODE Strings") & "】"
'Debug.Print "Jet OLEDB:Allow Zero Length=【" & colNew.Properties("Jet OLEDB:Allow Zero Length") & "】"
'Debug.Print "Jet OLEDB:Hyperlink=【" & colNew.Properties("Jet OLEDB:Hyperlink") & "】"
'Set colNew = Nothing
    
    
    '添加表gVar.tbDept
    Set tbNew = New ADOX.Table
    With tbNew
        .ParentCatalog = dbNew
        .Name = gVar.tbDept
        .Columns.Append gVar.fdDeptID, adInteger
        .Columns(gVar.fdDeptID).Properties("AutoIncrement") = True
        .Columns.Append gVar.fdDeptName, adVarWChar, 50
        .Columns.Append gVar.fdDeptParentID, adInteger
    End With
    dbNew.Tables.Append tbNew
    
    '添加表gVar.tbRole
    Set tbNew = New ADOX.Table
    With tbNew
        .Name = gVar.tbRole
        .ParentCatalog = dbNew
        .Columns.Append gVar.fdRoleID, adInteger
        .Columns(gVar.fdRoleID).Properties("AutoIncrement") = True
        .Columns.Append gVar.fdRoleName, adVarWChar, 50
        .Columns.Append gVar.fdRoleDeptID, adInteger
        .Columns.Refresh
    End With
    dbNew.Tables.Append tbNew
    
    '添加表gVar.tbFunc
    Set tbNew = New ADOX.Table
    With tbNew
        .ParentCatalog = dbNew
        .Name = gVar.tbFunc
        .Columns.Append gVar.fdFuncID, adInteger
        .Columns(gVar.fdFuncID).Properties("AutoIncrement") = True
        .Columns.Append gVar.fdFuncName, adVarWChar, 50
        .Columns.Append gVar.fdFuncTitle, adVarWChar, 50
        .Columns.Append gVar.fdFuncType, adVarWChar, 50
        .Columns.Append gVar.fdFuncParentID, adInteger
    End With
    dbNew.Tables.Append tbNew
    
'''    '单独直接修改某个字段示例
'''    dbNew.Tables(gVar.tbDept).Columns(gVar.fdDeptName).Name = "DeptTitle"
'''    dbNew.Tables(gVar.tbDept).Columns.Refresh
    
    '添加表gVar.tbOperation
    Set tbNew = New ADOX.Table
    With tbNew
        .ParentCatalog = dbNew
        .Name = gVar.tbOperation
        .Columns.Append gVar.fdLogID, adInteger
        .Columns(gVar.fdLogID).Properties("AutoIncrement") = True
        .Columns.Append gVar.fdLogType, adVarWChar, 50
        .Columns.Append gVar.fdLogContent, adVarWChar, 200
        .Columns.Append gVar.fdLogTime, adDate
        .Columns.Append gVar.fdLogTable, adVarWChar, 50
        .Columns.Append gVar.fdLogFormName, adVarWChar, 50
        .Columns.Append gVar.fdLogUserFullName, adVarWChar, 50
        .Columns.Append gVar.fdLogPCIP, adVarWChar, 50
        .Columns.Append gVar.fdLogPCName, adVarWChar, 50
    End With
    dbNew.Tables.Append tbNew
    
    '添加表gVar.tbRoleFunc。注意是先创建表，再添加的字段
    Set tbNew = New ADOX.Table
    tbNew.Name = gVar.tbRoleFunc
    dbNew.Tables.Append tbNew
    With tbNew.Columns
        .Append gVar.fdRFRoleID, adInteger
        .Append gVar.fdRFFuncID, adInteger
        .Refresh
    End With
    
    '添加表gVar.tbUserRole
    Set tbNew = New ADOX.Table
    tbNew.Name = gVar.tbUserRole
    dbNew.Tables.Append tbNew
    With tbNew.Columns
        .Append gVar.fdURUserID, adInteger
        .Append gVar.fdURRoleID, adInteger
        .Refresh
    End With
    
    '添加系统管理员用户
    Dim rsAdd As New ADODB.Recordset
    With rsAdd
        .CursorLocation = adUseClient
        .Open gVar.tbUser, dbNew.ActiveConnection, adOpenStatic, adLockBatchOptimistic
        If .RecordCount = 0 Then
            .AddNew
            .Fields(gVar.fdUserName) = gID.UserAdmin
            .Fields(gVar.fdUserPassword) = "6596841752979759m3swwgnCybHLiHNbcX22YewUqSuEjEUMgOoXgpj1IAkx"
            .Fields(gVar.fdUserFullName) = "系统管理员"
            .Fields(gVar.fdUserSex) = "男"
            .Fields(gVar.fdUserState) = "启用"
            .Fields(gVar.fdUserDeptID) = 0
            .Fields(gVar.fdUserMemo) = Now
            .AddNew
            .Fields(gVar.fdUserName) = gID.UserSystem
            .Fields(gVar.fdUserPassword) = "6845499556359466788m7sywYnAkocd1be8ag1ZqFHFnTlXowf62RqJ1JbwL"
            .Fields(gVar.fdUserFullName) = "系统人员"
            .Fields(gVar.fdUserSex) = "女"
            .Fields(gVar.fdUserState) = "启用"
            .Fields(gVar.fdUserDeptID) = 0
            .Fields(gVar.fdUserMemo) = Now
            .UpdateBatch
        End If
        .Close
    End With
    Set rsAdd = Nothing
    
    Set dbNew.ActiveConnection = Nothing
    Set tbNew = Nothing
    Set dbNew = Nothing
    
    Exit Sub
    
LineErr:
    If Not rsAdd Is Nothing Then If rsAdd.State = adStateOpen Then rsAdd.Close
    Set rsAdd = Nothing
    Set dbNew.ActiveConnection = Nothing
    Set tbNew = Nothing
    Set dbNew = Nothing
    
    On Error Resume Next
    strDir = Dir(gVar.dbPath)
    If Len(strDir) > 0 Then
        If InStr(strDir, gVar.dbName) > 0 Then
            Kill gVar.dbPath
        End If
    End If
    
    Call gsAlarmAndLog("返回记录集异常")
    
End Sub

