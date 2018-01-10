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
    
'    On Error GoTo LineErr
    
    strDir = Dir(gVar.dbPath)
    If Len(strDir) > 0 Then
        If InStr(strDir, gVar.dbName) > 0 Then
            Kill gVar.dbPath
        End If
    End If
    
    Set dbNew = New ADOX.Catalog
    dbNew.Create gVar.dbConn    '新建数据库文件[***.mdb]
    
    '添加表gVar.tbUser
    Set tbNew = New ADOX.Table
    tbNew.Name = gVar.tbUser
    With tbNew.Columns   '添加字段
        .Append gVar.fdUserID, adInteger
        .Append gVar.fdUserName, adVarWChar, 50
        .Append gVar.fdUserPassword, adVarWChar, 60
        .Append gVar.fdUserFullName, adVarWChar, 50
        .Append gVar.fdUserSex, adVarWChar, 2
        .Append gVar.fdUserState, adVarWChar, 50
        .Append gVar.fdUserDeptID, adInteger
        .Append gVar.fdUserMemo, adVarWChar, 200
    End With
    dbNew.Tables.Append tbNew   '创建表。注意参数是表对象，不是表名称
    
    Set tbNew = New ADOX.Table
    tbNew.Name = gVar.tbUser
    Set tbNew.ParentCatalog = dbNew
    Dim colNew As New ADOX.Column
    colNew.Name = gVar.fdUserID
    Set colNew.ParentCatalog = dbNew
Debug.Print colNew.Properties("AutoIncrement"), colNew.Properties("Seed") = 1, colNew.Properties("Increment"), colNew.Properties("Default")
    colNew.Properties("AutoIncrement") = True
    colNew.Properties("Seed") = 1
    colNew.Properties("Increment") = 1
    colNew.Properties("Default") = 1
'    tbNew.Columns.Refresh
Debug.Print colNew.Properties("AutoIncrement"), colNew.Properties("Seed") = 1, colNew.Properties("Increment"), colNew.Properties("Default")
'    Debug.Print colNew.Properties.Count
'    Dim pP As ADOX.Property
'    For Each pP In colNew.Properties
'        Debug.Print pP.Name
'    Next
    
    '添加表gVar.tbDept
    Set tbNew = New ADOX.Table
    tbNew.Name = gVar.tbDept
    With tbNew.Columns
        .Append gVar.fdDeptID, adInteger
        .Append gVar.fdDeptName, adVarWChar, 50
        .Append gVar.fdDeptParentID, adInteger
    End With
    dbNew.Tables.Append tbNew
    
    '添加表gVar.tbRole。注意是先创建表，再添加的字段
    Set tbNew = New ADOX.Table
    tbNew.Name = gVar.tbRole
    dbNew.Tables.Append tbNew
    With dbNew.Tables(gVar.tbRole).Columns
        .Append gVar.fdRoleID, adInteger
        .Append gVar.fdRoleName, adVarWChar, 50
        .Append gVar.fdRoleDeptID, adInteger
        .Refresh
    End With
    
    '添加表gVar.tbFunc
    Set tbNew = New ADOX.Table
    tbNew.Name = gVar.tbFunc
    dbNew.Tables.Append tbNew
    Set tbNew = dbNew.Tables(gVar.tbFunc)
    With tbNew.Columns
        .Append gVar.fdFuncID, adInteger
        .Append gVar.fdFuncName, adVarWChar, 50
        .Append gVar.fdFuncTitle, adVarWChar, 50
        .Append gVar.fdFuncType, adVarWChar, 50
        .Append gVar.fdFuncParentID, adInteger
        .Refresh
    End With
    
'''    '单独直接修改某个字段示例
'''    dbNew.Tables(gVar.tbDept).Columns(gVar.fdDeptName).Name = "DeptTitle"
'''    dbNew.Tables(gVar.tbDept).Columns.Refresh
    
    '添加表gVar.tbOperation
    Set tbNew = New ADOX.Table
    tbNew.Name = gVar.tbOperation
    With tbNew.Columns
        .Append gVar.fdLogID, adInteger
        .Append gVar.fdLogType, adVarWChar, 50
        .Append gVar.fdLogContent, adVarWChar, 200
        .Append gVar.fdLogTime, adDate
        .Append gVar.fdLogTable, adVarWChar, 50
        .Append gVar.fdLogFormName, adVarWChar, 50
        .Append gVar.fdLogUserFullName, adVarWChar, 50
        .Append gVar.fdLogPCIP, adVarWChar, 50
        .Append gVar.fdLogPCName, adVarWChar, 50
    End With
    dbNew.Tables.Append tbNew
    
    '添加表gVar.tbRoleFunc
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
    Set tbNew = dbNew.Tables(gVar.tbUser)
    
    Set dbNew.ActiveConnection = Nothing
    Set tbNew = Nothing
    Set dbNew = Nothing
    
    Exit Sub
    
LineErr:
    Set dbNew.ActiveConnection = Nothing
    Set tbNew = Nothing
    Set dbNew = Nothing
    
    strDir = Dir(gVar.dbPath)
    If Len(strDir) > 0 Then
        If InStr(strDir, gVar.dbName) > 0 Then
            Kill gVar.dbPath
        End If
    End If
    
    Call gsAlarmAndLog("返回记录集异常")
    
End Sub

