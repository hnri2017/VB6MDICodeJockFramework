Attribute VB_Name = "modDeclare"
Option Explicit


Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const SE_ERR_NOASSOC = 31 '没有关联的程序

Public Const gconAscAdd As Integer = 5      '简单加解密中字符转化的增量
Public Const gconAddLenStart As Integer = 10    '加在密文开始部分的字符个数
Public Const gconSumLen As Integer = 60     '密文的总字符数
Public Const gconMaxPWD As Integer = 20     '密码的最大字符数


Public Type gtypCommandBarID
    'CommandBars的ID集合
    
    Sys As Long
    
    SysExit As Long
    SysReLogin As Long
    SysModifyPassword As Long
    SysDepartment As Long
    SysUser As Long
    SysLog As Long
    SysRole As Long
    SysFunc As Long
    
    SysPageSet As Long
    SysPrint As Long
    SysPrintPreview As Long
    SysOutToExcel As Long
    SysOutToWord As Long
    SysOutToText As Long
    
    SysSearch As Long
    SysSearch1Label As Long
    SysSearch2TextBox As Long
    SysSearch3Button As Long
    SysSearch4ListBoxCaption As Long
    SysSearch4ListBoxFormID As Long
    SysSearch5Go As Long
    
    
    TestWindow As Long
    
    TestWindowFirst As Long
    TestWindowSecond As Long
    TestWindowThird As Long
    TestWindowThour As Long
    TestWindowMDB As Long
    
    Help As Long
    HelpAbout As Long
    HelpDocument As Long
    
    
    Wnd As Long
    
    WndResetLayout As Long
    
    WndThemeCommandBars As Long
    WndThemeCommandBarsOffice2000 As Long
    WndThemeCommandBarsOfficeXp As Long
    WndThemeCommandBarsOffice2003 As Long
    WndThemeCommandBarsWinXP As Long
    WndThemeCommandBarsWhidbey As Long
    WndThemeCommandBarsResource As Long
    WndThemeCommandBarsRibbon As Long
    WndThemeCommandBarsVS2008 As Long
    WndThemeCommandBarsVS6 As Long
    WndThemeCommandBarsVS2010 As Long
    
    WndThemeTaskPanel As Long
    WndThemeTaskPanelOffice2000 As Long
    WndThemeTaskPanelOffice2003 As Long
    WndThemeTaskPanelNativeWinXP As Long
    WndThemeTaskPanelOffice2000Plain As Long
    WndThemeTaskPanelOfficeXPPlain As Long
    WndThemeTaskPanelOffice2003Plain As Long
    WndThemeTaskPanelNativeWinXPPlain As Long
    WndThemeTaskPanelToolbox As Long
    WndThemeTaskPanelToolboxWhidbey As Long
    WndThemeTaskPanelListView As Long
    WndThemeTaskPanelListViewOfficeXP As Long
    WndThemeTaskPanelListViewOffice2003 As Long
    WndThemeTaskPanelShortcutBarOffice2003 As Long
    WndThemeTaskPanelResource As Long
    WndThemeTaskPanelVisualStudio2010 As Long
    
    WndThemeSkin As Long
    WndThemeSkinCodejock As Long
    WndThemeSkinOffice2007 As Long
    WndThemeSkinOffice2010 As Long
    WndThemeSkinVista As Long
    WndThemeSkinWinXPRoyale As Long
    WndThemeSkinWinXPLuna As Long
    WndThemeSkinZune As Long
    WndThemeSkinSet As Long
    
    WndSon As Long
    WndSonVbCascade As Long
    WndSonVbTileHorizontal As Long
    WndSonVbTileVertical As Long
    WndSonVbArrangeIcons As Long
    WndSonVbAllMin As Long
    WndSonVbAllBack As Long
    WndSonCloseAll As Long
    WndSonCloseCurrent As Long
    WndSonCloseLeft As Long
    WndSonCloseRight As Long
    WndSonCloseOther As Long
    
    
    Other As Long
    
    OtherPane As Long
    OtherPaneIDFirst As Long
    OtherPaneIDSecond As Long
    OtherPaneMenuTitle As Long
    OtherPaneMenuPopu As Long
    OtherPaneMenuPopuExpand As Long
    OtherPaneMenuPopuAutoFold As Long
    OtherPaneMenuPopuFold As Long
    
    OtherSave As Long               '注册表中相关值与名称
    OtherSaveAppName As String
    OtherSaveRegistryKey As String
    OtherSaveCommandBarsSection As String
    OtherSaveDockingPaneSection As String
    OtherSaveWidth As Long
    OtherSaveHeight As Long
    OtherSaveSettings As String
    OtherSaveSkinPath As String
    OtherSaveSkinIni As String
    OtherSaveSkinID As String
    OtherSaveUserLast As String     'Key名称-最后登陆过的用户名
    OtherSaveUserList As String     'Key名称-登陆过的用户名列表
    
    
    OtherTabWorkspacePopup As Long
    
    StatusBarPane As Long               '状态栏面板
    StatusBarPaneProgress As Long       '状态栏中进度条
    StatusBarPaneProgressText As Long   '状态栏中进度百分值
    StatusBarPaneUserInfo As Long       '状态栏中用户信息
    StatusBarPaneTime As Long           '状态中时间
    
    FolderStyles As String  '本地Style文件夹路径
    FolderBin As String     '本地Bin文件夹路径
    FolderNet As String     '网络共享文件夹路径
    FolderData As String    '本地Data文件夹路径
    
    FileLog As String       '日志文件全路径
    FileAppName As String   'App带扩展名的全名
    FileAppNet As String    '网络App全路径
    FileAppLoc As String    '本地App全路径
    FileSetupNet As String  '网络安装包全路径
    FileSetupLoc As String  '本地安装包存放全路径
    
    SkinPath As String      '主题资源文件名
    SkinIni As String       '主题配置文件名
    
    UserAutoID As String    '用户标识ID
    UserLoginName As String '用户登陆名
    UserNickName As String  '用户昵称
    UserFullName As String  '用户姓名
    UserPassword As String  '用户密码
    UserDepartment As String    '用户所在部门
    UserLoginIP As String       '用户登陆电脑IP
    UserComputerName As String  '用户登陆电脑名称
    UserAdmin As String         '特别用户名：系统管理员
    UserSystem As String        '特别用户名：系统管理员
    
    CnSource As String      '连接数据库服务器名称或IP地址
    CnUserID As String      '连接数据库用户名
    CnPassword As String    '连接数据库密码
    CnDatabase As String    '连接的数据库名
    CnString As String      '数据库连接字符串全称
    rsRF As New ADODB.Recordset '保存用户的所有权限信息
    
    FuncButton As String    '功能类别：按钮
    FuncForm As String      '功能类别：窗口
    FuncControl As String   '功能类别：其它控件
    FuncMainMenu As String  '功能类别：主菜单
    
    VSPrintPageSet As Boolean 'VS表格控件的页面设置状态
    
End Type

Public Type gtypValueAndErr '用于返回布尔值的过程，顺便返回异常代号
    Result As Boolean
    ErrNum As Long
End Type

Public Enum genmFileOpenType    '打开文件方式
    udAppend    '以顺序型访问，把字符追加到文件
    udBinary    '以二进制访问
    udInput     '以顺序型访问，从文件输入字符
    udOutput    '以顺序型访问，向文件输出字符
    udRandom    '以随机方式
End Enum

Public Enum genmFileWriteType   '写入文件方式
    udPut       '用Get读出.For Binary、Random.
    udWrite     '用Input读出
    udPrint     '用Line Input 或 Input读出
End Enum

Public Enum genmCharType    '返回字符类型
    udUpperCase = 4     '仅大写字母
    udLowerCase = 1     '仅小写字母
    udNumber = 2        '仅数字
    udUpperLowerNum = 7 '大写、小写、数字
End Enum

Public Enum genmLogType '操作日志类型增、删、改、查
    udSelect        '单个查询
    udInsert
    udDelete
    udUpdate
    udSelectBatch   '多个查询
    udInsertBatch
    udDeleteBatch
    udUpdateBatch
End Enum


Public gID As gtypCommandBarID   '全局CommandBars的ID变量
Public gMDI As MDIForm          '全局主窗体引用




