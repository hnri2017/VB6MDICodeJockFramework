Attribute VB_Name = "modSub"
Option Explicit


Public Sub Main()
    
    App.Title = "VB6+Codejock"
    Set gMDI = frmSysMDI    '初始化主窗体引用全局变量

    With gID
        .Sys = 100
        .SysExit = 101
        .SysModifyPassword = 102
        .SysReLogin = 103
        .SysDepartment = 104
        .SysUser = 105
        .SysLog = 106
        .SysRole = 107
        .SysFunc = 108
        
        
        .SysOutToExcel = 120
        .SysOutToText = 121
        .SysOutToWord = 122
        .SysPrint = 123
        .SysPrintPreview = 124
        
        .SysSearch = 110
        .SysSearch1Label = 111
        .SysSearch2TextBox = 112
        .SysSearch3Button = 113
        .SysSearch4ListBoxCaption = 114
        .SysSearch4ListBoxFormID = 115
        .SysSearch5Go = 116
        
        
        .TestWindow = 200
        .TestWindowFirst = 201
        .TestWindowSecond = 202
        .TestWindowThird = 203
        .TestWindowThour = 204
        
        .Wnd = 800
        
        .WndResetLayout = 801
        
        .WndThemeCommandBars = 810
        .WndThemeCommandBarsOffice2000 = 811
        .WndThemeCommandBarsOffice2003 = 812
        .WndThemeCommandBarsOfficeXp = 813
        .WndThemeCommandBarsResource = 814
        .WndThemeCommandBarsRibbon = 815
        .WndThemeCommandBarsVS2008 = 816
        .WndThemeCommandBarsVS2010 = 817
        .WndThemeCommandBarsVS6 = 818
        .WndThemeCommandBarsWhidbey = 819
        .WndThemeCommandBarsWinXP = 820

        .WndThemeTaskPanel = 840
        .WndThemeTaskPanelListView = 841
        .WndThemeTaskPanelListViewOffice2003 = 842
        .WndThemeTaskPanelListViewOfficeXP = 843
        .WndThemeTaskPanelNativeWinXP = 844
        .WndThemeTaskPanelNativeWinXPPlain = 845
        .WndThemeTaskPanelOffice2000 = 846
        .WndThemeTaskPanelOffice2000Plain = 847
        .WndThemeTaskPanelOffice2003 = 848
        .WndThemeTaskPanelOffice2003Plain = 849
        .WndThemeTaskPanelOfficeXPPlain = 850
        .WndThemeTaskPanelResource = 851
        .WndThemeTaskPanelShortcutBarOffice2003 = 852
        .WndThemeTaskPanelToolbox = 853
        .WndThemeTaskPanelToolboxWhidbey = 854
        .WndThemeTaskPanelVisualStudio2010 = 855
        
        .WndSon = 856
        .WndSonCloseAll = 857
        .WndSonCloseCurrent = 858
        .WndSonCloseLeft = 859
        .WndSonCloseOther = 860
        .WndSonCloseRight = 861
        .WndSonVbAllBack = 862
        .WndSonVbAllMin = 863
        .WndSonVbArrangeIcons = 864
        .WndSonVbCascade = 865
        .WndSonVbTileHorizontal = 866
        .WndSonVbTileVertical = 867
        
        
        .WndThemeSkin = 870
        .WndThemeSkinCodejock = 871
        .WndThemeSkinOffice2007 = 872
        .WndThemeSkinOffice2010 = 873
        .WndThemeSkinVista = 874
        .WndThemeSkinWinXPLuna = 875
        .WndThemeSkinWinXPRoyale = 876
        .WndThemeSkinZune = 877
        
        .WndThemeSkinSet = 802
        
        .Help = 900
        .HelpAbout = 901
        .HelpDocument = 902
        
        
        '请将所有菜单CommandBrs的ID值设置在2000以下，。
        
        .Other = 2000
        .OtherPane = 2100
        .OtherPaneIDFirst = 2101
        .OtherPaneIDSecond = 2102
        .OtherPaneMenuPopu = 2103
        .OtherPaneMenuPopuAutoFold = 2104
        .OtherPaneMenuPopuExpand = 2105
        .OtherPaneMenuPopuFold = 2106
        .OtherPaneMenuTitle = 2107
        
        .OtherSave = 2200
        .OtherSaveWidth = 15360
        .OtherSaveHeight = 11520
        .OtherSaveSettings = "Settings"
        .OtherTabWorkspacePopup = 2201
        .OtherSaveSkinID = "SkinFWID"
        .OtherSaveSkinIni = "SkinFWIni"
        .OtherSaveSkinPath = "SkinFWPath"
        .OtherSaveUserList = "UserList"
        .OtherSaveUserLast = "UserLast"
        
        .StatusBarPane = 2300
        .StatusBarPaneProgress = 2301
        .StatusBarPaneProgressText = 2302
        .StatusBarPaneTime = 2303
        .StatusBarPaneUserInfo = 2304
        
        .FolderBin = App.Path & "\Bin\"
        .FolderNet = "\\192.168.12.100\玮之度\部门数据\公共数据\WZDMS专用(勿动)\玮之度管理系统\"
        .FolderStyles = App.Path & "\Styles\"
        .FolderData = App.Path & "\Data\"
        
        .FileAppName = App.EXEName & ".exe"
        .FileAppLoc = App.Path & "\" & .FileAppName
'''        .FileAppNet = .FolderNet & .FileAppName
        .FileAppNet = .FileAppLoc
        .FileLog = App.Path & "\Data\Record.LOG"
        .FileSetupLoc = App.Path & "\" & App.EXEName & "Setup.exe"
'''        .FileSetupNet = .FolderNet & App.EXEName & "Setup.exe"
        .FileSetupNet = .FileSetupLoc
        
        .UserAdmin = "Admin"    '两个特殊用户
        .UserSystem = "System"  '两个特殊用户
        
        .CnDatabase = "db_Test"
        .CnPassword = "test"
        .CnSource = "192.168.2.9"
        .CnUserID = "wzd_test"
        .CnString = "Provider='SQLOLEDB';Persist Security Info=False;Data Source='" & .CnSource & _
                    "';User ID='" & .CnUserID & "';Password='" & .CnPassword & _
                    "';Initial Catalog='" & .CnDatabase & "';"   '在自己64位系统电脑上Data Source中间要空格隔开才能建立连接，在这里可以不用，不知为何
        
        .FuncButton = "按钮"
        .FuncControl = "其它"
        .FuncForm = "窗口"
        .FuncMainMenu = "主菜单"
        
    End With
    
    '设置窗口主题
    gMDI.skinFW.ApplyOptions = xtpSkinApplyColors Or xtpSkinApplyFrame Or xtpSkinApplyMenus Or xtpSkinApplyMetrics
    gMDI.skinFW.ApplyWindow gMDI.hwnd
    gID.SkinPath = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveSkinPath, "")
    gID.SkinIni = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveSkinIni, "")
    Call gMDI.gmsThemeSkinSet(gID.SkinPath, gID.SkinIni)

    frmSysLogin.Show    '显示登陆窗口

End Sub


Public Sub gsAlarmAndLog(Optional ByVal strErr As String, Optional ByVal MsgButton As VbMsgBoxStyle = vbCritical)
    '异常提示并写下异常日志
    
    Dim strMsg As String
    
    strMsg = "异常代号：" & Err.Number & vbCrLf & "异常描述：" & Err.Description
    MsgBox strMsg, MsgButton, strErr
    Call gsFileWrite(gID.FileLog, strErr & vbTab & Replace(strMsg, vbCrLf, vbTab))
    
End Sub


Public Sub gsFileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As genmFileOpenType = udAppend, _
    Optional ByVal WriteMode As genmFileWriteType = udPrint)
    '将指定内容以指定的方式写入指定文件中
    
    Dim intNum As Integer
    Dim strTime As String
    
    If Not gfFileRepair(strFile) Then Exit Sub
    intNum = FreeFile
    
    On Error Resume Next
    
    Select Case OpenMode
        Case udBinary
            Open strFile For Binary As #intNum
        Case udInput
            Open strFile For Input As #intNum
        Case udOutput
            Open strFile For Output As #intNum
        Case Else   '其余皆当作udAppend
            Open strFile For Append As #intNum
    End Select
    
    strTime = Format(Now, "yyyy-MM-dd hh:mm:ss")
    Select Case WriteMode
        Case udWrite
            Write #intNum, strTime, strContent
        Case udPut
            Put #intNum, , strTime & vbTab & strContent
        Case Else   '其余皆当作udPrint
            Print #intNum, strTime, strContent
    End Select
    
    Close #intNum
    
End Sub


Public Sub gsFormScrollBar(ByRef frmCur As Form, ByRef ctlMv As Control, _
    ByRef Hsb As HScrollBar, ByRef Vsb As VScrollBar, _
    Optional ByVal lngMW As Long = 12000, _
    Optional ByVal lngMH As Long = 9000, _
    Optional ByVal lngHV As Long = 255)
    
    'frmCur：滚动条所在的窗体
    'ctlMv：窗体中的控件（除滚动条以外）都在此容器控件中
    'Hsb：窗体frmCur中水平滚动条控件
    'Vsb：窗体frmCur中垂直滚动条控件
    'lngMW：窗体不出现滚动条的宽度
    'lngMH：窗体不出现滚动条的高度
    'lngHV：滚动条的窄边宽度或高度。
    '***注意注意注意：滚动条控件需最后添加至窗体中，且不能放在容器控件ctlMv中*******
    
    Dim lngW As Long
    Dim lngH As Long
    Dim lngSW As Long
    Dim lngSH As Long
    Dim lngMin As Long
    
    lngW = frmCur.Width
    lngH = frmCur.Height
    lngSW = frmCur.ScaleWidth
    lngSH = frmCur.ScaleHeight
    lngMin = -120
    
    On Error Resume Next
    
    If lngW >= lngMW Then
        Hsb.Visible = False
        ctlMv.Left = -lngMin
    Else
        With Hsb
            .Move 0, lngSH - lngHV, lngSW, lngHV
            .Min = lngMin
            .Max = lngMW - lngW + lngHV
            .SmallChange = 10
            .LargeChange = 50
            .Visible = True
        End With
    End If
    
    If lngH >= lngMH Then
        Vsb.Visible = False
        ctlMv.Top = -lngMin
    Else
        With Vsb
            .Move lngSW - lngHV, 0, lngHV, IIf(Hsb.Visible, lngSH - lngHV, lngSH)
            .Min = lngMin
            .Max = lngMH - lngH + lngHV
            .SmallChange = 10
            .LargeChange = 50
            .Visible = True
        End With
    End If
    
'    '在窗体中添加窗口控件ctlMove，将所有其它控件放入此容器中，然
'    '后添加名称分别为Hsb\Vsb的水平\垂直滚动条在窗体中，最好留到最后放入窗体中
'    '然后在窗体中添加以下事件调用即可
'Private Sub Form_Resize()
'    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 12000, 9000)  '注意长、宽的修改
'End Sub
'Private Sub Hsb_Change()
'    ctlMove.Left = -Hsb.Value
'End Sub
'
'Private Sub Hsb_Scroll()
'    Call Hsb_Change    '当滑动滚动条中的滑块时会同时更新对应内容，以下同。
'End Sub
'
'Private Sub Vsb_Change()
'    ctlMove.Top = -Vsb.Value
'End Sub
'
'Private Sub Vsb_Scroll()
'    Call Vsb_Change
'End Sub

End Sub

Public Sub gsGridPrint(ByRef gridControl As Control)
    '打印表格内容
    
    Dim blnFlexCell As Boolean
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    
End Sub

Public Sub gsGridPrintPreview(ByRef gridControl As Control)   'FlexCell.Grid
    '预览表格内容
    
    Dim blnFlexCell As Boolean
    Dim blnVSGrid As Boolean
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    If TypeOf gridControl Is VSFlex8Ctl.VSFlexGrid Then blnVSGrid = True
    
    If blnFlexCell Then
        gridControl.PrintPreview
    End If
    
End Sub

Public Sub gsGridToExcel(ByRef gridControl As Control, Optional ByVal TimeCol As Long = -1, Optional ByVal TimeStyle As String = "yyyy-MM-dd HH:mm:ss")  '导出至Excel
    '将表格控件中的内容导出至Excel中
    '参数TimeCol：为控件中的时间列的列号，TimeStyle设定格式
    '最好引用Excel对象。运行时电脑上应有MSOFFICE软件。
    
'    Dim xlsOut As Excel.Application    '用这个申明好编程但要引用，编完后改为Object
    Dim xlsOut As Object
'    Dim sheetOut As Excel.Worksheet
    Dim sheetOut  As Object
    Dim blnFlexCell As Boolean
    Dim R As Long, C As Long, I As Long, J As Long
    
    On Error Resume Next
    Screen.MousePointer = 13
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    Set xlsOut = CreateObject("Excel.Application")
    xlsOut.Workbooks.Add
    Set sheetOut = xlsOut.ActiveSheet
    
    With gridControl
        R = .Rows
        C = .Cols
        '表格内容复制到Excel中
        If blnFlexCell Then
            For I = 0 To R - 1
                For J = 0 To C - 1
                    sheetOut.Cells(I + 1, J + 1) = .Cell(I, J).Text
                Next
            Next
        Else
            For I = 0 To R - 1
                For J = 0 To C - 1
                    sheetOut.Cells(I + 1, J + 1) = .TextMatrix(I, J)
                Next
            Next
        End If
    End With
    
    With sheetOut
        If TimeCol > -1 Then
            .Columns(TimeCol + 1).NumberFormatLocal = TimeStyle
        End If
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Bold = True '加粗显示(第一行默认标题行)
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Size = 12   '第一行12号字大小
        .Range(.Cells(2, 1), .Cells(R, C)).Font.Size = 10   '第二行以后10号字大小
        .Range(.Cells(1, 1), .Cells(R, C)).HorizontalAlignment = -4108  'xlCenter= -4108(&HFFFFEFF4)   '居中显示
        .Range(.Cells(1, 1), .Cells(R, C)).Borders.Weight = 2   'xlThin=2  '单元格显示黑色线宽
        .Columns.EntireColumn.AutoFit   '自动列宽
        .Rows(1).rowHeight = 23 '第一行行高
    End With
    
    xlsOut.Visible = True   '显示Excel文档
    
    Set sheetOut = Nothing
    Set xlsOut = Nothing
    Screen.MousePointer = 0
    
End Sub


Public Sub gsGridToText(ByRef gridControl As Control)
    '将传入的表格控件中的内容导出为文本文件
    
    Dim strFileName As String
    Dim blnFlexCell As Boolean
    Dim intFree As Integer
    Dim R As Long, C As Long, I As Long, J As Long
    Dim strTxt As String
    
    For I = 1 To 8
        strFileName = strFileName & gfBackOneChar(udNumber + udUpperCase) '文件名中的8个随机字符，不含小写字母
    Next
    strFileName = gID.FolderData & Format(Now, "yyyyMMddHHmmss_") & strFileName & ".txt"
    If Not gfFileRepair(strFileName) Then
        MsgBox "创建文件失败，请重试！", vbExclamation, "文件生成警告"
        Exit Sub
    End If
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    intFree = FreeFile
    Open strFileName For Output As #intFree
    With gridControl
        R = .Rows - 1
        C = .Cols - 1
        If blnFlexCell Then
            For I = 0 To R
                strTxt = ""
                For J = 0 To C
                    strTxt = strTxt & .Cell(I, J).Text & vbTab
                Next
                Print #intFree, strTxt
            Next
        Else
            For I = 0 To R
                strTxt = ""
                For J = 0 To C
                    strTxt = strTxt & .TextMatrix(I, J) & vbTab
                Next
                Print #intFree, strTxt
            Next
        End If
    End With
    
    Close
    
    Call gfFileOpen(strFileName)    '打开
    
End Sub


Public Sub gsGridToWord(ByRef gridControl As Control)
    '将指定表格中的内容导出至Word文档中
    
'    Dim wordApp As Word.Application
    Dim wordApp As Object
'    Dim docOut As Word.Document
    Dim docOut As Object
'    Dim tbOut As Word.Table
    Dim tbOut As Object
    Dim lngRows As Long, lngCols As Long
    Dim I As Long, J As Long
    Dim blnFlexCell As Boolean
    
    lngRows = gridControl.Rows
    lngCols = gridControl.Cols
    
    On Error Resume Next
'    Set wordApp = New Word.Application
    Set wordApp = CreateObject("Word.Application")
    Set docOut = wordApp.Documents.Add()
    Set tbOut = docOut.Tables.Add(docOut.Range, lngRows, lngCols, True)
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    If blnFlexCell Then
        For I = 0 To lngRows - 1
            For J = 0 To lngCols - 1
                tbOut.Cell(I + 1, J + 1).Range.Text = gridControl.Cell(I, J).Text
            Next
        Next
    Else
        For I = 0 To lngRows - 1
            For J = 0 To lngCols - 1
                tbOut.Cell(I + 1, J + 1).Range.Text = gridControl.TextMatrix(I, J)
            Next
        Next
    End If
    tbOut.Rows(1).Range.Bold = True             '第一行内容加粗
    tbOut.Range.ParagraphFormat.Alignment = 1   '表格内容居中显示
    Call tbOut.AutoFitBehavior(1)               '根据内容自动调整列宽
    
    wordApp.Visible = True
    
    Set tbOut = Nothing
    Set docOut = Nothing
    Set wordApp = Nothing
    
End Sub

Public Sub gsLoadAuthority(ByRef frmCur As Form, ByRef ctlCur As Control)
    '加载窗口中的控制权限
    
    Dim strUser As String, strForm As String, strCtlName As String
    
    strUser = LCase(gID.UserLoginName)
    strForm = LCase(frmCur.Name)
    strCtlName = LCase(ctlCur.Name)
    
    If strUser = LCase(gID.UserAdmin) Or strUser = LCase(gID.UserSystem) Then Exit Sub
    ctlCur.Enabled = False
    
    With gID.rsRF
        If .State = adStateOpen Then
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If strForm = LCase(.Fields("FuncFormName")) Then
                        If strCtlName = LCase(.Fields("FuncName")) Then
                            ctlCur.Enabled = True
                            Exit Do
                        End If
                    End If
                    .MoveNext
                Loop
            End If
        End If
    End With
    
End Sub

Public Sub gsLogAdd(ByRef frmCur As Form, Optional ByVal LogType As genmLogType = udSelect, _
    Optional ByVal strTable As String = "", Optional ByVal strContent As String = "")
    '添加操作日志
    
    Dim strType As String
    Dim strSQL As String
    Dim rsLog As ADODB.Recordset
    
    strType = gfBackLogType(LogType)
    
    strSQL = "EXEC sp_Test_Sys_LogAdd '" & strType & "','" & frmCur.Name & "," & frmCur.Caption & "','" & strTable & _
             "','" & strContent & "','" & gID.UserLoginName & "," & gID.UserFullName & "','" & gID.UserLoginIP & "','" & gID.UserComputerName & "'"
'Debug.Print strSQL
    Set rsLog = gfBackRecordset(strSQL, , adLockOptimistic)
    If rsLog.State = adStateOpen Then rsLog.Close
    Set rsLog = Nothing
    
End Sub


Public Sub gsNodeCheckCascade(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '结点的Checked属性级联变化
    
    If blnCheck Then    '=False时不处理
        Call gsNodeCheckUp(nodeCheck)
    End If
    
    Call gsNodeCheckDown(nodeCheck, blnCheck)
    
End Sub

Public Sub gsNodeCheckDown(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '不/勾选结点的所有子结点
    
    Dim nodeSon As MSComctlLib.Node
    Dim C As Long, K As Long
    
    C = nodeCheck.Children
    If C > 0 Then
        For K = 1 To C
            If K = 1 Then
                Set nodeSon = nodeCheck.Child
            Else
                Set nodeSon = nodeSon.Next
            End If
            If nodeSon.Checked <> blnCheck Then nodeSon.Checked = blnCheck
            If nodeSon.Children > 0 Then
                Call gsNodeCheckDown(nodeSon, blnCheck)
            End If
        Next
    End If
    
End Sub

Public Sub gsNodeCheckUp(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean = True)
    '勾选结点的所有父结点
    
    Dim nodeDad As MSComctlLib.Node
    
    If Not nodeCheck.Parent Is Nothing Then
        Set nodeDad = nodeCheck.Parent
        If Not nodeDad.Checked Then nodeDad.Checked = blnCheck
        If Not nodeDad.Parent Is Nothing Then
            Call gsNodeCheckUp(nodeDad)
        End If
    End If
    
End Sub


Public Sub gsOpenTheWindow(ByVal strFormName As String, _
    Optional ByVal OpenMode As FormShowConstants = vbModeless, _
    Optional ByVal FormWndState As FormWindowStateConstants = vbMaximized)
    '以指定窗口模式OpenMode与窗口FormWndState状态来打开指定窗体strFormName
    
    Dim frmOpen As Form
    Dim C As Long
    
    strFormName = LCase(strFormName)
    If gfFormLoad(strFormName) Then
        For C = 0 To Forms.Count - 1
            If LCase(Forms(C).Name) = strFormName Then
                Set frmOpen = Forms(C)
                Exit For
            End If
        Next
    Else
        Set frmOpen = Forms.Add(strFormName)
    End If
    
    frmOpen.WindowState = FormWndState
    frmOpen.Show OpenMode               '此句放最后，不能放上句前面，否则退出程序时MDI窗体不能完全关闭，可能因为CommandBars控件的原因。
        
End Sub


Public Sub gsUnCheckedAction(ByVal strFormName As String)
    '当窗口关闭时，去掉主窗体中cBS控件中被勾选的对应Action
    
    Dim actionCur As CommandBarAction
    
    strFormName = LCase(strFormName)
    For Each actionCur In gMDI.cBS.Actions
        If Len(actionCur.Key) > 0 Then
            If LCase(actionCur.Key) = strFormName Then
                actionCur.Checked = False
                Exit For
            End If
        End If
    Next
    
End Sub


