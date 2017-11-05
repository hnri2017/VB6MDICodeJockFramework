Attribute VB_Name = "modSub"
Option Explicit


Public Sub Main()
    
    Set gMDI = frmSysMDI    '初始化主窗体引用全局变量

    With gID
        .Sys = 100
        .SysExit = 101
        .SysModifyPassword = 102
        .SysReLogin = 103
        
        .SysOutToExcel = 104
        .SysOutToText = 105
        .SysOutToWord = 106
        .SysPrint = 107
        .SysPrintPreview = 108
        
        .SysSearch = 110
        .SysSearch1Label = 111
        .SysSearch2TextBox = 112
        .SysSearch3Button = 113
        .SysSearch4ListBoxCaption = 114
        .SysSearch4ListBoxName = 115
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
        
        .WndThemeSkinSet = 899
        
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
        
        .FileAppName = App.EXEName & ".exe"
        .FileAppLoc = App.Path & "\" & .FileAppName
'''        .FileAppNet = .FolderNet & .FileAppName
        .FileAppNet = .FileAppLoc
        .FileLog = App.Path & "\Data\Record.LOG"
        .FileSetupLoc = App.Path & "\" & App.EXEName & "Setup.exe"
'''        .FileSetupNet = .FolderNet & App.EXEName & "Setup.exe"
        .FileSetupNet = .FileSetupLoc
        
        .CnDatabase = "db_Test"
        .CnPassword = "test"
        .CnSource = "192.168.2.9"
        .CnUserID = "wzd_test"
        .CnString = "Provider=SQLOLEDB;Persist Security Info=False;DataSource=" & .CnSource & _
                    ";UID=" & .CnUserID & ";PWD=" & .CnPassword & ";DataBase=" & .CnDatabase
        
        
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


Public Sub gsGridToExcel(ByRef gridControl As Control, Optional ByVal TimeCol As Long = -1, Optional ByVal TimeStyle As String = "yyyy-MM-dd HH:mm:ss")  '导出至Excel
    '将表格控件中的内容导出至Excel中
    '参数TimeCol：为控件中的时间列的列号，TimeStyle设定格式
    '最好引用Excel对象。运行时电脑上应有MSOFFICE软件。
    
'    Dim xlsOut As Excel.Application    '用这个申明好编程，编完后改为Object
    Dim xlsOut As Object
'    Dim sheetOut As Excel.Worksheet
    Dim sheetOut  As Object
    Dim R As Long, C As Long, I As Long, J As Long
    
    On Error Resume Next
    Screen.MousePointer = 13
    
    Set xlsOut = CreateObject("Excel.Application")
    xlsOut.Workbooks.Add
    Set sheetOut = xlsOut.ActiveSheet
    
    With gridControl
        R = .Rows
        C = .Cols
        For I = 0 To R - 1  '表格内容复制到Excel中
            For J = 0 To C - 1
                sheetOut.Cells(I + 1, J + 1) = .TextMatrix(I, J)
            Next
        Next
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
        .Rows(1).RowHeight = 23 '第一行行高
    End With
    
    xlsOut.Visible = True   '显示Excel文档
    
    Set sheetOut = Nothing
    Set xlsOut = Nothing
    Screen.MousePointer = 0
    
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
    
    frmOpen.Show OpenMode
    frmOpen.WindowState = FormWndState
    
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


