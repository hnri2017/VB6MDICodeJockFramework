Attribute VB_Name = "modSub"
Option Explicit


Public Sub Main()
    
    Set gMDI = frmSysMDI    '初始化主窗体引用全局变量

    With gID
        .Sys = 100
        .SysExit = 101
        .SysModifyPassword = 102
        .SysReLogin = 103
        
        
        .Help = 900
        .HelpAbout = 901
        .HelpDocument = 902
        
        
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
        
        .FolderStyles = App.Path & "\Styles\"
    End With
    
    '设置窗口主题
    gMDI.skinFW.ApplyOptions = xtpSkinApplyColors Or xtpSkinApplyFrame Or xtpSkinApplyMenus Or xtpSkinApplyMetrics
    gMDI.skinFW.ApplyWindow gMDI.hwnd
    gID.SkinPath = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveSkinPath, "")
    gID.SkinIni = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveSkinIni, "")
    Call gMDI.gmsThemeSkinSet(gID.SkinPath, gID.SkinIni)

    frmSysLogin.Show    '显示登陆窗口
    
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


