Attribute VB_Name = "modInitialize"
Option Explicit

Type typCommandBarID
    'CommandBars的ID集合
    
    Sys As Long
    SysExit As Long
    SysReLogin As Long
    SysModifyPassword As Long
    
    
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
    
    WndSon As Long
    WndSonVbCascade As Long
    WndSonVbTileHorizontal As Long
    WndSonVbTileVertical As Long
    WndSonVbArrangeIcons As Long
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
    
    OtherSave As Long
    OtherSaveAppName As String
    OtherSaveRegistryKey As String
    OtherSaveCommandBarsSection As String
    OtherSaveDockingPaneSection As String
    OtherSaveWidth As Long
    OtherSaveHeight As Long
    OtherSaveSettings As String
    OtherSaveSkinPath As String
    OtherSaveSkinIni As String
    
    OtherTabWorkspacePopup As Long
    
    StatusBarPane As Long
    StatusBarPaneProgress As Long
    StatusBarPaneProgressText As Long
    StatusBarPaneUserInfo As Long
    StatusBarPaneTime As Long
    
End Type

Public gID As typCommandBarID   '全局CommandBars的ID变量
Public gMDI As MDIForm          '全局主窗体引用

Sub Main()
    
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
        
        .WndSon = 860
        .WndSonCloseAll = 861
        .WndSonCloseCurrent = 862
        .WndSonCloseLeft = 863
        .WndSonCloseOther = 864
        .WndSonCloseRight = 865
        .WndSonVbArrangeIcons = 866
        .WndSonVbCascade = 867
        .WndSonVbTileHorizontal = 868
        .WndSonVbTileVertical = 869
        
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
        .OtherSaveSkinIni = "SkinFWIni"
        .OtherSaveSkinPath = "SkinFWPath"
        
        .StatusBarPane = 2300
        .StatusBarPaneProgress = 2301
        .StatusBarPaneProgressText = 2302
        .StatusBarPaneTime = 2303
        .StatusBarPaneUserInfo = 2304
        
    End With
    
    Set gMDI = frmSysMDI
    gMDI.skinFW.ApplyOptions = xtpSkinApplyColors Or xtpSkinApplyFrame Or xtpSkinApplyMenus Or xtpSkinApplyMetrics
    gMDI.skinFW.ApplyWindow gMDI.hWnd
    gMDI.skinFW.LoadSkin App.Path & "\styles\Codejock.cjstyles", "NormalBlack.ini"
    
    frmSysLogin.Show
    
End Sub
