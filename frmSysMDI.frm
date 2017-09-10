VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.DockingPane.v15.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "Codejock.TaskPanel.v15.3.1.ocx"
Begin VB.MDIForm frmSysMDI 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgListCommandBars 
      Left            =   3600
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picHide 
      Align           =   1  'Align Top
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   14700
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   14760
      Begin VB.PictureBox picTaskPL 
         Height          =   1455
         Left            =   480
         ScaleHeight     =   1395
         ScaleWidth      =   3195
         TabIndex        =   1
         Top             =   600
         Width           =   3255
         Begin XtremeTaskPanel.TaskPanel TaskPL 
            Height          =   615
            Left            =   720
            TabIndex        =   2
            Top             =   360
            Width           =   1335
            _Version        =   983043
            _ExtentX        =   2355
            _ExtentY        =   1085
            _StockProps     =   64
            ItemLayout      =   2
            HotTrackStyle   =   1
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cBS 
      Left            =   2880
      Top             =   5880
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPL 
      Left            =   720
      Top             =   4560
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSysMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim lngID As Long


Sub msAddAction()
    '创建CommandBars的Action
    
    Dim cbsAction As CommandBarAction
    
    With cBS.Actions
'        cbs.Actions.Add "Id","Caption","TooltipText","DescriptionText","Category"
        
        .Add gID.Sys, "系统", "", "", "Sys"
        .Add gID.SysExit, "退出", "", "", ""
        .Add gID.SysModifyPassword, "修改密码", "", "", ""
        .Add gID.SysReLogin, "重新登陆", "", "", ""
        
        
        
        .Add gID.Help, "帮助", "", "", "Help"
        .Add gID.HelpAbout, "关于", "", "", ""
        .Add gID.HelpDocument, "帮助文档", "", "", ""
        
        .Add gID.Wnd, "窗口", "", "", "Window"
        .Add gID.WndThemeCommandBars, "工具栏主题", "", "", ""
        .Add gID.WndThemeCommandBarsOffice2000, "Office2000", "", "", ""
        .Add gID.WndThemeCommandBarsOffice2003, "Office2003", "", "", ""
        .Add gID.WndThemeCommandBarsOfficeXp, "OfficeXp", "", "", ""
        .Add gID.WndThemeCommandBarsResource, "Resource", "", "", ""
        .Add gID.WndThemeCommandBarsRibbon, "Ribbon", "", "", ""
        .Add gID.WndThemeCommandBarsVS2008, "VS2008", "", "", ""
        .Add gID.WndThemeCommandBarsVS2010, "VS2010", "", "", ""
        .Add gID.WndThemeCommandBarsVS6, "VS6.0", "", "", ""
        .Add gID.WndThemeCommandBarsWhidbey, "Whidbey", "", "", ""
        .Add gID.WndThemeCommandBarsWinXP, "WinXP", "", "", ""

    End With
    

    For Each cbsAction In cBS.Actions
        cbsAction.ToolTipText = cbsAction.Caption
        cbsAction.DescriptionText = cbsAction.ToolTipText
        cbsAction.Key = cbsAction.Category
        cbsAction.Category = cBS.Actions(CLng(cbsAction.Id \ 100) * 100).Category
    Next

    
End Sub

Sub msAddMenu()
    '创建菜单栏
    
    Dim cbsMenuBar As XtremeCommandBars.MenuBar
    Dim cbsMenuMain As CommandBarPopup
    Dim cbsMenuCtrl As CommandBarControl
    
    
    '系统主菜单
    Set cbsMenuBar = cBS.ActiveMenuBar
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Sys, "")
    With cbsMenuMain.CommandBar.Controls
        .Add xtpControlButton, gID.SysModifyPassword, ""
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysReLogin, "")
        cbsMenuCtrl.BeginGroup = True
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysExit, "")
        cbsMenuCtrl.BeginGroup = True
    End With
    
    
    '窗口主菜单
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    
    'CommandBars工具栏主题子菜单
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, gID.WndThemeCommandBars, "")
    With cbsMenuCtrl.CommandBar.Controls
        For lngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            .Add xtpControlButton, lngID, ""
        Next
    End With
    
    
    '帮助主菜单
    Set cbsMenuBar = cBS.ActiveMenuBar
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Help, "")
    With cbsMenuMain.CommandBar.Controls
        .Add xtpControlButton, gID.HelpDocument, ""
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.HelpAbout, "")
        cbsMenuCtrl.BeginGroup = True
    End With
    
    
    
End Sub

Sub msAddToolBar()
    '创建工具栏
    
    Dim cbsBar As CommandBar
    Dim cbsCtr As CommandBarControl
    
    Set cbsBar = cBS.Add(cBS.Actions(gID.WndThemeCommandBars).Caption, xtpBarTop)
    With cbsBar.Controls
        For lngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            Set cbsCtr = .Add(xtpControlButton, lngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    
End Sub

Sub msAddTaskPanelItem()
    '创建导航菜单
    
    Dim paneLeft As XtremeDockingPane.Pane
    
    Set paneLeft = DockingPL.CreatePane(1, 150, 150, DockLeftOf, Nothing)
    paneLeft.Title = "导航菜单"
    
End Sub

Sub msAddStatuBar()
    '创建状态栏
    
    Dim statuBar As XtremeCommandBars.StatusBar
    
    Set statuBar = cBS.StatusBar
    statuBar.Visible = True
    
    statuBar.AddPane 0
'    statuBar.IdleText = "准备"
    statuBar.SetPaneStyle 0, SBPS_STRETCH
    
    statuBar.AddPane 59137  'CapsLock键的状态
    statuBar.AddPane 59138  'NumLK键的状态
    statuBar.AddPane 59139  'ScrLK键的状态
    
End Sub

Private Sub MDIForm_Load()

'    Debug.Print Screen.TwipsPerPixelX, Screen.TwipsPerPixelY    '返回水平与垂直度量的对象的每一像素中的缇数。测试结果：1像素=15缇

    Me.Width = 15360    '设置窗口大小1024*768像素
    Me.Height = 11520
    
    '注意：先往窗体中拖入DockingPanel控件，再拖入CommandBars控件，显示才正常。
    DockingPL.SetCommandBars Me.cBS     '使DockingPanel与CommandBars控件关联起来，子Pane与CommandBarControl控件在位置移动时才能显示正常。
    
    cBS.EnableActions   '启用CommandBars的Actions集合，否则msAddAction过程执行无效。
    cBS.VisualTheme = xtpThemeVisualStudio2008
    cBS.ShowTabWorkspace True
    cBS.TabWorkspace.AllowReorder = True
    cBS.TabWorkspace.AutoTheme = True
    cBS.TabWorkspace.Flags = xtpWorkspaceShowCloseSelectedTab Or xtpWorkspaceShowActiveFiles

    
    DockingPL.Options.AlphaDockingContext = True
    DockingPL.Options.ShowDockingContextStickers = True
    DockingPL.VisualTheme = ThemeWord2007
    
    Call msAddAction
    Call msAddMenu
    Call msAddToolBar
    Call msAddTaskPanelItem
    Call msAddStatuBar
    
    Dim frmNew As Form
    For lngID = 1 To 5
        Set frmNew = New frmSysTest
        frmNew.Show
    Next
    
End Sub

