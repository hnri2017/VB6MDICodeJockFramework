VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.DockingPane.v15.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "Codejock.TaskPanel.v15.3.1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm frmSysMDI 
   BackColor       =   &H8000000C&
   Caption         =   "���������"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.ImageList imgListCommandBars 
      Left            =   3960
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":0000
            Key             =   "cNativeWinXP"
            Object.Tag             =   "820"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":0357
            Key             =   "cOffice2000"
            Object.Tag             =   "811"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":064E
            Key             =   "cOffice2003"
            Object.Tag             =   "812"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":0BE2
            Key             =   "cOfficeXP"
            Object.Tag             =   "813"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":0F0E
            Key             =   "cResource"
            Object.Tag             =   "814"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1475
            Key             =   "cRibbon"
            Object.Tag             =   "815"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":18B6
            Key             =   "cVisualStudio6.0"
            Object.Tag             =   "818"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1BF9
            Key             =   "cVisualStudio2008"
            Object.Tag             =   "816"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":2116
            Key             =   "cVisualStudio2010"
            Object.Tag             =   "817"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":252D
            Key             =   "cWhidbey"
            Object.Tag             =   "819"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":2978
            Key             =   "tListView"
            Object.Tag             =   "841"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":2C36
            Key             =   "tListViewOffice2003"
            Object.Tag             =   "842"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":2F36
            Key             =   "tListViewOfficeXP"
            Object.Tag             =   "843"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":31F3
            Key             =   "tNativeWinXP"
            Object.Tag             =   "844"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":379F
            Key             =   "tNativeWinXPPlain"
            Object.Tag             =   "845"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":3C20
            Key             =   "tOffice2000"
            Object.Tag             =   "846"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":402F
            Key             =   "tOffice2000Plain"
            Object.Tag             =   "847"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":4444
            Key             =   "tOffice2003"
            Object.Tag             =   "848"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":474E
            Key             =   "tOffice2003Plain"
            Object.Tag             =   "849"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":4A51
            Key             =   "tOfficeXPPlain"
            Object.Tag             =   "850"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":4D0D
            Key             =   "tResource"
            Object.Tag             =   "851"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":5007
            Key             =   "tShortcutBarOffice2003"
            Object.Tag             =   "852"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":5327
            Key             =   "tToolbox"
            Object.Tag             =   "853"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":55E4
            Key             =   "tToolboxWhidbey"
            Object.Tag             =   "854"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":59A1
            Key             =   "tVisualStudio2010"
            Object.Tag             =   "855"
         EndProperty
      EndProperty
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
      Begin VB.PictureBox picList 
         Height          =   1455
         Left            =   5160
         ScaleHeight     =   1395
         ScaleWidth      =   3915
         TabIndex        =   3
         Top             =   600
         Width           =   3975
         Begin VB.ListBox listBlank 
            Height          =   960
            Left            =   480
            TabIndex        =   4
            Top             =   120
            Width           =   2895
         End
      End
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
      Left            =   3000
      Top             =   6120
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPN 
      Left            =   2400
      Top             =   6120
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


Dim mLngID As Long  'ѭ������ID
Dim mcbsActions As CommandBarActions    'cBS�ؼ�Actions���ϵ�����
Dim mcbsPopupNav As CommandBar      '���ڵ����˵���������ɱ����е�Popup�˵�
Dim mcbsPopupTab As CommandBar      '���ڶ��ǩ�Ҽ�Popup�˵�
Dim WithEvents mTabWorkspace As TabWorkspace    '���ڶ��ǩ�ؼ�
Attribute mTabWorkspace.VB_VarHelpID = -1


Sub msAddAction()
    '����CommandBars��Action
    
    Dim cbsAction As CommandBarAction
    
    cBS.EnableActions   '����CommandBars��Actions����
    Set mcbsActions = cBS.Actions
    
    With mcbsActions
'        mcbsActions.Add "Id","Caption","TooltipText","DescriptionText","Category"
        
        .Add gID.Sys, "ϵͳ", "", "", "ϵͳ"
        .Add gID.SysExit, "�˳�", "", "", ""
        .Add gID.SysModifyPassword, "�޸�����", "", "", ""
        .Add gID.SysReLogin, "���µ�½", "", "", ""
        
        
        
        .Add gID.Help, "����", "", "", "����"
        .Add gID.HelpAbout, "����", "", "", ""
        .Add gID.HelpDocument, "�����ĵ�", "", "", ""
        
        
        .Add gID.Wnd, "����", "", "", "����"
        
        .Add gID.WndResetLayout, "���ô��ڲ���", "", "", ""
        
        .Add gID.WndThemeCommandBars, "����������", "", "", ""
        .Add gID.WndThemeCommandBarsOffice2000, "Office2000", "", "", ""
        .Add gID.WndThemeCommandBarsOffice2003, "Office2003", "", "", ""
        .Add gID.WndThemeCommandBarsOfficeXp, "OfficeXp", "", "", ""
        .Add gID.WndThemeCommandBarsResource, "Resource", "", "", ""
        .Add gID.WndThemeCommandBarsRibbon, "Ribbon", "", "", ""
        .Add gID.WndThemeCommandBarsVS2008, "VisualStudio2008", "", "", ""
        .Add gID.WndThemeCommandBarsVS2010, "VisualStudio2010", "", "", ""
        .Add gID.WndThemeCommandBarsVS6, "VisualStudio6.0", "", "", ""
        .Add gID.WndThemeCommandBarsWhidbey, "Whidbey", "", "", ""
        .Add gID.WndThemeCommandBarsWinXP, "WinXP", "", "", ""
        
        .Add gID.WndThemeTaskPanel, "�����˵�����", "", "", ""
        .Add gID.WndThemeTaskPanelListView, "ListView", "", "", ""
        .Add gID.WndThemeTaskPanelListViewOffice2003, "ListViewOffice2003", "", "", ""
        .Add gID.WndThemeTaskPanelListViewOfficeXP, "ListViewOfficeXP", "", "", ""
        .Add gID.WndThemeTaskPanelNativeWinXP, "NativeWinXP", "", "", ""
        .Add gID.WndThemeTaskPanelNativeWinXPPlain, "NativeWinXPPlain", "", "", ""
        .Add gID.WndThemeTaskPanelOffice2000, "Office2000", "", "", ""
        .Add gID.WndThemeTaskPanelOffice2000Plain, "Office2000Plain", "", "", ""
        .Add gID.WndThemeTaskPanelOffice2003, "Office2003", "", "", ""
        .Add gID.WndThemeTaskPanelOffice2003Plain, "Office2003Plain", "", "", ""
        .Add gID.WndThemeTaskPanelOfficeXPPlain, "OfficeXPPlain", "", "", ""
        .Add gID.WndThemeTaskPanelResource, "Resource", "", "", ""
        .Add gID.WndThemeTaskPanelShortcutBarOffice2003, "ShortcutBarOffice2003", "", "", ""
        .Add gID.WndThemeTaskPanelToolbox, "Toolbox", "", "", ""
        .Add gID.WndThemeTaskPanelToolboxWhidbey, "ToolboxWhidbey", "", "", ""
        .Add gID.WndThemeTaskPanelVisualStudio2010, "VisualStudio2010", "", "", ""
        
        .Add gID.WndSon, "�Ӵ��ڿ���", "", "", ""
        .Add gID.WndSonCloseAll, "�ر����д���", "", "", ""
        .Add gID.WndSonCloseCurrent, "�رյ�ǰ����", "", "", ""
        .Add gID.WndSonCloseLeft, "�رյ�ǰ��ǩ��ര��", "", "", ""
        .Add gID.WndSonCloseOther, "�ر���������", "", "", ""
        .Add gID.WndSonCloseRight, "�رյ�ǰ��ǩ�Ҳര��", "", "", ""
        .Add gID.WndSonVbArrangeIcons, "������С��ͼ��", "", "", ""
        .Add gID.WndSonVbCascade, "���", "", "", ""
        .Add gID.WndSonVbTileHorizontal, "ˮƽƽ��", "", "", ""
        .Add gID.WndSonVbTileVertical, "��ֱƽ��", "", "", ""
        
        
        .Add gID.Other, "����", "", "", "����"
        
        .Add gID.OtherPane, "�������", "", "", ""
        .Add gID.OtherPaneMenuPopu, "PaneCaptionMenu", "", "", ""
        .Add gID.OtherPaneMenuPopuAutoFold, "�Զ���£", "", "", ""
        .Add gID.OtherPaneMenuPopuExpand, "ȫ��չ��", "", "", ""
        .Add gID.OtherPaneMenuPopuFold, "ȫ����£", "", "", ""
        .Add gID.OtherPaneMenuTitle, "�����˵�", "", "", ""
        
        .Add gID.OtherTabWorkspacePopup, "���ǩ�Ҽ��˵�", "", "", ""
        
        .Add gID.StatusBarPane, "״̬��", "", "", ""
        .Add gID.StatusBarPaneProgress, "������", "", "", ""
        .Add gID.StatusBarPaneProgressText, "���Ȱٷֱ�", "", "", ""
        .Add gID.StatusBarPaneTime, "ϵͳʱ��", "", "", ""
        .Add gID.StatusBarPaneUserInfo, "��ǰ�û�", "", "", ""
        
        
    End With
    

    
    For Each cbsAction In mcbsActions
        With cbsAction
            If .Id < 2000 Then
                .ToolTipText = .Caption
                .DescriptionText = .ToolTipText
                .Key = .Category
                .Category = mcbsActions((.Id \ 100) * 100).Category
            End If
        End With
    Next

End Sub

Sub msAddDesignerControls()
    '
    
    Dim cbsControls As CommandBarControls
    Dim cbsAction As CommandBarAction

    Set cbsControls = cBS.DesignerControls
    For Each cbsAction In mcbsActions
        If cbsAction.Id < 2000 Then
            cbsControls.Add xtpControlButton, cbsAction.Id, ""
        End If
    Next
    
End Sub

Sub msAddDockingPane()
    '�����������
    
    Dim paneLeft As XtremeDockingPane.Pane
    Dim paneList As XtremeDockingPane.Pane
    
    '���������˵���������һ��DockingPane
    Set paneLeft = DockingPN.CreatePane(gID.OtherPaneIDFirst, 240, 240, DockLeftOf, Nothing)
    paneLeft.Title = mcbsActions(gID.OtherPaneMenuTitle).Caption
    paneLeft.TitleToolTip = paneLeft.Title & mcbsActions(gID.OtherPane).Caption
    paneLeft.Handle = picTaskPL.hWnd    '���������TaskPanel������PictureBox�ؼ��ҿ��ڸ������PanelLeft��
    paneLeft.Options = PaneHasMenuButton    '��ʾPopu����
    
        
'    '�ڶ���DockingPane
'    Set paneList = DockingPN.CreatePane(gID.OtherPaneIDSecond, 240, 240, DockLeftOf, Nothing)
'    paneList.Title = "       "
'    paneList.Handle = picList.hWnd
'    paneList.AttachTo paneLeft  '��������һ��Pane��
'    paneLeft.Selected = True    '��ʾ��һ��Pane
 
End Sub

Sub msAddKeyBindings()
    '������ݼ�
    
    With cBS.KeyBindings
        .AddShortcut gID.HelpDocument, "F1"
'        .Add 0, &H70, gID.HelpDocument
    End With
    
End Sub

Sub msAddMenu()
    '�����˵���
    
    Dim cbsMenuBar As XtremeCommandBars.MenuBar
    Dim cbsMenuMain As CommandBarPopup
    Dim cbsMenuCtrl As CommandBarControl
    
    
    Set cbsMenuBar = cBS.ActiveMenuBar
    cbsMenuBar.ShowGripper = False  '����ʾ���϶����Ǹ������
    cbsMenuBar.EnableDocking xtpFlagStretched     '�˵�����ռһ���Ҳ��������϶�
    
    'ϵͳ���˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Sys, "")
    With cbsMenuMain.CommandBar.Controls
        .Add xtpControlButton, gID.SysModifyPassword, ""
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysReLogin, "")
        cbsMenuCtrl.BeginGroup = True
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysExit, "")
        cbsMenuCtrl.BeginGroup = True
    End With
    
    
    '�������˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    
    '���ò���
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, gID.WndResetLayout, "")
    cbsMenuCtrl.BeginGroup = True
    
    '����ID35001
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, XTP_ID_CUSTOMIZE, "�Զ��幤������")
    cbsMenuCtrl.BeginGroup = True
    
    '����ID59392
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, 0, "�������б�")
    cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, ""
    
    'CommandBars�����������Ӳ˵�
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, gID.WndThemeCommandBars, "")
    cbsMenuCtrl.BeginGroup = True
    With cbsMenuCtrl.CommandBar.Controls
        For mLngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            .Add xtpControlButton, mLngID, ""
        Next
    End With
    
    'TaskPanel�����˵������Ӳ˵�
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, gID.WndThemeTaskPanel, "")
    With cbsMenuCtrl.CommandBar.Controls
        For mLngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
            .Add xtpControlButton, mLngID, ""
        Next
    End With
    
    '�Ӵ��ڿ���
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, gID.WndSon, "")
    cbsMenuCtrl.BeginGroup = True
    With cbsMenuCtrl.CommandBar.Controls
        For mLngID = gID.WndSonCloseAll To gID.WndSonVbTileVertical
            .Add xtpControlButton, mLngID, ""
            If mLngID = gID.WndSonVbArrangeIcons Then .Find(, mLngID).BeginGroup = True
        Next
    End With
  
    
    '����ID35000
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, 0, "�Ѵ򿪴����б�")
    cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_WINDOWLIST, ""
    
    
    
    '�������˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Help, "")
    With cbsMenuMain.CommandBar.Controls
        .Add xtpControlButton, gID.HelpDocument, ""
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.HelpAbout, "")
        cbsMenuCtrl.BeginGroup = True
    End With
    
    
    
End Sub

Sub msAddPopupMenu()
    '����Popup�˵�,���Ҽ�������ʾ
    
    '����Pane�����е�Popup�˵�
    Set mcbsPopupNav = cBS.Add(mcbsActions(gID.OtherPaneMenuPopu).Caption, xtpBarPopup)
    With mcbsPopupNav.Controls
        .Add xtpControlButton, gID.OtherPaneMenuPopuAutoFold, ""
        .Add xtpControlButton, gID.OtherPaneMenuPopuExpand, ""
        .Add xtpControlButton, gID.OtherPaneMenuPopuFold, ""
    End With
    
    '�����Ӵ��ڶ��ǩ�ؼ����Ҽ��˵�
    Set mcbsPopupTab = cBS.Add(mcbsActions(gID.OtherTabWorkspacePopup).Caption, xtpBarPopup)
    mcbsPopupTab.BarID = gID.OtherTabWorkspacePopup
    With mcbsPopupTab.Controls
        For mLngID = gID.WndSonCloseAll To gID.WndSonCloseRight
            .Add xtpControlButton, mLngID, ""
        Next
    End With
    
End Sub

Sub msAddStatuBar()
    '����״̬��
    
    Dim statuBar As XtremeCommandBars.StatusBar
    
    Set statuBar = cBS.StatusBar
    With statuBar
        .AddPane 0      'ϵͳPane����ʾCommandBarActions��Description
        .SetPaneStyle 0, SBPS_STRETCH
        
        .AddPane gID.StatusBarPaneUserInfo
        .FindPane(gID.StatusBarPaneUserInfo).Caption = mcbsActions(gID.StatusBarPaneUserInfo).Caption
        .FindPane(gID.StatusBarPaneUserInfo).Text = "С��"
        
        .AddProgressPane gID.StatusBarPaneProgress
        .SetPaneText gID.StatusBarPaneProgress, mcbsActions(gID.StatusBarPaneProgress).Caption
        
        .AddPane gID.StatusBarPaneProgressText
        .FindPane(gID.StatusBarPaneProgressText).Caption = mcbsActions(gID.StatusBarPaneProgressText).Caption
        .FindPane(gID.StatusBarPaneProgressText).Width = 40
        
        .AddPane 59137  'CapsLock����״̬
        .AddPane 59138  'NumLK����״̬
        .AddPane 59139  'ScrLK����״̬
        
        .Visible = True
        .EnableCustomization True
        
    End With
    
End Sub

Sub msAddTaskPanelItem()
    '���������˵�
    'ע�⣺����ĵ����˵����ǲ˵�������һ����ʾ��ʽ
    
    Dim taskGroup As TaskPanelGroup
    Dim taskItem As TaskPanelGroupItem
    
    
    'ϵͳ
    Set taskGroup = TaskPL.Groups.Add(gID.Sys, mcbsActions(gID.Sys).Caption)
    With taskGroup.Items
        .Add gID.SysModifyPassword, mcbsActions(gID.SysModifyPassword).Caption, xtpTaskItemTypeLink
        .Add gID.SysReLogin, mcbsActions(gID.SysReLogin).Caption, xtpTaskItemTypeLink
        .Add gID.SysExit, mcbsActions(gID.SysExit).Caption, xtpTaskItemTypeLink
    End With
    
    
    '����
    Set taskGroup = TaskPL.Groups.Add(gID.Wnd, mcbsActions(gID.Wnd).Caption)
    
    '����
    Set taskItem = taskGroup.Items.Add(gID.WndResetLayout, mcbsActions(gID.WndResetLayout).Caption, xtpTaskItemTypeLink)
    
    '����������
    Set taskItem = taskGroup.Items.Add(gID.WndThemeCommandBars, mcbsActions(gID.WndThemeCommandBars).Caption, xtpTaskItemTypeText)
    taskItem.Bold = True
    For mLngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        taskGroup.Items.Add mLngID, mcbsActions(mLngID).Caption, xtpTaskItemTypeLink
    Next
    
    '�����˵�����
    Set taskItem = taskGroup.Items.Add(gID.WndThemeTaskPanel, mcbsActions(gID.WndThemeTaskPanel).Caption, xtpTaskItemTypeText)
    taskItem.Bold = True
    For mLngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        taskGroup.Items.Add mLngID, mcbsActions(mLngID).Caption, xtpTaskItemTypeLink
    Next
    
    
    '����
    Set taskGroup = TaskPL.Groups.Add(gID.Help, mcbsActions(gID.Help).Caption)
    taskGroup.Items.Add gID.HelpDocument, mcbsActions(gID.HelpDocument).Caption, xtpTaskItemTypeLink
    taskGroup.Items.Add gID.HelpAbout, mcbsActions(gID.HelpAbout).Caption, xtpTaskItemTypeLink
    
    
End Sub

Sub msAddToolBar()
    '����������
    
    Dim cbsBar As CommandBar
    Dim cbsCtr As CommandBarControl
    
    '����������
    Set cbsBar = cBS.Add(mcbsActions(gID.WndThemeCommandBars).Caption, xtpBarTop)
    With cbsBar.Controls
        For mLngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            Set cbsCtr = .Add(xtpControlButton, mLngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    '�����˵�����
    Set cbsBar = cBS.Add(mcbsActions(gID.WndThemeTaskPanel).Caption, xtpBarTop)
    With cbsBar.Controls
        For mLngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
            Set cbsCtr = .Add(xtpControlButton, mLngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    

End Sub

Sub msResetLayout()
    '���ô��ڲ��֣�CommandBars��Dockingpane�ؼ�����
    
    Dim cBar As CommandBar
    Dim L As Long, T As Long, R As Long, B As Long

    For Each cBar In cBS
        cBar.Reset
        cBar.Visible = True
    Next
    
    For mLngID = 2 To cBS.Count
        cBS.GetClientRect L, T, R, B
        cBS.DockToolBar cBS(mLngID), 0, B, xtpBarTop
    Next

    Dim pnRe As XtremeDockingPane.Pane
    For Each pnRe In DockingPN
        pnRe.Closed = False
        pnRe.Hidden = False
    Next

End Sub

Sub msThemeCommandBar(ByVal CID As Long)
    'CommandBars�������
    
    Select Case CID
        Case gID.WndThemeCommandBarsOffice2000
            cBS.VisualTheme = xtpThemeOffice2000
        Case gID.WndThemeCommandBarsOffice2003
            cBS.VisualTheme = xtpThemeOffice2003
        Case gID.WndThemeCommandBarsOfficeXp
            cBS.VisualTheme = xtpThemeOfficeXP
        Case gID.WndThemeCommandBarsResource
            cBS.VisualTheme = xtpThemeResource
        Case gID.WndThemeCommandBarsRibbon
            cBS.VisualTheme = xtpThemeRibbon
        Case gID.WndThemeCommandBarsVS2008
            cBS.VisualTheme = xtpThemeVisualStudio2008
        Case gID.WndThemeCommandBarsVS2010
            cBS.VisualTheme = xtpThemeVisualStudio2010
        Case gID.WndThemeCommandBarsVS6
            cBS.VisualTheme = xtpThemeVisualStudio6
        Case gID.WndThemeCommandBarsWhidbey
            cBS.VisualTheme = xtpThemeWhidbey
        Case gID.WndThemeCommandBarsWinXP
            cBS.VisualTheme = xtpThemeNativeWinXP
    End Select
    
    For mLngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        mcbsActions(mLngID).Checked = False
    Next
    mcbsActions(CID).Checked = True
    
End Sub

Sub msThemeTaskPanel(ByVal TID As Long)
    'Taskpanel�������
    
    Select Case TID
        Case gID.WndThemeTaskPanelListView
            TaskPL.VisualTheme = xtpTaskPanelThemeListView
        Case gID.WndThemeTaskPanelListViewOffice2003
            TaskPL.VisualTheme = xtpTaskPanelThemeListViewOffice2003
        Case gID.WndThemeTaskPanelListViewOfficeXP
            TaskPL.VisualTheme = xtpTaskPanelThemeListViewOfficeXP
        Case gID.WndThemeTaskPanelNativeWinXP
            TaskPL.VisualTheme = xtpTaskPanelThemeNativeWinXP
        Case gID.WndThemeTaskPanelNativeWinXPPlain
            TaskPL.VisualTheme = xtpTaskPanelThemeNativeWinXPPlain
        Case gID.WndThemeTaskPanelOffice2000
            TaskPL.VisualTheme = xtpTaskPanelThemeOffice2000
        Case gID.WndThemeTaskPanelOffice2000Plain
            TaskPL.VisualTheme = xtpTaskPanelThemeOffice2000Plain
        Case gID.WndThemeTaskPanelOffice2003
            TaskPL.VisualTheme = xtpTaskPanelThemeOffice2003
        Case gID.WndThemeTaskPanelOffice2003Plain
            TaskPL.VisualTheme = xtpTaskPanelThemeOffice2003Plain
        Case gID.WndThemeTaskPanelOfficeXPPlain
            TaskPL.VisualTheme = xtpTaskPanelThemeOfficeXPPlain
        Case gID.WndThemeTaskPanelResource
            TaskPL.VisualTheme = xtpTaskPanelThemeResource
        Case gID.WndThemeTaskPanelShortcutBarOffice2003
            TaskPL.VisualTheme = xtpTaskPanelThemeShortcutBarOffice2003
        Case gID.WndThemeTaskPanelToolbox
            TaskPL.VisualTheme = xtpTaskPanelThemeToolbox
        Case gID.WndThemeTaskPanelToolboxWhidbey
            TaskPL.VisualTheme = xtpTaskPanelThemeToolboxWhidbey
        Case gID.WndThemeTaskPanelVisualStudio2010
            TaskPL.VisualTheme = xtpTaskPanelThemeVisualStudio2010
    End Select
    
    For mLngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        mcbsActions(mLngID).Checked = False
    Next
    mcbsActions(TID).Checked = True
    
End Sub

Sub msWindowControl(ByVal WID As Long)
    '�Ӵ��ڿ���
    
    Dim frmTag As Form
    Dim C As Long
    Dim itemCur As XtremeCommandBars.TabControlItem
    
    With gID
        Select Case WID
            Case .WndSonCloseAll
                For Each frmTag In Forms
                    If frmTag.Name <> gMDI.Name Then Unload frmTag
                Next
            Case .WndSonCloseCurrent
                If Not ActiveForm Is Nothing Then Unload ActiveForm
            Case .WndSonCloseLeft
                If Forms.Count > 2 Then
                    Set itemCur = mTabWorkspace.Selected
                    itemCur.Tag = "c"   '��ǵ�ǰ���ڣ���ΪIndexֵ�ڴ��������仯ʱ��仯��������ΪΨһ�ж�����
                    For C = 0 To mTabWorkspace.ItemCount - 1
                        If mTabWorkspace.Item(0).Tag = itemCur.Tag Then
                            itemCur.Tag = ""    '�ǵ���ա�Tag����Ĭ��ֵ���ǿ��ַ���
                            Exit For
                        Else
                            mTabWorkspace.Item(0).Selected = True   '����Ҫɾ���Ĵ���
                            Unload ActiveForm
                        End If
                    Next
                End If
            Case .WndSonCloseOther
                If Forms.Count > 1 Then
                    For Each frmTag In Forms
                        If frmTag.Name <> gMDI.Name Then
                            If Not (frmTag.Name = ActiveForm.Name And frmTag.Caption = ActiveForm.Caption) Then
                                Unload frmTag
                            End If
                        End If
                    Next
                End If
            Case .WndSonCloseRight
                If Forms.Count > 2 Then
                    Set itemCur = mTabWorkspace.Selected
                    itemCur.Tag = "c"
                    For C = mTabWorkspace.ItemCount - 1 To 0 Step -1
                        If mTabWorkspace.Item(C).Tag = itemCur.Tag Then
                            itemCur.Tag = ""
                            Exit For
                        Else
                            mTabWorkspace.Item(C).Selected = True
                            Unload ActiveForm
                        End If
                    Next
                End If
            Case .WndSonVbCascade
                Me.Arrange vbCascade
            Case .WndSonVbArrangeIcons
                Me.Arrange vbArrangeIcons
            Case .WndSonVbTileHorizontal
                Me.Arrange vbTileHorizontal
            Case .WndSonVbTileVertical
                Me.Arrange vbTileVertical
        End Select
    End With
    
End Sub

Sub msCommandBarPopu(ByVal PID As Long)
    'Popu�˵���Ӧ
    
    Dim taskGroup As TaskPanelGroup
            
    Select Case PID
        Case gID.OtherPaneMenuPopuAutoFold
            mcbsActions(PID).Checked = Not mcbsActions(PID).Checked
        Case gID.OtherPaneMenuPopuExpand
            For Each taskGroup In TaskPL.Groups
                taskGroup.Expanded = True
            Next
        Case gID.OtherPaneMenuPopuFold
            For Each taskGroup In TaskPL.Groups
                taskGroup.Expanded = False
            Next
    End Select
    
End Sub

Sub msLeftClick(ByVal CID As Long)
    'CommandBar��TaskPanelGroupItem����������Ӧ��������
    
    With gID
        Select Case CID
            Case .WndThemeCommandBarsOffice2000 To .WndThemeCommandBarsWinXP
                Call msThemeCommandBar(CID)
            Case .OtherPaneMenuPopuAutoFold To .OtherPaneMenuPopuFold
                Call msCommandBarPopu(CID)
            Case .WndThemeTaskPanelListView To .WndThemeTaskPanelVisualStudio2010
                Call msThemeTaskPanel(CID)
            Case .WndSonCloseAll To .WndSonVbTileVertical
                Call msWindowControl(CID)
            
            Case .WndResetLayout
                Call msResetLayout
            Case Else
                MsgBox "��" & mcbsActions(CID).Caption & "������δ���壡", vbExclamation, "�����"
        End Select
    End With
    
End Sub

Private Sub cBS_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '������¼�

    Call msLeftClick(Control.Id)
    
End Sub


Private Sub DockingPn_PanePopupMenu(ByVal Pane As XtremeDockingPane.IPane, ByVal x As Long, ByVal y As Long, Handled As Boolean)
    '�����˵������е�Popu�˵�����

    If Pane.Id = gID.OtherPaneIDFirst Then
        mcbsPopupNav.ShowPopup , x * 15, y * 15     'ֻ֪������15��λ�ò��ԣ�����x��y�ĵ�λ�����أ�������Ҫ��羡�
    End If
    
End Sub

Private Sub MDIForm_Load()

'    Debug.Print Screen.TwipsPerPixelX, Screen.TwipsPerPixelY    '����ˮƽ�봹ֱ�����Ķ����ÿһ�����е���������Խ����1����=15�
'    Me.Width = 15360    '���ô��ڴ�С1024*768����
'    Me.Height = 11520
    
    CommandBarsGlobalSettings.App = App
    
    Call msAddAction        '����Actions����
    Call msAddMenu          '�����˵���
    Call msAddToolBar       '����������
    Call msAddDockingPane   '�����������
    Call msAddPopupMenu     '����Popup�˵�
    Call msAddTaskPanelItem '���������˵�
    Call msAddStatuBar      '����״̬��
'    Call msAddKeyBindings   '��ӿ�ݼ�,�ŵ�LoadCommandBars�������������Ч
    Call msAddDesignerControls  'CommandBars�Զ���Ի�����ʹ�õ�
    

    cBS.AddImageList imgListCommandBars '���ͼ��
    cBS.EnableCustomization True        '�����Զ��壬��������÷�������CommandBars�趨֮��
    
    Set mTabWorkspace = cBS.ShowTabWorkspace(True)    '�����ڶ��ǩ��ʾ
'    mTabWorkspace.AllowReorder = False
    mTabWorkspace.Flags = xtpWorkspaceShowCloseSelectedTab Or xtpWorkspaceShowActiveFiles
    
    
    'ע�⣺��������������DockingPanel�ؼ���������CommandBars�ؼ��������Ҽ�CommandBars�ؼ���ѡ���Ƶ�����,��ʾ��������
    'ʹDockingPanel��CommandBars�ؼ�������������Pane��CommandBar�ؼ���λ���ƶ�����С�仯ʱ������ʾ������
    DockingPN.SetCommandBars Me.cBS
    
    DockingPN.Options.AlphaDockingContext = True    '��ʾDockingλ��ָ���ǩ��Ӱ��
    DockingPN.Options.ShowDockingContextStickers = True
    DockingPN.VisualTheme = ThemeWord2007
    
    Dim frmNew As Form
    For mLngID = 2 To 15
        Set frmNew = New frmSysTest
        frmNew.Caption = "Form" & mLngID
        frmNew.Command1.Caption = frmNew.Caption & "cmd1"
        frmNew.Show
    Next
    
    
    'ע����б����õļ�������ֵ��ʼ��
    With gID
        .OtherSaveRegistryKey = Me.Name
        .OtherSaveAppName = Me.Name & "Layout"
        .OtherSaveCommandBarsSection = "CommandBarsLayout"
        .OtherSaveDockingPaneSection = "DockingPaneLayout"
    End With

    '����λ��
    Dim WS As Long, L As Long, T As Long, W As Long, H As Long
    WS = Val(GetSetting(Me.Name, "Settings", "WindowState", 2))
    If WS = 2 Then
        Me.WindowState = 2  '���
    Else
        Me.WindowState = 0
        L = Val(GetSetting(Me.Name, gID.OtherSaveSettings, "Left", 0))
        T = Val(GetSetting(Me.Name, gID.OtherSaveSettings, "Top", 0))
        W = Val(GetSetting(Me.Name, gID.OtherSaveSettings, "Width", gID.OtherSaveWidth))
        H = Val(GetSetting(Me.Name, gID.OtherSaveSettings, "Height", gID.OtherSaveHeight))
        If Val(L) < 0 Then L = 0
        If Val(T) < 0 Then T = 0
        If Val(W) < gID.OtherSaveWidth Then W = gID.OtherSaveWidth
        If Val(H) < gID.OtherSaveHeight Then H = gID.OtherSaveHeight
        Me.Move L, T, W, H
    End If

'    CommandBars��������
    cBS.LoadCommandBars gID.OtherSaveRegistryKey, gID.OtherSaveAppName, gID.OtherSaveCommandBarsSection
    Call msAddKeyBindings   '��ӿ�ݼ�
    
    'CommandBars��������
    Call msThemeCommandBar(Val(GetSetting(Me.Name, gID.OtherSaveSettings, "ThemeCommandBas", gID.WndThemeCommandBarsVS2008)))
    
    ''TaskPanel��Popu����
    mcbsActions(gID.OtherPaneMenuPopuAutoFold).Checked = Val(GetSetting(Me.Name, gID.OtherSaveSettings, "AutoFold", 1))
    
'''    'DockingPaneλ��,�ݲ�֪��ô��
'''    DockingPN.LoadState gID.OtherSaveRegistryKey, gID.OtherSaveAppName, gID.OtherSaveDockingPaneSection

    'TaskPanel����������  �ϴε�������˵�λ��
    Call msThemeTaskPanel(Val(GetSetting(Me.Name, gID.OtherSaveSettings, "ThemeTaskPanel", gID.WndThemeTaskPanelNativeWinXP)))
    
    'TaskPanel�ϵ����˵�չ������£����
    Dim taskGroup As TaskPanelGroup
    For Each taskGroup In TaskPL.Groups
        taskGroup.Expanded = Val(GetSetting(Me.Name, gID.OtherSaveSettings, "TaskPL" & taskGroup.Id, 0))
    Next

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    '����λ�ñ���
    Dim L As Long, T As Long, W As Long, H As Long
    If Me.WindowState = 2 Then
        SaveSetting Me.Name, gID.OtherSaveSettings, "WindowState", 2
    Else
        SaveSetting Me.Name, gID.OtherSaveSettings, "WindowState", 0
        L = Me.Left
        T = Me.Top
        W = Me.Width
        H = Me.Height
        If L < 0 Then L = 0
        If T < 0 Then T = 0
        If W < gID.OtherSaveWidth Then W = gID.OtherSaveWidth
        If H < gID.OtherSaveHeight Then H = gID.OtherSaveHeight
        SaveSetting Me.Name, gID.OtherSaveSettings, "Left", L
        SaveSetting Me.Name, gID.OtherSaveSettings, "Top", T
        SaveSetting Me.Name, gID.OtherSaveSettings, "Width", W
        SaveSetting Me.Name, gID.OtherSaveSettings, "Height", H
    End If
    
    'CommandBars���ֱ���
    cBS.SaveCommandBars gID.OtherSaveRegistryKey, gID.OtherSaveAppName, gID.OtherSaveCommandBarsSection
    
    'CommandBas���Ᵽ��
    Dim lngSaveID As Long
    lngSaveID = gID.WndThemeCommandBarsVS2008
    For mLngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        If mcbsActions(mLngID).Checked Then
            lngSaveID = mLngID
            Exit For
        End If
    Next
    SaveSetting Me.Name, gID.OtherSaveSettings, "ThemeCommandBas", lngSaveID
    
    'Taskpanels��Popu����
    lngSaveID = 0
    If mcbsActions(gID.OtherPaneMenuPopuAutoFold).Checked Then lngSaveID = 1
    SaveSetting Me.Name, gID.OtherSaveSettings, "AutoFold", lngSaveID
    
'''    'DockingPaneλ�ñ���
'''    DockingPN.SaveState gID.OtherSaveRegistryKey, gID.OtherSaveAppName, gID.OtherSaveDockingPaneSection
    
    'TaskPanel��Popu
    lngSaveID = gID.WndThemeTaskPanelNativeWinXP
    For mLngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        If mcbsActions(mLngID).Checked Then
            lngSaveID = mLngID
            Exit For
        End If
    Next
    SaveSetting Me.Name, gID.OtherSaveSettings, "ThemeTaskPanel", lngSaveID
    
    'TaskPanel�ϵ����˵�չ������£����
    Dim taskGroup As TaskPanelGroup
    For Each taskGroup In TaskPL.Groups
        lngSaveID = IIf(taskGroup.Expanded, 1, 0)
        SaveSetting Me.Name, gID.OtherSaveSettings, "TaskPL" & taskGroup.Id, lngSaveID
    Next
    
End Sub

Private Sub mTabWorkspace_RClick(ByVal Item As XtremeCommandBars.ITabControlItem)
    '�Ҽ��˵��ĵ���
    
    If Not Item Is Nothing Then
        Item.Selected = True
        mTabWorkspace.Refresh
        mcbsPopupTab.ShowPopup
    End If
End Sub

Private Sub picList_Resize()
    
    listBlank.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
    
End Sub

Private Sub picTaskPL_Resize()
    '��������С��ҿ��ڸ�������ϵ�PictureBox�ؼ��Ĵ�С�仯���仯
    
    TaskPL.Move 0, 0, picTaskPL.ScaleWidth, picTaskPL.ScaleHeight
    
End Sub

Private Sub TaskPL_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
    '�����˵������¼���
    'ע�⣺�򵼺��˵��Ǹ��ƵĲ˵���������Item��IDֵ��cBS�ؼ���IDֵһ�¡�
    
    Dim taskGroup As TaskPanelGroup
    
    '�Զ���£
    If mcbsActions(gID.OtherPaneMenuPopuAutoFold).Checked Then
        For Each taskGroup In TaskPL.Groups
            If taskGroup.Id <> Item.Group.Id Then taskGroup.Expanded = False
        Next
    End If
    
    Call msLeftClick(Item.Id)
    
End Sub
