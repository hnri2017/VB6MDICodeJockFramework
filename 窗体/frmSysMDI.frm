VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "Codejock.SkinFramework.v15.3.1.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.DockingPane.v15.3.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#15.3#0"; "Codejock.CommandBars.v15.3.1.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "Codejock.TaskPanel.v15.3.1.ocx"
Begin VB.MDIForm frmSysMDI 
   BackColor       =   &H8000000C&
   Caption         =   "���������"
   ClientHeight    =   4590
   ClientLeft      =   2880
   ClientTop       =   2415
   ClientWidth     =   9510
   Icon            =   "frmSysMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.ImageList imgListCommandBars 
      Left            =   4200
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   66
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":068A
            Key             =   "cNativeWinXP"
            Object.Tag             =   "820"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":09E1
            Key             =   "cOffice2000"
            Object.Tag             =   "811"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":0CD8
            Key             =   "cOffice2003"
            Object.Tag             =   "812"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":126C
            Key             =   "cOfficeXP"
            Object.Tag             =   "813"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1598
            Key             =   "cResource"
            Object.Tag             =   "814"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1AFF
            Key             =   "cRibbon"
            Object.Tag             =   "815"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1F40
            Key             =   "cVisualStudio6.0"
            Object.Tag             =   "818"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":2283
            Key             =   "cVisualStudio2008"
            Object.Tag             =   "816"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":27A0
            Key             =   "cVisualStudio2010"
            Object.Tag             =   "817"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":2BB7
            Key             =   "cWhidbey"
            Object.Tag             =   "819"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":3002
            Key             =   "tListView"
            Object.Tag             =   "841"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":32C0
            Key             =   "tListViewOffice2003"
            Object.Tag             =   "842"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":35C0
            Key             =   "tListViewOfficeXP"
            Object.Tag             =   "843"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":387D
            Key             =   "tNativeWinXP"
            Object.Tag             =   "844"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":3E29
            Key             =   "tNativeWinXPPlain"
            Object.Tag             =   "845"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":42AA
            Key             =   "tOffice2000"
            Object.Tag             =   "846"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":46B9
            Key             =   "tOffice2000Plain"
            Object.Tag             =   "847"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":4ACE
            Key             =   "tOffice2003"
            Object.Tag             =   "848"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":4DD8
            Key             =   "tOffice2003Plain"
            Object.Tag             =   "849"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":50DB
            Key             =   "tOfficeXPPlain"
            Object.Tag             =   "850"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":5397
            Key             =   "tResource"
            Object.Tag             =   "851"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":5691
            Key             =   "tShortcutBarOffice2003"
            Object.Tag             =   "852"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":59B1
            Key             =   "tToolbox"
            Object.Tag             =   "853"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":5C6E
            Key             =   "tToolboxWhidbey"
            Object.Tag             =   "854"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":602B
            Key             =   "tVisualStudio2010"
            Object.Tag             =   "855"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":63F3
            Key             =   "sCodejock"
            Object.Tag             =   "871"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":7445
            Key             =   "sOffice2007"
            Object.Tag             =   "872"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":8497
            Key             =   "sOffice2010"
            Object.Tag             =   "873"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":94E9
            Key             =   "sOrangina"
            Object.Tag             =   "878"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":A53B
            Key             =   "sVista"
            Object.Tag             =   "874"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":B58D
            Key             =   "sWinXPLuna"
            Object.Tag             =   "875"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":C5DF
            Key             =   "sWinXPRoyale"
            Object.Tag             =   "876"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":D631
            Key             =   "sZune"
            Object.Tag             =   "877"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":E683
            Key             =   ""
            Object.Tag             =   "901"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":ED1D
            Key             =   "SysWord"
            Object.Tag             =   "122"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":F9F7
            Key             =   "SysText"
            Object.Tag             =   "121"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":106D1
            Key             =   "SysExcel"
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":113AB
            Key             =   "SysSearch"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":114BD
            Key             =   "SysPageSet"
            Object.Tag             =   "123"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":11C0F
            Key             =   "SysPreview"
            Object.Tag             =   "125"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":12861
            Key             =   "SysPrint"
            Object.Tag             =   "124"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":134B3
            Key             =   "SysGo"
            Object.Tag             =   "116"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1418D
            Key             =   "SysExit"
            Object.Tag             =   "101"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":14E67
            Key             =   "SysRelogin"
            Object.Tag             =   "103"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":15B41
            Key             =   "SysCompany"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":16793
            Key             =   "SysDepartment"
            Object.Tag             =   "104"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":173E5
            Key             =   "threemen"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":18037
            Key             =   "SysUser"
            Object.Tag             =   "105"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":18C89
            Key             =   "man"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":198DB
            Key             =   "woman"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1A52D
            Key             =   "SysPassword"
            Object.Tag             =   "102"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1B17F
            Key             =   ""
            Object.Tag             =   "902"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1B4D1
            Key             =   "themes"
            Object.Tag             =   "801"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1C123
            Key             =   "SelectedMen"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1CD75
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1CE7F
            Key             =   "SysLog"
            Object.Tag             =   "106"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1DAD1
            Key             =   "SysRole"
            Object.Tag             =   "107"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1E723
            Key             =   "RoleSelect"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1F375
            Key             =   "SysFunc"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":1FFC7
            Key             =   "FuncHead"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":20C19
            Key             =   "FuncSelect"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":2186B
            Key             =   "FuncControl"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":224BD
            Key             =   "FuncButton"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":22BCF
            Key             =   "FuncForm"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":23821
            Key             =   "FuncMainMenu"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSysMDI.frx":25373
            Key             =   "themeSet"
            Object.Tag             =   "802"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picHide 
      Align           =   1  'Align Top
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   9450
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9510
      Begin VSPrinter8LibCtl.VSPrinter VPMain 
         Height          =   495
         Left            =   3960
         TabIndex        =   5
         Top             =   1080
         Width           =   855
         _cx             =   1508
         _cy             =   873
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoRTF         =   -1  'True
         Preview         =   -1  'True
         DefaultDevice   =   0   'False
         PhysicalPage    =   -1  'True
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         MarginHeader    =   0
         MarginFooter    =   0
         IndentLeft      =   0
         IndentRight     =   0
         IndentFirst     =   0
         IndentTab       =   720
         SpaceBefore     =   0
         SpaceAfter      =   0
         LineSpacing     =   100
         Columns         =   1
         ColumnSpacing   =   180
         ShowGuides      =   2
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   30
         SmallChangeVert =   30
         Track           =   0   'False
         ProportionalBars=   -1  'True
         Zoom            =   -2.6448362720403
         ZoomMode        =   3
         ZoomMax         =   400
         ZoomMin         =   10
         ZoomStep        =   25
         EmptyColor      =   -2147483636
         TextColor       =   0
         HdrColor        =   0
         BrushColor      =   0
         BrushStyle      =   0
         PenColor        =   0
         PenStyle        =   0
         PenWidth        =   0
         PageBorder      =   0
         Header          =   ""
         Footer          =   ""
         TableSep        =   "|;"
         TableBorder     =   7
         TablePen        =   0
         TablePenLR      =   0
         TablePenTB      =   0
         NavBar          =   3
         NavBarColor     =   -2147483633
         ExportFormat    =   0
         URL             =   ""
         Navigation      =   3
         NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
         AutoLinkNavigate=   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
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
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   600
      Top             =   3720
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSysMDI.frx":25FC5
   End
   Begin XtremeSkinFramework.SkinFramework skinFW 
      Left            =   3720
      Top             =   3600
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars cBS 
      Left            =   3240
      Top             =   3600
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPN 
      Left            =   2640
      Top             =   3600
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


Dim mlngID As Long  'ѭ������ID
Dim mcbsActions As CommandBarActions    'cBS�ؼ�Actions���ϵ�����
Dim mcbsPopupNav As CommandBar      '���ڵ����˵���������ɱ����е�Popup�˵�
Dim mcbsPopupTab As CommandBar      '���ڶ��ǩ�Ҽ�Popup�˵�
Dim WithEvents mTabWorkspace As TabWorkspace    '���ڶ��ǩ�ؼ�
Attribute mTabWorkspace.VB_VarHelpID = -1
Dim mlngWindowCount As Long         '��¼�Ѵ򿪴��ڵ�������������MDI����
Dim marrWindowName() As String      '��¼�Ѵ򿪴��ڵ�Name��������MDI����



Public Sub gmsThemeSkinSet(ByVal skinFile As String, ByVal SkinIni As String)
    '������������
    
    skinFW.LoadSkin gID.FolderStyles & skinFile, SkinIni
    
    Dim lngID As Long
    
    Select Case LCase(skinFile)
        Case LCase("Codejock.cjstyles")
            lngID = gID.WndThemeSkinCodejock
        Case LCase("Office2007.cjstyles")
            lngID = gID.WndThemeSkinOffice2007
        Case LCase("Office2010.cjstyles")
            lngID = gID.WndThemeSkinOffice2010
        Case LCase("Vista.cjstyles")
            lngID = gID.WndThemeSkinVista
        Case LCase("WinXPLuna.cjstyles")
            lngID = gID.WndThemeSkinWinXPLuna
        Case LCase("WinXPRoyale.cjstyles")
            lngID = gID.WndThemeSkinWinXPRoyale
        Case LCase("Zune.msstyles")
            lngID = gID.WndThemeSkinZune
    End Select
    For mlngID = gID.WndThemeSkinCodejock To gID.WndThemeSkinZune
        mcbsActions(mlngID).Checked = False
    Next
    If lngID > 0 Then
        mcbsActions(lngID).Checked = True
    End If
    
End Sub

Private Sub msAddAction()
    '����CommandBars��Action
    
    Dim cbsAction As CommandBarAction
    
    cBS.EnableActions   '����CommandBars��Actions����
    Set mcbsActions = cBS.Actions
    
    With mcbsActions
'        mcbsActions.Add "Id","Caption","TooltipText","DescriptionText","Category"
        
        .Add gID.Sys, "ϵͳ", "", "", "ϵͳ"
        
        .Add gID.SysExit, "�˳�", "", "", ""
        .Add gID.SysReLogin, "���µ�½", "", "", ""
        
        .Add gID.SysModifyPassword, "�����޸�", "", "", "frmSysAlterPWD"
        .Add gID.SysDepartment, "���Ź���", "", "", "frmSysDepartment"
        .Add gID.SysUser, "�û�����", "", "", "frmSysUser"
        .Add gID.SysLog, "��־�鿴", "", "", "frmSysLog"
        .Add gID.SysRole, "��ɫ����", "", "", "frmSysRole"
        .Add gID.SysFunc, "��������", "", "", "frmSysFunc"


        .Add gID.SysOutToExcel, "������Excel", "", "", ""
        .Add gID.SysOutToText, "���������±�", "", "", ""
        .Add gID.SysOutToWord, "������Word", "", "", ""
        .Add gID.SysPageSet, "��ӡҳ������", "", "", "frmSysPageSet"
        .Add gID.SysPrint, "��ӡ��", "", "", ""
        .Add gID.SysPrintPreview, "��ӡԤ��", "", "", ""
        
        .Add gID.SysSearch, "���ڼ���", "", "", ""
        .Add gID.SysSearch1Label, "���봰�����ƹؼ���", "", "", ""
        .Add gID.SysSearch2TextBox, "�ؼ��������", "", "", ""
        .Add gID.SysSearch3Button, "��������", "", "", ""
        .Add gID.SysSearch4ListBoxCaption, "�������Ĵ��ڱ����б�", "", "", ""
        .Add gID.SysSearch4ListBoxFormID, "�������Ĵ��������б�", "", "", ""
        .Add gID.SysSearch5Go, "��ת��ѡ������", "", "", ""
        
        
        
        .Add gID.TestWindow, "���Դ��ڲ˵�", "", "", "���Դ���"
        .Add gID.TestWindowFirst, "���Դ���һ", "", "", ""
        .Add gID.TestWindowSecond, "���Դ��ڶ�", "", "", ""
        .Add gID.TestWindowThird, "���Դ�����", "", "", ""
        .Add gID.TestWindowThour, "���Դ�����", "", "", "frmForm4"
        
        .Add gID.Wnd, "����", "", "", "����"
        
        .Add gID.WndResetLayout, "���ô��ڲ���", "", "", ""
        .Add gID.WndThemeSkinSet, "������������...", "", "", "frmSysSetSkin"
        
        .Add gID.WndThemeSkin, "��������", "", "", ""
        .Add gID.WndThemeSkinCodejock, "Codejock", "", "", ""
        .Add gID.WndThemeSkinOffice2007, "Office2007", "", "", ""
        .Add gID.WndThemeSkinOffice2010, "Office2010", "", "", ""
        .Add gID.WndThemeSkinVista, "Vista", "", "", ""
        .Add gID.WndThemeSkinWinXPLuna, "XPLuna", "", "", ""
        .Add gID.WndThemeSkinWinXPRoyale, "XPRoyale", "", "", ""
        .Add gID.WndThemeSkinZune, "msZune", "", "", ""
        
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
        .Add gID.WndSonVbAllBack, "�ָ�����", "", "", ""
        .Add gID.WndSonVbAllMin, "��С�������Ӵ���", "", "", ""
        .Add gID.WndSonVbArrangeIcons, "������С��ͼ��", "", "", ""
        .Add gID.WndSonVbCascade, "���", "", "", ""
        .Add gID.WndSonVbTileHorizontal, "ˮƽƽ��", "", "", ""
        .Add gID.WndSonVbTileVertical, "��ֱƽ��", "", "", ""
        
        
        .Add gID.Help, "����", "", "", "����"
        .Add gID.HelpAbout, "����", "", "", ""
        .Add gID.HelpDocument, "�����ĵ�", "", "", ""
        
        
        .Add gID.Other, "����", "", "", "����"
        
        .Add gID.OtherPane, "�������", "", "", ""
        .Add gID.OtherPaneMenuPopu, "PaneCaptionMenu", "", "", ""
        .Add gID.OtherPaneMenuPopuAutoFold, "�Զ���£", "", "", ""
        .Add gID.OtherPaneMenuPopuExpand, "ȫ��չ��", "", "", ""
        .Add gID.OtherPaneMenuPopuFold, "ȫ����£", "", "", ""
        .Add gID.OtherPaneMenuTitle, "�����˵�", "", "", ""
        .Add gID.OtherPaneIDFirst, "(��/��)" & mcbsActions(gID.OtherPaneMenuTitle).Caption, "", "", ""
        
        
        .Add gID.OtherTabWorkspacePopup, "���ǩ�Ҽ��˵�", "", "", ""
        
        .Add gID.StatusBarPane, "״̬��", "", "", ""
        .Add gID.StatusBarPaneProgress, "������", "", "", ""
        .Add gID.StatusBarPaneProgressText, "���Ȱٷֱ�", "", "", ""
        .Add gID.StatusBarPaneTime, "ϵͳʱ��", "", "", ""
        .Add gID.StatusBarPaneUserInfo, "��ǰ�û�", "", "", ""
        
       
    End With
    

    '���mcbsActions����������
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
    
    
    '******�ر����д��ڴ򿪵�Ȩ��******
    For Each cbsAction In mcbsActions
        If cbsAction.Id < 2000 Then         '�����ڶ����д��ڵ�IDֵ<2000
            If Len(cbsAction.Key) > 0 Then  '�����ڶ����д��ڵ�Nameֵ����Action�����Key������
                If Left(LCase(cbsAction.Key), 3) = "frm" Then   '�����������д��ڵ�Nameֵ��frm����ĸ��ͷ
                    cbsAction.Enabled = False
                End If
            End If
        End If
    Next
    '******����Ҫ����Ȩ�޵Ĵ���******
    mcbsActions(gID.SysModifyPassword).Enabled = True
    mcbsActions(gID.WndThemeSkinSet).Enabled = True
    mcbsActions(gID.TestWindowThour).Enabled = True
    
    '���ϵ�е�mcbsActions���������Ե���������
    For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        mcbsActions.Action(mlngID).DescriptionText = mcbsActions.Action(gID.WndThemeCommandBars).Caption & "����Ϊ��" & mcbsActions.Action(mlngID).DescriptionText
        mcbsActions.Action(mlngID).ToolTipText = mcbsActions.Action(mlngID).DescriptionText
    Next
    For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        mcbsActions.Action(mlngID).DescriptionText = mcbsActions.Action(gID.WndThemeTaskPanel).Caption & "����Ϊ��" & mcbsActions.Action(mlngID).DescriptionText
        mcbsActions.Action(mlngID).ToolTipText = mcbsActions.Action(mlngID).DescriptionText
    Next
    
End Sub

Private Sub msAddDesignerControls()
    'CommandBars�Զ���Ի���������������
    
    Dim cbsControls As CommandBarControls
    Dim cbsAction As CommandBarAction

    Set cbsControls = cBS.DesignerControls
    For Each cbsAction In mcbsActions
        If cbsAction.Id < 2000 Then
            cbsControls.Add xtpControlButton, cbsAction.Id, ""
        End If
    Next
    
End Sub

Private Sub msAddDockingPane()
    '�����������
    
    Dim paneLeft As XtremeDockingPane.Pane
    Dim paneList As XtremeDockingPane.Pane
    
    '���������˵���������һ��DockingPane
    Set paneLeft = DockingPN.CreatePane(gID.OtherPaneIDFirst, 240, 240, DockLeftOf, Nothing)
    paneLeft.Title = mcbsActions(gID.OtherPaneMenuTitle).Caption
    paneLeft.TitleToolTip = paneLeft.Title & mcbsActions(gID.OtherPane).Caption
    paneLeft.Handle = picTaskPL.hwnd    '���������TaskPanel������PictureBox�ؼ��ҿ��ڸ������PanelLeft��
    paneLeft.Options = PaneHasMenuButton    '��ʾPopu����
    
        
'    '�ڶ���DockingPane
'    Set paneList = DockingPN.CreatePane(gID.OtherPaneIDSecond, 240, 240, DockLeftOf, Nothing)
'    paneList.Title = "       "
'    paneList.Handle = picList.hWnd
'    paneList.AttachTo paneLeft  '��������һ��Pane��
'    paneLeft.Selected = True    '��ʾ��һ��Pane
 
End Sub

Private Sub msAddKeyBindings()
    '������ݼ�
    
    With cBS.KeyBindings
'''        .Add 0, &H70, gID.HelpDocument
        .AddShortcut gID.HelpDocument, "F1"
        .AddShortcut gID.SysExit, "F10"
    End With
    
End Sub

Private Sub msAddMenu()
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
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysDepartment, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysUser, ""
        .Add xtpControlButton, gID.SysRole, ""
        .Add xtpControlButton, gID.SysFunc, ""
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysLog, "")
        cbsMenuCtrl.BeginGroup = True
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysOutToExcel, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysOutToText, ""
        .Add xtpControlButton, gID.SysOutToWord, ""
        
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysPageSet, "")
        cbsMenuCtrl.BeginGroup = True
        .Add xtpControlButton, gID.SysPrintPreview, ""
        .Add xtpControlButton, gID.SysPrint, ""
                
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysReLogin, "")
        cbsMenuCtrl.BeginGroup = True
        Set cbsMenuCtrl = .Add(xtpControlButton, gID.SysExit, "")
        cbsMenuCtrl.BeginGroup = True
    End With
    
    
    '���Դ��ڲ˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.TestWindow, "")
    With cbsMenuMain.CommandBar.Controls
        For mlngID = gID.TestWindowFirst To gID.TestWindowThour
            .Add xtpControlButton, mlngID, ""
        Next
    End With
    
    
    '�������˵�
    Set cbsMenuMain = cbsMenuBar.Controls.Add(xtpControlPopup, gID.Wnd, "")
    
    '��/�� �����˵�
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, gID.OtherPaneIDFirst, "")
'    cbsMenuCtrl.BeginGroup = True
    
    '���ò���
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, gID.WndResetLayout, "")
    cbsMenuCtrl.BeginGroup = True
    
    '������������
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, gID.WndThemeSkinSet, "")
    cbsMenuCtrl.BeginGroup = True
    
    '����ID35001
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlButton, XTP_ID_CUSTOMIZE, "�Զ��幤����...")
    cbsMenuCtrl.BeginGroup = True
    
    '����ID59392
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, 0, "�������б�")
    cbsMenuCtrl.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, ""
    
    '����������ʽ
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, gID.WndThemeSkin, "")
    cbsMenuCtrl.BeginGroup = True
    With cbsMenuCtrl.CommandBar.Controls
        For mlngID = gID.WndThemeSkinCodejock To gID.WndThemeSkinZune
            .Add xtpControlButton, mlngID, ""
        Next
    End With
    
    'CommandBars�����������Ӳ˵�
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, gID.WndThemeCommandBars, "")
    With cbsMenuCtrl.CommandBar.Controls
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            .Add xtpControlButton, mlngID, ""
        Next
    End With
    
    'TaskPanel�����˵������Ӳ˵�
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, gID.WndThemeTaskPanel, "")
    With cbsMenuCtrl.CommandBar.Controls
        For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
            .Add xtpControlButton, mlngID, ""
        Next
    End With
    
    '�Ӵ��ڿ���
    Set cbsMenuCtrl = cbsMenuMain.CommandBar.Controls.Add(xtpControlPopup, gID.WndSon, "")
    cbsMenuCtrl.BeginGroup = True
    With cbsMenuCtrl.CommandBar.Controls
        For mlngID = gID.WndSonCloseAll To gID.WndSonVbTileVertical
            .Add xtpControlButton, mlngID, ""
            If mlngID = gID.WndSonVbAllBack Then .Find(, mlngID).BeginGroup = True
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

Private Sub msAddPopupMenu()
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
        For mlngID = gID.WndSonCloseAll To gID.WndSonCloseRight
            .Add xtpControlButton, mlngID, ""
        Next
    End With
    
End Sub

Private Sub msAddStatuBar()
    '����״̬��
    
    Dim statuBar As XtremeCommandBars.StatusBar
    
    Set statuBar = cBS.StatusBar
    With statuBar
        .AddPane 0      'ϵͳPane����ʾCommandBarActions��Description
        .SetPaneStyle 0, SBPS_STRETCH
        
        .AddPane gID.StatusBarPaneUserInfo
        .FindPane(gID.StatusBarPaneUserInfo).Caption = mcbsActions(gID.StatusBarPaneUserInfo).Caption
'        .FindPane(gID.StatusBarPaneUserInfo).Text = "С��"
        
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

Private Sub msAddTaskPanelItem()
    '���������˵�
    'ע�⣺����ĵ����˵����ǲ˵�������һ����ʾ��ʽ
    
    Dim taskGroup As TaskPanelGroup
    Dim taskItem As TaskPanelGroupItem
    
    
    'ϵͳ
    Set taskGroup = TaskPL.Groups.Add(gID.Sys, mcbsActions(gID.Sys).Caption)
    With taskGroup.Items
        .Add gID.SysModifyPassword, mcbsActions(gID.SysModifyPassword).Caption, xtpTaskItemTypeLink
        .Add gID.SysDepartment, mcbsActions(gID.SysDepartment).Caption, xtpTaskItemTypeLink
        .Add gID.SysUser, mcbsActions(gID.SysUser).Caption, xtpTaskItemTypeLink
        .Add gID.SysRole, mcbsActions(gID.SysRole).Caption, xtpTaskItemTypeLink
        .Add gID.SysFunc, mcbsActions(gID.SysFunc).Caption, xtpTaskItemTypeLink
        .Add gID.SysLog, mcbsActions(gID.SysLog).Caption, xtpTaskItemTypeLink
        
        For mlngID = gID.SysOutToExcel To gID.SysPrintPreview
            .Add mlngID, mcbsActions(mlngID).Caption, xtpTaskItemTypeLink
        Next
        
        .Add gID.SysReLogin, mcbsActions(gID.SysReLogin).Caption, xtpTaskItemTypeLink
        .Add gID.SysExit, mcbsActions(gID.SysExit).Caption, xtpTaskItemTypeLink
    End With
    
    
    '���Դ���
    Set taskGroup = TaskPL.Groups.Add(gID.TestWindow, mcbsActions(gID.TestWindow).Caption)
    With taskGroup.Items
        For mlngID = gID.TestWindowFirst To gID.TestWindowThour
            .Add mlngID, mcbsActions(mlngID).Caption, xtpTaskItemTypeLink
        Next
    End With
    
    
    '����
    Set taskGroup = TaskPL.Groups.Add(gID.Wnd, mcbsActions(gID.Wnd).Caption)
    
    '����
    Set taskItem = taskGroup.Items.Add(gID.WndResetLayout, mcbsActions(gID.WndResetLayout).Caption, xtpTaskItemTypeLink)
    
    Set taskItem = taskGroup.Items.Add(gID.WndThemeSkinSet, mcbsActions(gID.WndThemeSkinSet).Caption, xtpTaskItemTypeLink)
    
    '��������
    Set taskItem = taskGroup.Items.Add(gID.WndThemeSkin, mcbsActions(gID.WndThemeSkin).Caption, xtpTaskItemTypeText)
    taskItem.Bold = True
    For mlngID = gID.WndThemeSkinCodejock To gID.WndThemeSkinZune
        taskGroup.Items.Add mlngID, mcbsActions(mlngID).Caption, xtpTaskItemTypeLink
    Next
    
    '����������
    Set taskItem = taskGroup.Items.Add(gID.WndThemeCommandBars, mcbsActions(gID.WndThemeCommandBars).Caption, xtpTaskItemTypeText)
    taskItem.Bold = True
    For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        taskGroup.Items.Add mlngID, mcbsActions(mlngID).Caption, xtpTaskItemTypeLink
    Next
    
    '�����˵�����
    Set taskItem = taskGroup.Items.Add(gID.WndThemeTaskPanel, mcbsActions(gID.WndThemeTaskPanel).Caption, xtpTaskItemTypeText)
    taskItem.Bold = True
    For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        taskGroup.Items.Add mlngID, mcbsActions(mlngID).Caption, xtpTaskItemTypeLink
    Next
    
    
    '����
    Set taskGroup = TaskPL.Groups.Add(gID.Help, mcbsActions(gID.Help).Caption)
    taskGroup.Items.Add gID.HelpDocument, mcbsActions(gID.HelpDocument).Caption, xtpTaskItemTypeLink
    taskGroup.Items.Add gID.HelpAbout, mcbsActions(gID.HelpAbout).Caption, xtpTaskItemTypeLink
    
    '���ͼ��
    Dim imgIcon As MSComctlLib.ListImage
    For Each taskGroup In TaskPL.Groups
        For Each taskItem In taskGroup.Items
            For Each imgIcon In imgListCommandBars.ListImages
                If Val(imgIcon.Tag) = Val(taskItem.Id) Then
                    taskItem.IconIndex = imgIcon.Index
                    Exit For
                End If
            Next
        Next
    Next
    
End Sub

Private Sub msAddToolBar()
    '����������
    
    Dim cbsBar As CommandBar
    Dim cbsCtr As CommandBarControl
    
    
    'ϵͳ����������
    Set cbsBar = cBS.Add(mcbsActions(gID.Sys).Caption, xtpBarTop)
    With cbsBar.Controls
        .Add xtpControlButton, gID.SysReLogin, ""
        .Add xtpControlButton, gID.SysExit, ""
        
        For mlngID = gID.SysOutToExcel To gID.SysPrintPreview
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    
    '���ڼ���������
    Set cbsBar = cBS.Add(mcbsActions(gID.SysSearch).Caption, xtpBarTop)
    With cbsBar.Controls
        .Add xtpControlLabel, gID.SysSearch1Label, ""
        Set cbsCtr = .Add(xtpControlEdit, gID.SysSearch2TextBox, "")
        cbsCtr.EditHint = "���봰�ڹؼ���"
        .Add xtpControlButton, gID.SysSearch3Button, ""
        Set cbsCtr = .Add(xtpControlComboBox, gID.SysSearch4ListBoxCaption, "")
        cbsCtr.Width = 200
        cbsCtr.EditHint = "���б���ѡ��һ�����ڱ���"
        Set cbsCtr = .Add(xtpControlComboBox, gID.SysSearch4ListBoxFormID, "")
        cbsCtr.Visible = False  '����������β����û��������������洰���Ӧ��IDֵ
        .Add xtpControlButton, gID.SysSearch5Go, ""
    End With
    
    
    '��������
    Set cbsBar = cBS.Add(mcbsActions(gID.WndThemeSkin).Caption, xtpBarTop)
    cbsBar.Visible = False
    With cbsBar.Controls
        For mlngID = gID.WndThemeSkinCodejock To gID.WndThemeSkinZune
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    '����������
    Set cbsBar = cBS.Add(mcbsActions(gID.WndThemeCommandBars).Caption, xtpBarTop)
    cbsBar.Visible = False
    With cbsBar.Controls
        For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    '�����˵�����
    Set cbsBar = cBS.Add(mcbsActions(gID.WndThemeTaskPanel).Caption, xtpBarTop)
    cbsBar.Visible = False
    With cbsBar.Controls
        For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
            Set cbsCtr = .Add(xtpControlButton, mlngID, "")
            cbsCtr.BeginGroup = True
        Next
    End With
    
    

End Sub

Private Sub msCommandBarPopu(ByVal PID As Long)
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

Private Sub msLeftClick(ByVal CID As Long)
    'CommandBar��TaskPanelGroupItem����������Ӧ��������
    
    Dim strKey As String
    
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
            Case .WndThemeSkinCodejock To .WndThemeSkinZune
                Call msThemeSkin(CID)
            Case .WndResetLayout
                Call msResetLayout
            Case .OtherPaneIDFirst
                DockingPN.FindPane(CID).Closed = Not DockingPN.FindPane(CID).Closed
            Case .SysReLogin
                Dim strName As String, strPWD As String
                strName = gID.UserLoginName
                strPWD = gID.UserPassword
                Unload Me
                Call Main
                frmSysLogin.ucTC = strName
                frmSysLogin.Text1.Text = strPWD
            Case .SysExit
                Unload Me
            Case .HelpAbout
                Dim strAbout As String
                strAbout = "���ƣ�" & App.Title & vbCrLf & _
                           "�汾��" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
                           "��Ȩ���У�WZD_XMH"
                MsgBox strAbout, vbInformation, "����" & App.Title
            Case .SysOutToExcel
                If MsgBox("ȷ������ǰ������ݵ���ΪExcel�ļ���", vbQuestion + vbOKCancel, "����ѯ��") = vbOK Then Call gsGridToExcel(ActiveForm.ActiveControl)
            Case .SysOutToText
                If MsgBox("ȷ������ǰ������ݵ���Ϊ�ı��ļ���", vbQuestion + vbOKCancel, "����ѯ��") = vbOK Then Call gsGridToText(ActiveForm.ActiveControl)
            Case .SysOutToWord
                If MsgBox("ȷ������ǰ������ݵ���ΪWord�ĵ���", vbQuestion + vbOKCancel, "����ѯ��") = vbOK Then Call gsGridToWord(ActiveForm.ActiveControl)
            Case .SysPrint
                If MsgBox("ȷ����ӡ��ǰ���������", vbQuestion + vbOKCancel, "��ӡѯ��") = vbOK Then Call gsGridPrint
            Case .SysPrintPreview
                Call gsGridPrintPreview
            Case .SysPageSet
                Call gsGridPageSet
            Case .SysSearch3Button
                Call msSearchWindow
            Case .SysSearch5Go, .SysSearch4ListBoxCaption
                If Len(cBS.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxCaption).Text) > 0 Then
                    Call msLeftClick(CLng(cBS.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxFormID).List(cBS.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxCaption).ListIndex)))
                End If
            Case Else
                
                strKey = LCase(mcbsActions.Action(CID).Key)
                If Left(strKey, 3) = "frm" Then
                    If mcbsActions.Action(CID).Enabled Then
                        Select Case strKey
                            Case LCase("frmSysSetSkin"), LCase("frmSysAlterPWD")
                                Call gsOpenTheWindow(strKey, vbModal, vbNormal)
                            Case Else
                                Call gsOpenTheWindow(strKey)
                                mcbsActions.Action(CID).Checked = True  '��Ǹô��ڱ���
                                If mlngWindowCount < Forms.Count - 1 Then
                                    mlngWindowCount = Forms.Count - 1
                                    Call msWindowNameAdd(strKey)        '�����Ѵ򿪴��ڵ�����
                                End If
                        End Select
                    End If
                Else
                    MsgBox "��" & mcbsActions(CID).Caption & "������δ���壡", vbExclamation, "�����"
                End If
        End Select
    End With
    
End Sub

Private Sub msResetLayout()
    '���ô��ڲ��֣�CommandBars��Dockingpane�ؼ�����
    
    Dim cBar As CommandBar
    Dim L As Long, T As Long, R As Long, B As Long

    For Each cBar In cBS
        cBar.Reset
        cBar.Visible = True
    Next
    
    For mlngID = 2 To cBS.Count
        cBS.GetClientRect L, T, R, B
        cBS.DockToolBar cBS(mlngID), 0, B, xtpBarTop
    Next

    Dim pnRe As XtremeDockingPane.Pane
    For Each pnRe In DockingPN
        pnRe.Closed = False
        pnRe.Hidden = False
        DockingPN.DockPane pnRe, 240, 240, DockLeftOf
    Next

End Sub

Private Sub msSearchWindow()
    '��������ָ���ؼ��ֵĴ���
    
    Dim strName As String
    Dim cbsAction As CommandBarAction
    Dim cbsCtrlCaption As CommandBarComboBox
    Dim cbsCtrlFormID As CommandBarComboBox
    Dim blnClear As Boolean
    
    strName = LCase(Trim(cBS.FindControl(xtpControlEdit, gID.SysSearch2TextBox).Text))
    If Len(strName) = 0 Then Exit Sub
    
    Set cbsCtrlCaption = cBS.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxCaption)
    Set cbsCtrlFormID = cBS.FindControl(xtpControlComboBox, gID.SysSearch4ListBoxFormID)
    
    For Each cbsAction In mcbsActions
        If cbsAction.Id < 2000 Then     '���д��ڵ�IDС��2000
            If LCase(Left(cbsAction.Key, 3)) = "frm" Then   '���ڵ�Name������frm��ͷ
                If InStr(LCase(cbsAction.Caption), strName) > 0 Then
                    If Not blnClear Then
                        cbsCtrlCaption.Clear
                        cbsCtrlFormID.Clear
                        blnClear = True
                    End If
                    cbsCtrlCaption.AddItem cbsAction.Caption
                    cbsCtrlFormID.AddItem cbsAction.Id
                End If
            End If
        End If
    Next
    
    If blnClear Then
        If cbsCtrlCaption.ListCount > 0 Then cbsCtrlCaption.ListIndex = 1
    Else
        cbsCtrlCaption.ListIndex = 0
    End If
    
    Set cbsCtrlCaption = Nothing
    Set cbsCtrlFormID = Nothing
    
End Sub

Private Sub msThemeCommandBar(ByVal CID As Long)
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
    
    For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        mcbsActions(mlngID).Checked = False
    Next
    mcbsActions(CID).Checked = True
    
End Sub

Private Sub msThemeSkin(ByVal SID As Long)
    '����Ƥ�����Ӵ��ڡ����������������ʽ���������仯
    
    Dim strFile As String, strIni As String

    Select Case SID
        Case gID.WndThemeSkinCodejock
            strFile = "Codejock.cjstyles"
        Case gID.WndThemeSkinOffice2007
            strFile = "Office2007.cjstyles"
        Case gID.WndThemeSkinOffice2010
            strFile = "Office2010.cjstyles"
        Case gID.WndThemeSkinVista
            strFile = "Vista.cjstyles"
        Case gID.WndThemeSkinWinXPLuna
            strFile = "WinXPLuna.cjstyles"
        Case gID.WndThemeSkinWinXPRoyale
            strFile = "WinXPRoyale.cjstyles"
        Case gID.WndThemeSkinZune
            strFile = "Zune.msstyles"
        Case Else
            strFile = ""
    End Select
    
    gID.SkinPath = strFile
    gID.SkinIni = strIni
    Call gmsThemeSkinSet(strFile, strIni)

End Sub

Private Sub msThemeTaskPanel(ByVal TID As Long)
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
    
    For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        mcbsActions(mlngID).Checked = False
    Next
    mcbsActions(TID).Checked = True
    
End Sub

Private Sub msWindowControl(ByVal WID As Long)
    '�Ӵ��ڿ���
    
    Dim frmTag As Form
    Dim C As Long
    Dim itemCur As XtremeCommandBars.TabControlItem
    
    With gID
        Select Case WID
            Case .WndSonCloseAll    '�ر����д���
                For Each frmTag In Forms
                    If frmTag.Name <> gMDI.Name Then Unload frmTag
                Next
            Case .WndSonCloseCurrent    '�رյ�ǰ����
                If Not ActiveForm Is Nothing Then Unload ActiveForm
            Case .WndSonCloseLeft   '�ر���ര��
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
            Case .WndSonCloseOther  '�ر���������
                If Forms.Count > 1 Then
                    For Each frmTag In Forms
                        If frmTag.Name <> gMDI.Name Then
                            If Not (frmTag.Name = ActiveForm.Name And frmTag.Caption = ActiveForm.Caption) Then
                                Unload frmTag
                            End If
                        End If
                    Next
                End If
            Case .WndSonCloseRight  '�ر��Ҳര��
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
            Case .WndSonVbAllBack
                For Each frmTag In Forms
                    If frmTag.Name <> gMDI.Name Then frmTag.WindowState = vbNormal
                Next
            Case .WndSonVbAllMin
                For Each frmTag In Forms
                    If frmTag.Name <> gMDI.Name Then frmTag.WindowState = vbMinimized
                Next
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

Private Sub msWindowNameAdd(ByVal strFormName As String)
    '���Ѵ򿪴��ڵ�Name��ӽ�������
    
    strFormName = LCase(strFormName)
    ReDim Preserve marrWindowName(mlngWindowCount - 1)
    marrWindowName(mlngWindowCount - 1) = strFormName

End Sub

Private Sub msWindowNameDel()
    '���ѹرյĴ���Name��������ɾ��
    
    Dim strFormName As Variant
    Dim lngCount As Long
    
    'Ѱ�ұ��رմ��ڵ�Nameֵ
    lngCount = Forms.Count - 1
    For Each strFormName In marrWindowName
        strFormName = LCase(strFormName)
        mlngID = 0
        Do While (mlngID <= lngCount)
            If LCase(Forms(mlngID).Name) = strFormName Then Exit Do
            mlngID = mlngID + 1
        Loop
        If mlngID > lngCount Then Exit For
    Next
    
    'ɾ�����رմ��ڵ�Nameֵ
    For mlngID = 0 To UBound(marrWindowName)
        If LCase(marrWindowName(mlngID)) = strFormName Then
            If mlngID < UBound(marrWindowName) Then
                For lngCount = mlngID To UBound(marrWindowName) - 1
                    marrWindowName(lngCount) = marrWindowName(lngCount + 1)
                Next
            End If
            Exit For
        End If
    Next
    ReDim Preserve marrWindowName(mlngWindowCount)
    
    'ȥ��Action��Checked����
    Call gsUnCheckedAction(strFormName)

End Sub



Private Sub cBS_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '������¼�

    Call msLeftClick(Control.Id)

End Sub

Private Sub cBS_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    'ȥ�������������е��ѹرմ���
    
    If mlngWindowCount > Forms.Count - 1 Then
        mlngWindowCount = Forms.Count - 1
        Call msWindowNameDel
    End If
    
End Sub

Private Sub cBS_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '����CommandBars�����еĿؼ���Enabled��Checked״̬
    
    Dim blnActForm As Boolean
    Dim blnActCtrl As Boolean
    Dim blnGrid As Boolean
    Dim blnResult As Boolean
    Dim ctlCur As Control
    Dim strType As String
    
    Select Case Control.Id
        Case gID.SysOutToExcel To gID.SysPrintPreview
            If Not ActiveForm Is Nothing Then
        
                blnActForm = True   '�Ƿ��л����
                Set ctlCur = ActiveForm.ActiveControl
                If Not ctlCur Is Nothing Then
                    
                    blnActCtrl = True   '�Ƿ��л�ؼ�
                    If (TypeOf ctlCur Is VSFlex8Ctl.VSFlexGrid) Or (TypeOf ctlCur Is FlexCell.Grid) Then
                        
                        blnGrid = True  '�Ƿ��Ǳ��ؼ����ݽ�֧��VSFlexGrid��FlexCell��
                        strType = TypeName(ctlCur)
                        'Ȩ���ж�
                        
                    End If
                End If
            End If
            
            blnResult = blnActForm And blnActCtrl And blnGrid
            mcbsActions(Control.Id).Enabled = blnResult
            
        Case Else
                
    End Select
    
    
End Sub

Private Sub DockingPN_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    
    If Action = PaneActionClosed Then
        If Pane.Id = gID.OtherPaneIDFirst Then
'            Debug.Print Pane.Id, Pane.Title, Pane.TitleToolTip
            mcbsActions(Pane.Id).Checked = False
        End If
    End If
    
End Sub

Private Sub DockingPn_PanePopupMenu(ByVal Pane As XtremeDockingPane.IPane, ByVal X As Long, ByVal Y As Long, Handled As Boolean)
    '�����˵������е�Popu�˵�����

    If Pane.Id = gID.OtherPaneIDFirst Then
        mcbsPopupNav.ShowPopup , X * 15, Y * 15     'ֻ֪������15��λ�ò��ԣ�����x��y�ĵ�λ�����أ�������Ҫ��羡�
    End If
    
End Sub

Private Sub MDIForm_Load()

'''    Debug.Print Screen.TwipsPerPixelX, Screen.TwipsPerPixelY    '����ˮƽ�봹ֱ�����Ķ����ÿһ�����е���������Խ����1����=15�
'''    Me.Width = 15360    '���ô��ڴ�С1024*768����
'''    Me.Height = 11520
    Me.Caption = App.Title
    
    CommandBarsGlobalSettings.App = App
    
    Call msAddAction        '����Actions����
    Call msAddMenu          '�����˵���
    Call msAddToolBar       '����������
    Call msAddDockingPane   '�����������
    Call msAddPopupMenu     '����Popup�˵�
    Call msAddTaskPanelItem '���������˵�
    Call msAddStatuBar      '����״̬��
    Call msAddKeyBindings   '��ӿ�ݼ�,�ŵ�LoadCommandBars�������������Ч������
    Call msAddDesignerControls  'CommandBars�Զ���Ի�����ʹ�õ�
    

    cBS.AddImageList imgListCommandBars '���ͼ��
    cBS.EnableCustomization True        '�����Զ��壬��������÷�������CommandBars�趨֮��
    cBS.Options.UpdatePeriod = 250      '����CommandBars��Update�¼���ִ�����ڣ�Ĭ��100ms
    TaskPL.SetImageList imgListCommandBars  '��ӵ����˵�ͼ�꣬��cBS����һ��

    Set mTabWorkspace = cBS.ShowTabWorkspace(True)    '�����ڶ��ǩ��ʾ
    mTabWorkspace.Flags = xtpWorkspaceShowActiveFiles Or xtpWorkspaceShowCloseSelectedTab
    
    
    'ע�⣺��������������DockingPanel�ؼ���������CommandBars�ؼ��������Ҽ�CommandBars�ؼ���ѡ���Ƶ�����,��ʾ��������
    'ʹDockingPanel��CommandBars�ؼ�������������Pane��CommandBar�ؼ���λ���ƶ�����С�仯ʱ������ʾ������
    DockingPN.SetCommandBars Me.cBS
    
    DockingPN.Options.AlphaDockingContext = True    '��ʾDockingλ��ָ���ǩ��Ӱ��
    DockingPN.Options.ShowDockingContextStickers = True
    DockingPN.VisualTheme = ThemeWord2007
    
    
    
    '******��ע����ж�ȡ����ϴ��˳�ʱ����������������******
    'ע����б����õļ�������ֵ��ʼ��
    With gID
        .OtherSaveRegistryKey = Me.Name
        .OtherSaveAppName = Me.Name & "Layout"
        .OtherSaveCommandBarsSection = "CommandBarsLayout"
        .OtherSaveDockingPaneSection = "DockingPaneLayout"
    End With

    '����λ��
    Dim WS As Long, L As Long, T As Long, W As Long, H As Long
    WS = Val(GetSetting(Me.Name, gID.OtherSaveSettings, "WindowState", 2))
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

    'CommandBars��������
    cBS.LoadCommandBars gID.OtherSaveRegistryKey, gID.OtherSaveAppName, gID.OtherSaveCommandBarsSection
'''    Call msAddKeyBindings   '��ӿ�ݼ�

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
    
    Dim lngSaveID As Long
    
    'skinFW�������Ᵽ��
    lngSaveID = 0
    For mlngID = gID.WndThemeSkinCodejock To gID.WndThemeSkinZune
        If mcbsActions(mlngID).Checked Then
            lngSaveID = mlngID
            Exit For
        End If
    Next
    SaveSetting Me.Name, gID.OtherSaveSettings, gID.OtherSaveSkinID, lngSaveID
    SaveSetting Me.Name, gID.OtherSaveSettings, gID.OtherSaveSkinPath, gID.SkinPath
    SaveSetting Me.Name, gID.OtherSaveSettings, gID.OtherSaveSkinIni, gID.SkinIni
    
    'CommandBars���ֱ���
    cBS.SaveCommandBars gID.OtherSaveRegistryKey, gID.OtherSaveAppName, gID.OtherSaveCommandBarsSection
    
    'CommandBas���Ᵽ��
    lngSaveID = gID.WndThemeCommandBarsVS2008
    For mlngID = gID.WndThemeCommandBarsOffice2000 To gID.WndThemeCommandBarsWinXP
        If mcbsActions(mlngID).Checked Then
            lngSaveID = mlngID
            Exit For
        End If
    Next
    SaveSetting Me.Name, gID.OtherSaveSettings, "ThemeCommandBas", lngSaveID
    
    'Taskpanels��Popu����
    lngSaveID = 0
    If mcbsActions(gID.OtherPaneMenuPopuAutoFold).Checked Then lngSaveID = 1
    SaveSetting Me.Name, gID.OtherSaveSettings, "AutoFold", lngSaveID
    
    'DockingPaneλ�ñ���
    DockingPN.SaveState gID.OtherSaveRegistryKey, gID.OtherSaveAppName, gID.OtherSaveDockingPaneSection
    
    'TaskPanel��Popu
    lngSaveID = gID.WndThemeTaskPanelNativeWinXP
    For mlngID = gID.WndThemeTaskPanelListView To gID.WndThemeTaskPanelVisualStudio2010
        If mcbsActions(mlngID).Checked Then
            lngSaveID = mlngID
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
    
    '
    Set mcbsActions = Nothing
    Set mcbsPopupNav = Nothing
    Set mcbsPopupTab = Nothing
    Set mTabWorkspace = Nothing
    Set gMDI = Nothing
    
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
    
    If mcbsActions(Item.Id).Enabled Then
        Call msLeftClick(Item.Id)
    Else
        MsgBox "״̬������ �� û�����Ȩ�ޣ�", vbExclamation
    End If
    
End Sub
