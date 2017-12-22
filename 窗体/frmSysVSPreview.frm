VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form frmSysVSPreview 
   Caption         =   "¥Ú”°‘§¿¿"
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5550
   StartUpPosition =   3  '¥∞ø⁄»± °
   WindowState     =   2  'Maximized
   Begin VSPrinter8LibCtl.VSPrinter VP 
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      _cx             =   8070
      _cy             =   1931
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
      Zoom            =   1.69340463458111
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
End
Attribute VB_Name = "frmSysVSPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    
    Dim gridCtl As Control
    
        
    Me.Caption = gMDI.cBS.Actions(gID.SysPrintPreview).Caption
    Me.Icon = gMDI.imgListCommandBars.ListImages("SysPreview").Picture
    
    If gMDI.ActiveForm Is Nothing Then GoTo LineBreak
    If gMDI.ActiveForm.ActiveControl Is Nothing Then GoTo LineBreak
     
    Set gridCtl = gMDI.ActiveForm.ActiveControl
    If Not (TypeOf gridCtl Is VSFlex8Ctl.VSFlexGrid) Then GoTo LineBreak
    
    With VP
        .ShowGuides = gdShow
        .Navigation = vpnvAll
        
        .PaperSize = gMDI.VPMain.PaperSize
        .PaperBin = gMDI.VPMain.PaperBin
        .Orientation = gMDI.VPMain.Orientation
        .MarginTop = gMDI.VPMain.MarginTop
        .MarginBottom = gMDI.VPMain.MarginBottom
        .MarginLeft = gMDI.VPMain.MarginLeft
        .MarginRight = gMDI.VPMain.MarginRight
        
        If gID.VSPrintPageSet Then .PrintDialog pdPageSetup
        .StartDoc
        .RenderControl = gridCtl.hwnd
        .EndDoc
    End With
    
    Exit Sub
    
LineBreak:
    MsgBox "“≥√ÊºÏ≤‚“Ï≥££¨«Î÷ÿ ‘£°", vbExclamation
    Unload Me
    
End Sub

Private Sub Form_Resize()
    VP.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With gMDI.VPMain
        .PaperSize = VP.PaperSize
        .PaperBin = VP.PaperBin
        .Orientation = VP.Orientation
        .MarginTop = VP.MarginTop
        .MarginBottom = VP.MarginBottom
        .MarginLeft = VP.MarginLeft
        .MarginRight = VP.MarginRight
    End With
End Sub
