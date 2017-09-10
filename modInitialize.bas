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
    
End Type

Public gID As typCommandBarID   '全局CommandBars的ID变量


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

        
    End With
    
    frmSysTest.Show
    
End Sub
