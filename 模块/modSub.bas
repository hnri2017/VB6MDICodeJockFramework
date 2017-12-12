Attribute VB_Name = "modSub"
Option Explicit


Public Sub Main()
    
    App.Title = "VB6+Codejock"
    Set gMDI = frmSysMDI    '��ʼ������������ȫ�ֱ���

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
        
        
        '�뽫���в˵�CommandBrs��IDֵ������2000���£���
        
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
        .FolderNet = "\\192.168.12.100\��֮��\��������\��������\WZDMSר��(��)\��֮�ȹ���ϵͳ\"
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
        
        .UserAdmin = "Admin"    '���������û�
        .UserSystem = "System"  '���������û�
        
        .CnDatabase = "db_Test"
        .CnPassword = "test"
        .CnSource = "192.168.2.9"
        .CnUserID = "wzd_test"
        .CnString = "Provider='SQLOLEDB';Persist Security Info=False;Data Source='" & .CnSource & _
                    "';User ID='" & .CnUserID & "';Password='" & .CnPassword & _
                    "';Initial Catalog='" & .CnDatabase & "';"   '���Լ�64λϵͳ������Data Source�м�Ҫ�ո�������ܽ������ӣ���������Բ��ã���֪Ϊ��
        
        .FuncButton = "��ť"
        .FuncControl = "����"
        .FuncForm = "����"
        .FuncMainMenu = "���˵�"
        
    End With
    
    '���ô�������
    gMDI.skinFW.ApplyOptions = xtpSkinApplyColors Or xtpSkinApplyFrame Or xtpSkinApplyMenus Or xtpSkinApplyMetrics
    gMDI.skinFW.ApplyWindow gMDI.hwnd
    gID.SkinPath = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveSkinPath, "")
    gID.SkinIni = GetSetting(gMDI.Name, gID.OtherSaveSettings, gID.OtherSaveSkinIni, "")
    Call gMDI.gmsThemeSkinSet(gID.SkinPath, gID.SkinIni)

    frmSysLogin.Show    '��ʾ��½����

End Sub


Public Sub gsAlarmAndLog(Optional ByVal strErr As String, Optional ByVal MsgButton As VbMsgBoxStyle = vbCritical)
    '�쳣��ʾ��д���쳣��־
    
    Dim strMsg As String
    
    strMsg = "�쳣���ţ�" & Err.Number & vbCrLf & "�쳣������" & Err.Description
    MsgBox strMsg, MsgButton, strErr
    Call gsFileWrite(gID.FileLog, strErr & vbTab & Replace(strMsg, vbCrLf, vbTab))
    
End Sub


Public Sub gsFileWrite(ByVal strFile As String, ByVal strContent As String, _
    Optional ByVal OpenMode As genmFileOpenType = udAppend, _
    Optional ByVal WriteMode As genmFileWriteType = udPrint)
    '��ָ��������ָ���ķ�ʽд��ָ���ļ���
    
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
        Case Else   '����Ե���udAppend
            Open strFile For Append As #intNum
    End Select
    
    strTime = Format(Now, "yyyy-MM-dd hh:mm:ss")
    Select Case WriteMode
        Case udWrite
            Write #intNum, strTime, strContent
        Case udPut
            Put #intNum, , strTime & vbTab & strContent
        Case Else   '����Ե���udPrint
            Print #intNum, strTime, strContent
    End Select
    
    Close #intNum
    
End Sub


Public Sub gsFormScrollBar(ByRef frmCur As Form, ByRef ctlMv As Control, _
    ByRef Hsb As HScrollBar, ByRef Vsb As VScrollBar, _
    Optional ByVal lngMW As Long = 12000, _
    Optional ByVal lngMH As Long = 9000, _
    Optional ByVal lngHV As Long = 255)
    
    'frmCur�����������ڵĴ���
    'ctlMv�������еĿؼ��������������⣩���ڴ������ؼ���
    'Hsb������frmCur��ˮƽ�������ؼ�
    'Vsb������frmCur�д�ֱ�������ؼ�
    'lngMW�����岻���ֹ������Ŀ��
    'lngMH�����岻���ֹ������ĸ߶�
    'lngHV����������խ�߿�Ȼ�߶ȡ�
    '***ע��ע��ע�⣺�������ؼ����������������У��Ҳ��ܷ��������ؼ�ctlMv��*******
    
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
    
'    '�ڴ�������Ӵ��ڿؼ�ctlMove�������������ؼ�����������У�Ȼ
'    '��������Ʒֱ�ΪHsb\Vsb��ˮƽ\��ֱ�������ڴ����У�������������봰����
'    'Ȼ���ڴ�������������¼����ü���
'Private Sub Form_Resize()
'    Call gsFormScrollBar(Me, Me.ctlMove, Me.Hsb, Me.Vsb, 12000, 9000)  'ע�ⳤ������޸�
'End Sub
'Private Sub Hsb_Change()
'    ctlMove.Left = -Hsb.Value
'End Sub
'
'Private Sub Hsb_Scroll()
'    Call Hsb_Change    '�������������еĻ���ʱ��ͬʱ���¶�Ӧ���ݣ�����ͬ��
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
    '��ӡ�������
    
    Dim blnFlexCell As Boolean
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    
    
End Sub

Public Sub gsGridPrintPreview(ByRef gridControl As Control)   'FlexCell.Grid
    'Ԥ���������
    
    Dim blnFlexCell As Boolean
    Dim blnVSGrid As Boolean
    
    If TypeOf gridControl Is FlexCell.Grid Then blnFlexCell = True
    If TypeOf gridControl Is VSFlex8Ctl.VSFlexGrid Then blnVSGrid = True
    
    If blnFlexCell Then
        gridControl.PrintPreview
    End If
    
End Sub

Public Sub gsGridToExcel(ByRef gridControl As Control, Optional ByVal TimeCol As Long = -1, Optional ByVal TimeStyle As String = "yyyy-MM-dd HH:mm:ss")  '������Excel
    '�����ؼ��е����ݵ�����Excel��
    '����TimeCol��Ϊ�ؼ��е�ʱ���е��кţ�TimeStyle�趨��ʽ
    '�������Excel��������ʱ������Ӧ��MSOFFICE�����
    
'    Dim xlsOut As Excel.Application    '����������ñ�̵�Ҫ���ã�������ΪObject
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
        '������ݸ��Ƶ�Excel��
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
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Bold = True '�Ӵ���ʾ(��һ��Ĭ�ϱ�����)
        .Range(.Cells(1, 1), .Cells(1, C)).Font.Size = 12   '��һ��12���ִ�С
        .Range(.Cells(2, 1), .Cells(R, C)).Font.Size = 10   '�ڶ����Ժ�10���ִ�С
        .Range(.Cells(1, 1), .Cells(R, C)).HorizontalAlignment = -4108  'xlCenter= -4108(&HFFFFEFF4)   '������ʾ
        .Range(.Cells(1, 1), .Cells(R, C)).Borders.Weight = 2   'xlThin=2  '��Ԫ����ʾ��ɫ�߿�
        .Columns.EntireColumn.AutoFit   '�Զ��п�
        .Rows(1).rowHeight = 23 '��һ���и�
    End With
    
    xlsOut.Visible = True   '��ʾExcel�ĵ�
    
    Set sheetOut = Nothing
    Set xlsOut = Nothing
    Screen.MousePointer = 0
    
End Sub


Public Sub gsGridToText(ByRef gridControl As Control)
    '������ı��ؼ��е����ݵ���Ϊ�ı��ļ�
    
    Dim strFileName As String
    Dim blnFlexCell As Boolean
    Dim intFree As Integer
    Dim R As Long, C As Long, I As Long, J As Long
    Dim strTxt As String
    
    For I = 1 To 8
        strFileName = strFileName & gfBackOneChar(udNumber + udUpperCase) '�ļ����е�8������ַ�������Сд��ĸ
    Next
    strFileName = gID.FolderData & Format(Now, "yyyyMMddHHmmss_") & strFileName & ".txt"
    If Not gfFileRepair(strFileName) Then
        MsgBox "�����ļ�ʧ�ܣ������ԣ�", vbExclamation, "�ļ����ɾ���"
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
    
    Call gfFileOpen(strFileName)    '��
    
End Sub


Public Sub gsGridToWord(ByRef gridControl As Control)
    '��ָ������е����ݵ�����Word�ĵ���
    
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
    tbOut.Rows(1).Range.Bold = True             '��һ�����ݼӴ�
    tbOut.Range.ParagraphFormat.Alignment = 1   '������ݾ�����ʾ
    Call tbOut.AutoFitBehavior(1)               '���������Զ������п�
    
    wordApp.Visible = True
    
    Set tbOut = Nothing
    Set docOut = Nothing
    Set wordApp = Nothing
    
End Sub

Public Sub gsLoadAuthority(ByRef frmCur As Form, ByRef ctlCur As Control)
    '���ش����еĿ���Ȩ��
    
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
    '��Ӳ�����־
    
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
    '����Checked���Լ����仯
    
    If blnCheck Then    '=Falseʱ������
        Call gsNodeCheckUp(nodeCheck)
    End If
    
    Call gsNodeCheckDown(nodeCheck, blnCheck)
    
End Sub

Public Sub gsNodeCheckDown(ByRef nodeCheck As MSComctlLib.Node, Optional ByVal blnCheck As Boolean)
    '��/��ѡ���������ӽ��
    
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
    '��ѡ�������и����
    
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
    '��ָ������ģʽOpenMode�봰��FormWndState״̬����ָ������strFormName
    
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
    frmOpen.Show OpenMode               '�˾����󣬲��ܷ��Ͼ�ǰ�棬�����˳�����ʱMDI���岻����ȫ�رգ�������ΪCommandBars�ؼ���ԭ��
        
End Sub


Public Sub gsUnCheckedAction(ByVal strFormName As String)
    '�����ڹر�ʱ��ȥ����������cBS�ؼ��б���ѡ�Ķ�ӦAction
    
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


