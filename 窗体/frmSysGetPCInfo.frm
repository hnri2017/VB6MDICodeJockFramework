VERSION 5.00
Begin VB.Form frmSysGetPCInfo 
   Caption         =   "获取电脑相关信息"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15915
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   15915
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "硬盘、内存信息"
      Height          =   615
      Left            =   9720
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "获取硬件信息"
      Height          =   615
      Left            =   9720
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   8655
   End
End
Attribute VB_Name = "frmSysGetPCInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function mfGetPCInfo() As String
    
    Dim objWMI As Object, objMain As Object, objItem As Object
    Dim strText As String, strTemp As String
    Dim K As Long, C As Double
    Dim sngTime As Single
    
    sngTime = Timer
    Me.MousePointer = 13
    strTemp = "电脑相关信息：" & vbCrLf
    
    On Error Resume Next
    
    'WMI的全称是Windows Management Instrumentation，即Windows管理工具
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    
    '查询主板信息
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_BaseBoard")
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "主板型号：" & objItem.Product
        strTemp = strTemp & vbCrLf & "主板序列号：" & objItem.SerialNumber
        strTemp = strTemp & vbCrLf & "主板制造商：" & objItem.Manufacturer
    Next
    
    '查询CPU信息
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_Processor")
    strTemp = strTemp & vbCrLf
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "CPU名称：" & objItem.Name '这一句很影响速度
        strTemp = strTemp & vbCrLf & "CPU序列号：" & objItem.ProcessorId
        strTemp = strTemp & vbCrLf & "CPU物理/逻辑核心数：" & objItem.NumberOfCores & "." & objItem.NumberOfLogicalProcessors
        Select Case objItem.Architecture
            Case 0
                strText = "X86"
            Case 1
                strText = "MIPS"
            Case 2
                strText = "Alpha"
            Case 3
                strText = "PowerPC"
            Case 5
                strText = "ARM"
            Case 6
                strText = "Itanium-based systems"
            Case 9
                strText = "x64"
            Case Else
                strText = "未知"
        End Select
        strTemp = strTemp & vbCrLf & "CPU构架：" & strText
        strTemp = strTemp & vbCrLf & "CPU当前主频：" & objItem.CurrentClockSpeed & "MHz"
        strTemp = strTemp & vbCrLf & "CPU外频：" & objItem.ExtClock & "MHz"
        strTemp = strTemp & vbCrLf & "系统类型：" & objItem.AddressWidth & "位操作系统"
    Next
    
    '查询硬盘信息
    '首先查询系统盘所有硬盘ID
    Set objMain = objWMI.ExecQuery("SELECT DiskIndex FROM Win32_DiskPartition WHERE Bootable = TRUE")
    For Each objItem In objMain
        strText = objItem.DiskIndex
    Next
    '查询系统所在硬盘的信息
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_DiskDrive WHERE Index = " & strText)
    strTemp = strTemp & vbCrLf
    strTemp = strTemp & vbCrLf & "主硬盘Index：" & strText
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "主硬盘型号：" & objItem.Model
        strTemp = strTemp & vbCrLf & "主硬盘序列号：" & objItem.SerialNumber
        strTemp = strTemp & vbCrLf & "主硬盘接口类型：" & objItem.InterfaceType
        strTemp = strTemp & vbCrLf & "主硬盘大小：" & CLng(Val(objItem.Size) / 1024 / 1024 / 1024) & "GB"
    Next

    '查询内存信息
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
    strText = ""
    C = 0
    For Each objItem In objMain
        C = C + Val(objItem.Capacity)
        strText = objItem.Manufacturer
    Next
    strTemp = strTemp & vbCrLf
    strTemp = strTemp & vbCrLf & "内存制造商：" & strText
    strTemp = strTemp & vbCrLf & "内存大小：" & CLng(Val(C) / 1024 / 1024 / 1024) & "GB"
    
    '显卡
    Set objMain = objWMI.ExecQuery("SELECT * FROM win32_VideoController")
    strTemp = strTemp & vbCrLf
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "显卡名称：" & objItem.Description
        strTemp = strTemp & vbCrLf & "显卡分辨率：" & objItem.VideoModeDescription
    Next
    
    '网卡。查询时间要5秒
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
    strTemp = strTemp & vbCrLf
    For Each objItem In objMain
        If Left(objItem.NetConnectionID, 4) = "本地连接" Then
            strTemp = strTemp & vbCrLf & "网卡名称：" & objItem.NetConnectionID & "----" & objItem.Name
            Exit For
        End If
    Next
    
    '系统
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    strTemp = strTemp & vbCrLf
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "计算机名称：" & objItem.Name
        strTemp = strTemp & vbCrLf & "当前用户：" & objItem.UserName
    Next
    
    strTemp = strTemp & vbCrLf & vbCrLf & "总共用时" & CStr(Timer - sngTime) & "秒"
    mfGetPCInfo = strTemp
    Me.MousePointer = 0
    Set objItem = Nothing
    Set objMain = Nothing
    Set objWMI = Nothing
    
End Function

Private Sub Command1_Click()
    Text1.Text = mfGetPCInfo
End Sub

Private Sub Command2_Click()
    Dim objWMI As Object, objMain As Object, objItem As Object
    Dim strText As String, strTemp As String
    Dim K As Long, N As Long, C As Double
    
    On Error Resume Next
    
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    
    '查询硬盘信息
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_DiskDrive")
    strTemp = strTemp & vbCrLf
    K = 0: N = 0: strText = ""
    For Each objItem In objMain
        K = K + 1
        C = C + Val(objItem.Size)
    Next
    If K > 1 Then strTemp = strTemp & vbCrLf & "硬盘" & strText & "大小：" & CLng(Val(C) / 1024 / 1024 / 1024) & "GB"
    For Each objItem In objMain
        If K > 1 Then
            N = N + 1
            strText = N
        End If
        strTemp = strTemp & vbCrLf & "硬盘" & strText & "型号：" & objItem.Model
        strTemp = strTemp & vbCrLf & "硬盘" & strText & "序列号：" & Trim(objItem.SerialNumber)
        strTemp = strTemp & vbCrLf & "硬盘" & strText & "接口类型：" & objItem.InterfaceType
        strTemp = strTemp & vbCrLf & "硬盘" & strText & "大小：" & CLng(Val(objItem.Size) / 1024 / 1024 / 1024) & "GB"
        strTemp = strTemp & vbCrLf & "硬盘" & strText & "分区数：" & objItem.Partitions
    Next

    '查询内存信息
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
    C = 0: K = 0: N = 0: strText = ""
    For Each objItem In objMain
        K = K + 1
        C = C + Val(objItem.Capacity)
    Next
    strTemp = strTemp & vbCrLf
    If K > 1 Then strTemp = strTemp & vbCrLf & "内存总大小：" & CLng(Val(C) / 1024 / 1024 / 1024) & "GB"
    For Each objItem In objMain
        If K > 1 Then
            N = N + 1
            strText = N
        End If
        strTemp = strTemp & vbCrLf & "内存" & strText & "制造商：" & objItem.Manufacturer
        strTemp = strTemp & vbCrLf & "内存" & strText & "大小：" & CLng(Val(objItem.Capacity) / 1024 / 1024 / 1024) & "GB"
    Next

    Text1.Text = strTemp
    Set objItem = Nothing
    Set objMain = Nothing
    Set objWMI = Nothing
    
End Sub
