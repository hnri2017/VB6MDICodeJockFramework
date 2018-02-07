VERSION 5.00
Begin VB.Form frmSysGetPCInfo 
   Caption         =   "��ȡ���������Ϣ"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15915
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   15915
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "Ӳ�̡��ڴ���Ϣ"
      Height          =   615
      Left            =   9720
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ȡӲ����Ϣ"
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
    strTemp = "���������Ϣ��" & vbCrLf
    
    On Error Resume Next
    
    'WMI��ȫ����Windows Management Instrumentation����Windows������
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    
    '��ѯ������Ϣ
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_BaseBoard")
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "�����ͺţ�" & objItem.Product
        strTemp = strTemp & vbCrLf & "�������кţ�" & objItem.SerialNumber
        strTemp = strTemp & vbCrLf & "���������̣�" & objItem.Manufacturer
    Next
    
    '��ѯCPU��Ϣ
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_Processor")
    strTemp = strTemp & vbCrLf
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "CPU���ƣ�" & objItem.Name '��һ���Ӱ���ٶ�
        strTemp = strTemp & vbCrLf & "CPU���кţ�" & objItem.ProcessorId
        strTemp = strTemp & vbCrLf & "CPU����/�߼���������" & objItem.NumberOfCores & "." & objItem.NumberOfLogicalProcessors
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
                strText = "δ֪"
        End Select
        strTemp = strTemp & vbCrLf & "CPU���ܣ�" & strText
        strTemp = strTemp & vbCrLf & "CPU��ǰ��Ƶ��" & objItem.CurrentClockSpeed & "MHz"
        strTemp = strTemp & vbCrLf & "CPU��Ƶ��" & objItem.ExtClock & "MHz"
        strTemp = strTemp & vbCrLf & "ϵͳ���ͣ�" & objItem.AddressWidth & "λ����ϵͳ"
    Next
    
    '��ѯӲ����Ϣ
    '���Ȳ�ѯϵͳ������Ӳ��ID
    Set objMain = objWMI.ExecQuery("SELECT DiskIndex FROM Win32_DiskPartition WHERE Bootable = TRUE")
    For Each objItem In objMain
        strText = objItem.DiskIndex
    Next
    '��ѯϵͳ����Ӳ�̵���Ϣ
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_DiskDrive WHERE Index = " & strText)
    strTemp = strTemp & vbCrLf
    strTemp = strTemp & vbCrLf & "��Ӳ��Index��" & strText
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "��Ӳ���ͺţ�" & objItem.Model
        strTemp = strTemp & vbCrLf & "��Ӳ�����кţ�" & objItem.SerialNumber
        strTemp = strTemp & vbCrLf & "��Ӳ�̽ӿ����ͣ�" & objItem.InterfaceType
        strTemp = strTemp & vbCrLf & "��Ӳ�̴�С��" & CLng(Val(objItem.Size) / 1024 / 1024 / 1024) & "GB"
    Next

    '��ѯ�ڴ���Ϣ
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
    strText = ""
    C = 0
    For Each objItem In objMain
        C = C + Val(objItem.Capacity)
        strText = objItem.Manufacturer
    Next
    strTemp = strTemp & vbCrLf
    strTemp = strTemp & vbCrLf & "�ڴ������̣�" & strText
    strTemp = strTemp & vbCrLf & "�ڴ��С��" & CLng(Val(C) / 1024 / 1024 / 1024) & "GB"
    
    '�Կ�
    Set objMain = objWMI.ExecQuery("SELECT * FROM win32_VideoController")
    strTemp = strTemp & vbCrLf
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "�Կ����ƣ�" & objItem.Description
        strTemp = strTemp & vbCrLf & "�Կ��ֱ��ʣ�" & objItem.VideoModeDescription
    Next
    
    '��������ѯʱ��Ҫ5��
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_NetworkAdapter")
    strTemp = strTemp & vbCrLf
    For Each objItem In objMain
        If Left(objItem.NetConnectionID, 4) = "��������" Then
            strTemp = strTemp & vbCrLf & "�������ƣ�" & objItem.NetConnectionID & "----" & objItem.Name
            Exit For
        End If
    Next
    
    'ϵͳ
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    strTemp = strTemp & vbCrLf
    For Each objItem In objMain
        strTemp = strTemp & vbCrLf & "��������ƣ�" & objItem.Name
        strTemp = strTemp & vbCrLf & "��ǰ�û���" & objItem.UserName
    Next
    
    strTemp = strTemp & vbCrLf & vbCrLf & "�ܹ���ʱ" & CStr(Timer - sngTime) & "��"
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
    
    '��ѯӲ����Ϣ
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_DiskDrive")
    strTemp = strTemp & vbCrLf
    K = 0: N = 0: strText = ""
    For Each objItem In objMain
        K = K + 1
        C = C + Val(objItem.Size)
    Next
    If K > 1 Then strTemp = strTemp & vbCrLf & "Ӳ��" & strText & "��С��" & CLng(Val(C) / 1024 / 1024 / 1024) & "GB"
    For Each objItem In objMain
        If K > 1 Then
            N = N + 1
            strText = N
        End If
        strTemp = strTemp & vbCrLf & "Ӳ��" & strText & "�ͺţ�" & objItem.Model
        strTemp = strTemp & vbCrLf & "Ӳ��" & strText & "���кţ�" & Trim(objItem.SerialNumber)
        strTemp = strTemp & vbCrLf & "Ӳ��" & strText & "�ӿ����ͣ�" & objItem.InterfaceType
        strTemp = strTemp & vbCrLf & "Ӳ��" & strText & "��С��" & CLng(Val(objItem.Size) / 1024 / 1024 / 1024) & "GB"
        strTemp = strTemp & vbCrLf & "Ӳ��" & strText & "��������" & objItem.Partitions
    Next

    '��ѯ�ڴ���Ϣ
    Set objMain = objWMI.ExecQuery("SELECT * FROM Win32_PhysicalMemory")
    C = 0: K = 0: N = 0: strText = ""
    For Each objItem In objMain
        K = K + 1
        C = C + Val(objItem.Capacity)
    Next
    strTemp = strTemp & vbCrLf
    If K > 1 Then strTemp = strTemp & vbCrLf & "�ڴ��ܴ�С��" & CLng(Val(C) / 1024 / 1024 / 1024) & "GB"
    For Each objItem In objMain
        If K > 1 Then
            N = N + 1
            strText = N
        End If
        strTemp = strTemp & vbCrLf & "�ڴ�" & strText & "�����̣�" & objItem.Manufacturer
        strTemp = strTemp & vbCrLf & "�ڴ�" & strText & "��С��" & CLng(Val(objItem.Capacity) / 1024 / 1024 / 1024) & "GB"
    Next

    Text1.Text = strTemp
    Set objItem = Nothing
    Set objMain = Nothing
    Set objWMI = Nothing
    
End Sub
