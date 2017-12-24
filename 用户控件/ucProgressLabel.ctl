VERSION 5.00
Begin VB.UserControl ProgressLabel 
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   ScaleHeight     =   1605
   ScaleWidth      =   4530
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   180
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Label2"
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label1"
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "ProgressLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'常量
Private Const mconBackColorBack As Long = &H80FFFF
Private Const mconBackColorFore As Long = &HFF00&
Private Const mconForeColor As Long = &H80000012
Private Const mconMax As Long = 100
Private Const mconMin As Long = 0
Private Const mconValue As Long = 50

'变量
Dim mFont As New StdFont
Dim mlngMax As Long
Dim mlngMin As Long
Dim mlngVal As Long

'方法
Private Sub msReadValue(ByVal lngVal As Long)
    
End Sub

'属性定义
Public Property Get BackColorBack() As OLE_COLOR
    BackColorBack = Label1.BackColor
End Property

Public Property Let BackColorBack(newColor As OLE_COLOR)
    Label1.BackColor = newColor
    PropertyChanged "BackColorBack"
End Property

Public Property Get BackColorFore() As OLE_COLOR
    BackColorFore = Label2.BackColor
End Property

Public Property Let BackColorFore(newColor As OLE_COLOR)
    Label2.BackColor = newColor
    PropertyChanged "BackColorFore"
End Property

Public Property Get Font() As StdFont
    Set Font = Label3.Font
End Property

Public Property Set Font(newFont As StdFont)
    Set Label3.Font = newFont
    PropertyChanged "Font"
    Call UserControl_Resize
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Label3.ForeColor
End Property

Public Property Let ForeColor(newColor As OLE_COLOR)
    Label3.ForeColor = newColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Max() As Long
    Max = mlngMax
End Property

Public Property Let Max(newMax As Long)
    If newMax <= Min Then newMax = Min + 1
    mlngMax = newMax
    PropertyChanged "Max"
End Property

Public Property Get Min() As Long
    Min = mlngMin
End Property

Public Property Let Min(newMin As Long)
    mlngMin = newMin
    PropertyChanged "Min"
End Property

Public Property Get Value() As Long
    Value = mlngVal
End Property

Public Property Let Value(ByVal newVal As Long)
    Dim sngPer As Single
    If newVal < Min Or newVal > Max Then newVal = Max
    mlngVal = newVal
    sngPer = newVal / (Max - Min)
    Label3.Caption = CLng(sngPer * 100) & "%"
    Label2.Width = UserControl.Width * sngPer
    PropertyChanged "Value"
End Property


'UserControl事件
Private Sub UserControl_Initialize()
    Label1.Caption = ""
    Label2.Caption = ""
    mFont.Size = 9
    mlngMax = mconMax
    mlngVal = mconMax
    Label3.Caption = "100%"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '
    BackColorBack = PropBag.ReadProperty("BackColorBack", mconBackColorBack)
    BackColorFore = PropBag.ReadProperty("BackColorFore", mconBackColorFore)
    ForeColor = PropBag.ReadProperty("ForeColor", mconForeColor)
    Set Font = PropBag.ReadProperty("Font", mFont)
    Max = PropBag.ReadProperty("Max", mconMax)
    Min = PropBag.ReadProperty("Min", mconMin)
    Value = PropBag.ReadProperty("Value", mconMax)
    
End Sub

Private Sub UserControl_Resize()
    Label1.Move 0, 0, UserControl.Width, UserControl.Height
    Label2.Move 0, 0, UserControl.Width * (Value / (Max - Min)), UserControl.Height
    With Label3
        .AutoSize = True
        .Move 0, (UserControl.Height - .Height) / 2, UserControl.Width
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    PropBag.WriteProperty "BackColorBack", BackColorBack, mconBackColorBack
    PropBag.WriteProperty "BackColorFore", BackColorFore, mconBackColorFore
    PropBag.WriteProperty "ForeColor", ForeColor, mconForeColor
    PropBag.WriteProperty "Font", Font, mFont
    PropBag.WriteProperty "Max", Max, mconMax
    PropBag.WriteProperty "Min", Min, mconMin
    PropBag.WriteProperty "Value", Value, mconMax
    
End Sub
