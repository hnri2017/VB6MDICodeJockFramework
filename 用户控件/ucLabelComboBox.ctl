VERSION 5.00
Begin VB.UserControl LabelCombo 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "用户控件"
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "LabelCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API声明
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long


'常量
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const conPropCaption As String = ""
Private Const conPropBackColor As Long = &H8000000F
Private Const conPropForeColor As Long = &H80000012

''枚举
'Enum AppearanceConstants
'    ucFlat
'    uc3D
'End Enum
'
'Enum BorderStyleConstants
'    ucNone
'    ucFixedSingle
'End Enum

'变量
Private mFont As New StdFont
Private mblnLocked As Boolean


'事件声明
Public Event Change()
Public Event ClickDropDown()
Public Event DropDown()



'自定义过程
Private Sub msAutoSize()
    '当控件的AutoSzie属性为True时，应实时更新控件的宽度
    Label1.AutoSize = True
    Label1.Height = Combo1.Height
    UserControl.Width = Label1.Width
    Combo1.Width = Label1.Width
End Sub




'属性
Public Property Get Alignment() As AlignmentConstants
    Alignment = Label1.Alignment
End Property

Public Property Let Alignment(ByVal newAlignment As AlignmentConstants)
    Label1.Alignment = newAlignment
    PropertyChanged "Alignment"
End Property

Public Property Get Appearance() As AppearanceConstants
    Appearance = Label1.Appearance
End Property

Public Property Let Appearance(ByVal newAppearance As AppearanceConstants)
    Label1.Appearance = newAppearance
    PropertyChanged "Appearance"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = Label1.AutoSize
End Property

Public Property Let AutoSize(ByVal newAutoSize As Boolean)
    Label1.AutoSize = newAutoSize
    If AutoSize Then Call msAutoSize
    PropertyChanged "AutoSize"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = Label1.BackColor
End Property

Public Property Let BackColor(ByVal newBackColor As OLE_COLOR)
    Label1.BackColor = newBackColor
    Combo1.BackColor = newBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As BorderStyleConstants
    BorderStyle = Label1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal newBackStyle As BorderStyleConstants)
    Label1.BorderStyle = newBackStyle
    If AutoSize Then Call msAutoSize
    PropertyChanged "BorderStyle"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
    Caption = Combo1.Text
End Property

Public Property Let Caption(ByVal newCaption As String)
    Label1.Caption = newCaption
    Combo1.Text = newCaption
    If AutoSize Then Call msAutoSize
End Property

Public Property Get Enabled() As Boolean
    Enabled = Label1.Enabled
End Property

Public Property Let Enabled(ByVal newEnabled As Boolean)
    Label1.Enabled = newEnabled
    Combo1.Enabled = newEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = Combo1.Font
End Property

Public Property Set Font(ByVal newFont As StdFont)
    Set Label1.Font = newFont
    Set Combo1.Font = newFont
    If AutoSize Then Call msAutoSize
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal newForeColor As OLE_COLOR)
    Label1.ForeColor = newForeColor
    Combo1.ForeColor = newForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get List(ByVal Index As Long) As String
    List = Combo1.List(Index)
End Property

Public Property Get ListCount() As Long
    ListCount = Combo1.ListCount
End Property

Public Property Get ListIndex() As Long
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = Combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal newListIndex As Long)
    Combo1.ListIndex = newListIndex
    Label1.Caption = Combo1.Text
    If AutoSize Then Call msAutoSize
End Property

Public Property Get Locked() As Boolean
    Locked = mblnLocked
End Property

Public Property Let Locked(ByVal newLocked As Boolean)
    mblnLocked = newLocked
    PropertyChanged "Locked"
End Property



'添加方法
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Long = -1)
    
    Dim lngLC As Long
    
    With Combo1
        lngLC = .ListCount
        If (Index < 0) Or (Index > lngLC) Then
            Index = lngLC
        End If
        .AddItem Item, Index
    End With
    
End Sub

Public Sub Clear()
    Combo1.Clear
End Sub



'子控件事件
Private Sub Combo1_Change()
    RaiseEvent Change
End Sub

Private Sub Combo1_Click()
    Combo1.Visible = False
    Label1.Caption = Combo1.Text
    If AutoSize Then Call msAutoSize
    
    RaiseEvent ClickDropDown
    
End Sub

Private Sub Combo1_DropDown()
    RaiseEvent DropDown
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If Locked Then KeyAscii = 0
End Sub

Private Sub Combo1_LostFocus()
    Combo1.Visible = False
    Label1.Caption = Combo1.Text
    If AutoSize Then Call msAutoSize
End Sub

Private Sub Label1_Click()
    With Combo1
        .Visible = True
        .SetFocus
        .SelLength = Len(.Text)
        If .ListCount > 0 Then
            Call SendMessage(.hwnd, CB_SHOWDROPDOWN, 1, 0)
        End If
    End With
End Sub


'用户控件事件
Private Sub UserControl_Initialize()
    
    mblnLocked = True
    Label1.Move 0, 0
    Combo1.Move 0, 0
    Combo1.Visible = False
    
End Sub

Private Sub UserControl_InitProperties()
    '拖入控件后将控件Name值赋给Caption值
    
    Dim ctlUC As Control
    Dim strName As String
    
    For Each ctlUC In UserControl.Parent.Controls
        If TypeOf ctlUC Is LabelCombo Then
            strName = ctlUC.Name
        End If
    Next
    
    Caption = strName
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    '
    Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
    Appearance = PropBag.ReadProperty("Appearance", uc3D)
    AutoSize = PropBag.ReadProperty("AutoSize", False)
    BackColor = PropBag.ReadProperty("BackColor", conPropBackColor)
    BorderStyle = PropBag.ReadProperty("BorderStyle", ucNone)
    Caption = PropBag.ReadProperty("Caption", conPropCaption)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", mFont)
    ForeColor = PropBag.ReadProperty("ForeColor", conPropForeColor)
    Locked = PropBag.ReadProperty("Locked", True)
    
End Sub

Private Sub UserControl_Resize()
    '
    With UserControl
        .Height = Combo1.Height
        Label1.Height = Combo1.Height
        
        Combo1.Width = .Width
        Label1.Width = .Width
    End With
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    '
    PropBag.WriteProperty "Alignment", Alignment, vbLeftJustify
    PropBag.WriteProperty "Appearance", Appearance, uc3D
    PropBag.WriteProperty "AutoSize", AutoSize, False
    PropBag.WriteProperty "BackColor", BackColor, conPropBackColor
    PropBag.WriteProperty "BorderStyle", BorderStyle, ucNone
    PropBag.WriteProperty "Caption", Caption, conPropCaption
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Font", Font, mFont
    PropBag.WriteProperty "ForeColor", ForeColor, conPropForeColor
    PropBag.WriteProperty "Locked", Locked, True

End Sub
