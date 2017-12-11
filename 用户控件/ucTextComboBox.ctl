VERSION 5.00
Begin VB.UserControl TextCombo 
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   ScaleHeight     =   1890
   ScaleWidth      =   3615
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "TextCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API����
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'��������
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const conPropText As String = ""
Private Const conPropFontSize As Long = 9
Private Const conPropForeColor As Long = &H80000008
Private Const conPropBackColor As Long = &H80000005

'ö��
Enum AppearanceConstants
    ucFlat
    uc3D
End Enum

Enum BorderStyleConstants
    ucNone
    ucFixedSingle
End Enum

'��������������
Private fontProperty As New StdFont
Private mblnLocked As Boolean


'�¼�����
Public Event Change()
Public Event ClickDropDown()
Public Event DropDown()



'�������
Public Property Get Alignment() As AlignmentConstants
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal newAlignment As AlignmentConstants)
    Text1.Alignment = newAlignment
    PropertyChanged "Alignment"
End Property

Public Property Get Appearance() As AppearanceConstants
    Appearance = Combo1.Appearance
End Property

Public Property Let Appearance(ByVal newAppearance As AppearanceConstants)
    Text1.Appearance = newAppearance
    Combo1.Appearance = newAppearance
    PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = Combo1.BackColor
End Property

Public Property Let BackColor(newBackColor As OLE_COLOR)
    Text1.BackColor = newBackColor
    Combo1.BackColor = newBackColor
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As BorderStyleConstants
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal newBorderStyle As BorderStyleConstants)
    Text1.BorderStyle = newBorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Enabled() As Boolean
    Enabled = Combo1.Enabled
End Property

Public Property Let Enabled(ByVal newEnabled As Boolean)
    Text1.Enabled = newEnabled
    Combo1.Enabled = newEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = Combo1.Font
End Property

Public Property Set Font(ByRef newFont As StdFont)  '����������Property Set��Let
    Set Combo1.Font = newFont
    Set Text1.Font = newFont
    FontSize = newFont.Size 'ͬʱ�޸�FontSize����
    PropertyChanged "Font"
End Property

Public Property Get FontSize() As Long
    FontSize = Combo1.FontSize
End Property

Public Property Let FontSize(ByVal newFontSize As Long)
    Text1.FontSize = newFontSize
    Combo1.FontSize = newFontSize
    
    Call UserControl_Resize
    PropertyChanged "FontSize"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Combo1.ForeColor
End Property

Public Property Let ForeColor(newColor As OLE_COLOR)
    Text1.ForeColor = newColor
    Combo1.ForeColor = newColor
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
    Text1.Text = Combo1.Text
End Property

Public Property Get Locked() As Boolean
    Locked = mblnLocked
End Property

Public Property Let Locked(ByVal newLocked As Boolean)
    mblnLocked = newLocked
    PropertyChanged "Locked"
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "200"
    Text = Combo1.Text
End Property

Public Property Let Text(ByVal newText As String)
    Text1.Text = newText
    Combo1.Text = newText
    PropertyChanged "Text"
End Property



'��ӷ���
Public Sub AddItem(Item As String, Optional ByVal Index As Long = -1)
    Dim lngC As Long
    
    lngC = Combo1.ListCount

    If (Index < 0) Or (Index > lngC) Then  'Indexֵ�����ķ�����
        Index = lngC
    End If
    
    Combo1.AddItem Item, Index
    
End Sub

Public Sub Clear()
    Combo1.Clear
End Sub




'�ӿؼ��¼�
Private Sub Combo1_Change()
    RaiseEvent Change
End Sub

Private Sub Combo1_Click()
    Text1.Text = Combo1.Text
    Text1.ZOrder
    If Text1.Visible Then Text1.SetFocus
    Text1.SelStart = Len(Combo1.Text)
    
    RaiseEvent ClickDropDown
    
End Sub

Private Sub Combo1_DropDown()
    RaiseEvent DropDown
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If Locked Then KeyAscii = 0
End Sub

Private Sub Combo1_LostFocus()
    Text1.Text = Combo1.Text
    Text1.ZOrder
End Sub

Private Sub Text1_Click()
    Combo1.ZOrder
    
    Combo1.SetFocus
    Combo1.SelStart = Len(Combo1.Text)
    If Combo1.ListCount > 0 Then    '�ж��Ƿ񵯳������б�
        Call SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, 0)
    End If
    
End Sub

Private Sub Text1_GotFocus()
    Combo1.ZOrder
    Combo1.SetFocus
    Combo1.SelStart = 0
    Combo1.SelLength = Len(Combo1.Text)
End Sub



'�û��ؼ��¼�
Private Sub UserControl_Initialize()

    mblnLocked = True
    Text1.Move 0, 0
    Combo1.Move 0, 0
    Text1.ZOrder

End Sub

Private Sub UserControl_InitProperties()
    
    Dim ctlUC As Control
    Dim strName As String

    '���Զ������������б��������Զ���ؼ���
    '�ҵ����һ���Զ���ؼ�����ȡ��Nameֵ����
    For Each ctlUC In UserControl.Parent.Controls
        If TypeOf ctlUC Is TextCombo Then
            strName = ctlUC.Name
        End If
    Next

    Text = strName  'Ȼ��Text����Ĭ��ֵ��Ϊ�ؼ���Nameֵ

End Sub

Private Sub UserControl_LostFocus()
    Text1.ZOrder
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
    Appearance = PropBag.ReadProperty("Appearance", uc3D)
    BackColor = PropBag.ReadProperty("BackColor", conPropBackColor)
    BorderStyle = PropBag.ReadProperty("BorderStyle", ucFixedSingle)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Set Font = PropBag.ReadProperty("Font", fontProperty)
    FontSize = PropBag.ReadProperty("FontSize", conPropFontSize)
    ForeColor = PropBag.ReadProperty("ForeColor", conPropForeColor)
    Locked = PropBag.ReadProperty("Locked", True)
    Text = PropBag.ReadProperty("Text", conPropText)
    
End Sub

Private Sub UserControl_Resize()

    With UserControl
        .Height = Combo1.Height
        Combo1.Width = .Width
        Text1.Height = .Height
        Text1.Width = .Width
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    PropBag.WriteProperty "Alignment", Alignment, vbLeftJustify
    PropBag.WriteProperty "Appearance", Appearance, uc3D
    PropBag.WriteProperty "BackColor", BackColor, conPropBackColor
    PropBag.WriteProperty "BorderStyle", BorderStyle, ucFixedSingle
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "Font", Font, fontProperty
    PropBag.WriteProperty "FontSize", FontSize, conPropFontSize
    PropBag.WriteProperty "ForeColor", ForeColor, conPropForeColor
    PropBag.WriteProperty "Locked", Locked, True
    PropBag.WriteProperty "Text", Text, conPropText

End Sub
