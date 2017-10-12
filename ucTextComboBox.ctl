VERSION 5.00
Begin VB.UserControl ucTextComboBox 
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
Attribute VB_Name = "ucTextComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
Private Const CB_SHOWDROPDOWN As Long = &H14F


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

Public Property Get Text() As String
    Text = Combo1.Text
End Property

Public Property Let Text(ByVal newText As String)
    Text1.Text = newText
    Combo1.Text = newText
    PropertyChanged "Text"
End Property




Public Sub AddItem(Item As String, Optional ByVal Index As Long)
    Dim lngC As Long
    
    lngC = Combo1.ListCount
    If Index <> 0 Then
        If lngC = 0 Then
            Index = 0
        ElseIf (Index < 0) Or (Index > lngC) Then
            Index = lngC
        End If
    End If
    Combo1.AddItem Item, Index
    
End Sub

Private Sub Combo1_Click()
    Text1.Text = Combo1.Text
    Text1.ZOrder
    Text1.SetFocus
    Text1.SelStart = Len(Combo1.Text)
End Sub

Private Sub Combo1_LostFocus()
    Text1.Text = Combo1.Text
    Text1.ZOrder
End Sub

Private Sub Text1_Click()
    Combo1.ZOrder
    
    Combo1.SetFocus
    Combo1.SelStart = Len(Combo1.Text)
    If Combo1.ListCount > 0 Then
        Call SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, 1, 0)
    End If
    
End Sub

Private Sub UserControl_Initialize()
    
    Text1.Text = ""
    Combo1.Text = ""
    
    Text1.Move 0, 0
    Combo1.Move 0, 0
    
    Text1.ZOrder

    Text = UserControl.Name
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Error Resume Next
    FontSize = PropBag.ReadProperty("FontSize", Combo1.FontSize)
    ForeColor = PropBag.ReadProperty("ForeColor", Combo1.ForeColor)
    Text = PropBag.ReadProperty("Text", Combo1.Text)
    
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

    PropBag.WriteProperty "FontSize", FontSize, 9
    PropBag.WriteProperty "ForeColor", ForeColor, &H80000008
    PropBag.WriteProperty "Text", Text, ""
    
End Sub
