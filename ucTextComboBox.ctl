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

Private Sub UserControl_Initialize()
    
    Text1.Text = ""
    Combo1.Text = ""
    
    Text1.Move 0, 0
    Combo1.Move 0, 0
    
    Text1.ZOrder
    
End Sub

Private Sub UserControl_Resize()

    With UserControl
        .Height = Combo1.Height
        Combo1.Width = .Width
        Text1.Height = .Height
        Text1.Width = .Width
    End With
    
End Sub


