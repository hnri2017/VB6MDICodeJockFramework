VERSION 5.00
Object = "{E08BA07E-6463-4EAB-8437-99F08000BAD9}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmForm3 
   Caption         =   "²âÊÔ´°¿Ú3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin FlexCell.Grid Grid1 
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2566
      Cols            =   5
      GridColor       =   12632256
      Rows            =   30
   End
End
Attribute VB_Name = "frmForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Set Me.Icon = gMDI.imgListCommandBars.ListImages("SysPassword").Picture
End Sub
