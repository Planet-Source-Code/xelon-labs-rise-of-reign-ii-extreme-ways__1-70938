VERSION 5.00
Begin VB.UserControl LoadBar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label lblshad 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00481A1D&
      Height          =   255
      Left            =   495
      TabIndex        =   1
      Top             =   15
      Width           =   4575
   End
   Begin VB.Shape brdr 
      BorderColor     =   &H00FFFFFF&
      Height          =   1695
      Left            =   960
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Image Glass 
      Height          =   300
      Left            =   0
      Picture         =   "LoadBar.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "LoadBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_prec As Integer
Dim m_by As Integer
Dim i As Integer
Dim rand As Long
Dim Color As Long
Event Change()

Sub SetPro(Prec As Integer, by As Integer)
On Error Resume Next
Dim np As Integer
RaiseEvent Change
DoEvents
Glass.Width = (((Prec / by) * 100) / 100) * Width
np = (Glass.Width / Width) * 100
m_prec = Prec
m_by = by
DoEvents
End Sub

Private Sub UserControl_Initialize()
Randomize
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Glass.Height
brdr.Left = 0
brdr.Top = 0
brdr.Width = UserControl.Width
brdr.Height = UserControl.Height
SetPro m_prec, m_by
End Sub

Sub Set_Text(str As String)
Lbl = str
lblshad = str
Refresh
RaiseEvent Change
End Sub
