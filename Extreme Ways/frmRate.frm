VERSION 5.00
Begin VB.Form frmRate 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1335
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   1800
   End
   Begin VB.Label ldng 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Content Rated By Entertainment Software Rating Board"
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
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   120
      Picture         =   "frmRate.frx":0000
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
On Error Resume Next
Dim i As Double
FormOnTop Me
MakeTransparent hwnd, 0
Me.Show
Play_Sound "ion cannon.wav"
For i = 0 To 255 Step 0.5
MakeTransparent hwnd, Val(i)
Image1.Width = Val(i) * 2.9411
DoEvents
Next
ldng.Visible = True

MakeTransparent frm8X.hwnd, 0
FormOnTop frm8X
frm8X.Show

For i = 255 To 0 Step -0.5
MakeTransparent hwnd, Val(i)
MakeTransparent frm8X.hwnd, 255 - Val(i)
ldng.Left = (Val(i) * 7.0588) + 960
DoEvents
Next

Unload Me

For i = 255 To 0 Step -0.5
MakeTransparent frm8X.hwnd, Val(i)
DoEvents
Next
Unload frm8X
frmmnu.Show
End Sub
