VERSION 5.00
Begin VB.UserControl SPs 
   BackColor       =   &H00F0E7E3&
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   Begin VB.Timer Tmr_Nuke 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1680
   End
   Begin VB.Timer Tmr_Ion 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   480
   End
   Begin RoR.Bar Bar_Nuke 
      Height          =   45
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   79
   End
   Begin RoR.Bar Bar_Ion 
      Height          =   45
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   79
   End
   Begin VB.Label time2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label time1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ready"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label UD 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "\/"
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image Img_Nuke 
      Height          =   960
      Left            =   120
      Top             =   1440
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Img_Ion 
      Height          =   960
      Left            =   120
      Top             =   240
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image Nuke 
      Height          =   1200
      Left            =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Ion 
      Height          =   1200
      Left            =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "SPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Special Weapons Data
Public IC_time As Integer
Public IC_enabled As Boolean

Public Nuke_time As Integer
Public Nuke_enabled As Boolean

Public Nuke_Ready As Boolean
Public IC_Ready As Boolean

Public inite As Integer

Event ReadyIon()
Event ReadyNuke()
Event ClickIon()
Event ClickNuke()
Event Over(init As Integer)

Private Sub Img_Ion_Click()
RaiseEvent ClickIon
End Sub

Private Sub Img_Ion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Over(1)
End Sub

Private Sub Img_Nuke_Click()
RaiseEvent ClickNuke
End Sub

Private Sub Img_Nuke_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Over(2)
End Sub

Private Sub Tmr_Ion_Timer()
If IC_enabled = True Then
IC_time = IC_time - 1
Bar_Ion.SetPro IC_time, 500
time1 = IC_time
End If
If IC_time < 0 Then
Tmr_Ion = False
IC_Ready = True
time1 = "Ready"
Play_Sound "Nuke.wav"
RaiseEvent ReadyIon
End If
End Sub

Private Sub Tmr_Nuke_Timer()
If Nuke_enabled = True Then
Nuke_time = Nuke_time - 1
Bar_Nuke.SetPro Nuke_time, 400
time2 = Nuke_time
End If
If Nuke_time < 0 Then
Tmr_Nuke = False
Nuke_Ready = True
time2 = "Ready"
Play_Sound "Nuke.wav"
RaiseEvent ReadyNuke
End If
End Sub

Private Sub UD_Click()
If UD = "/\" Then
Height = 9 * 15
UD = "\/"
ElseIf UD = "\/" Then

If IC_enabled = True And Nuke_enabled = False Then
Height = (Ion.Height * 15) + (9 * 15)
ElseIf Nuke_enabled = True Then
Height = (Ion.Height * 30) + (9 * 15)
ElseIf IC_enabled = False And Nuke_enabled = False Then
Height = 9 * 15
End If

UD = "/\"
End If
End Sub

Private Sub UserControl_Initialize()
Set Nuke.Picture = LoadPicture(App.path & "\images\buildbar\Powers.gif")
Set Ion.Picture = LoadPicture(App.path & "\images\buildbar\Powers.gif")
Set Img_Ion.Picture = LoadPicture(App.path & "\images\buildbar\IonLaser.gif")
Set Img_Nuke.Picture = LoadPicture(App.path & "\images\buildbar\Nuke.gif")
End Sub

Private Sub UserControl_Resize()
Width = Ion.Width * 15
End Sub

Sub Activate_Nuke()
Nuke_time = 400
Nuke_enabled = True
Tmr_Nuke = True
Nuke_Ready = False
Img_Nuke.Visible = True
Nuke.Visible = True
Bar_Nuke.Visible = True
time2.Visible = True
End Sub

Sub Activate_IC()
IC_time = 500
IC_enabled = True
Tmr_Ion = True
IC_Ready = False
Img_Ion.Visible = True
Ion.Visible = True
Bar_Ion.Visible = True
time1.Visible = True
End Sub

Sub DeActivate_Nuke()
Nuke_time = 400
Tmr_Nuke = False
Nuke_Ready = False
Nuke_enabled = False
Img_Nuke.Visible = False
Nuke.Visible = False
Bar_Nuke.Visible = False
time2.Visible = False
End Sub

Sub DeActivate_IC()
IC_time = 500
Tmr_Ion = False
IC_Ready = False
IC_enabled = False
Img_Ion.Visible = False
Ion.Visible = False
Bar_Ion.Visible = False
time1.Visible = False
End Sub
