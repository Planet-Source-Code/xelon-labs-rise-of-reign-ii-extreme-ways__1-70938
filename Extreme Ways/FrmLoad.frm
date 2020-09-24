VERSION 5.00
Begin VB.Form FrmLoad 
   BackColor       =   &H00592D2B&
   BorderStyle     =   0  'None
   Caption         =   "RoR-II Loading"
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox imgview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00592D2B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   4200
      ScaleHeight     =   720
      ScaleWidth      =   750
      TabIndex        =   12
      Top             =   4080
      Width           =   750
   End
   Begin VB.Timer tmrload 
      Interval        =   1000
      Left            =   600
      Top             =   5760
   End
   Begin RoR.LoadBar LoadBar 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   6840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   529
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   3120
      Width           =   4775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   525
      Left            =   3960
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7shad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   525
      Left            =   3960
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   375
      Left            =   975
      TabIndex        =   8
      Top             =   5040
      Width           =   4775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tech Level :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   4680
      Width           =   4775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Power :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   4320
      Width           =   4775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3960
      Width           =   4775
   End
   Begin VB.Label Speed 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0E7E3&
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   3600
      Width           =   4775
   End
   Begin VB.Label UnName 
      BackStyle       =   0  'Transparent
      Caption         =   "UnName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Line lnview 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   3600
      X2              =   5640
      Y1              =   1605
      Y2              =   3720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Featured Unit :-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Mission ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "FrmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idx As Integer

Private Sub Form_Load()
Set_featured
End Sub

Private Sub Form_Resize()
LoadBar.Left = 512
LoadBar.Width = Width - 1024
LoadBar.Top = Height - LoadBar.Height - 512
End Sub

Sub Set_featured()
Dim cnt As Integer
Dim idx As Integer
Dim img As String
Dim pth As String
Randomize

If Sgn(Rnd - Rnd) = 1 Or 0 Then
pth = App.path & "\Rules\tanks.ini"
Label8 = "Type : Tank"
cnt = Val(GetFromIni("Main", "Count", pth))
idx = Round(Rnd * (cnt - 1))
UnName = GetFromIni("Main", "t" & CStr((idx + 1)), pth)

Else
pth = App.path & "\Rules\Aircrafts.ini"
Label8 = "Type : Aircraft"
cnt = Val(GetFromIni("Main", "Count", pth))
idx = Round(Rnd * (cnt - 1))
UnName = GetFromIni("Main", "a" & CStr((idx + 1)), pth)
End If

Speed = "Speed : " & GetFromIni(UnName, "Speed", pth)
Label3 = "Weapon : " & GetFromIni(UnName, "Weapon", pth)
Label4 = "Power : " & GetFromIni(UnName, "Power", pth)
Label5 = "Tech Level : " & GetFromIni(UnName, "techlevel", pth)
Label6 = "Cost : " & GetFromIni(UnName, "cost", pth)
img = GetFromIni(UnName, "image", pth)
imgview.Tag = GetFromIni(UnName, "image", pth)
imgview.Picture = LoadPicture(App.path & "\images\" & img & "\" & img & "17 copy.gif")
imgview.Move (Screen.Width / 2) - (imgview.Width / 2), (Screen.Height / 2) - (imgview.Height / 2)
lnview.X2 = (Screen.Width / 2) - (imgview.Width / 2)
lnview.Y2 = (Screen.Height / 2) - (imgview.Height / 2)
idx = 17
tmrView = True
End Sub

Private Sub LoadBar_Change()
On Error Resume Next
If idx > 20 Then idx = 0
imgview.Picture = LoadPicture(App.path & "\images\" & imgview.Tag & "\" & imgview.Tag & CStr(idx) & " copy.gif")
DoEvents
idx = idx + 1
End Sub

Private Sub tmrload_Timer()
tmrload = False
Label7 = GetFromIni("Main", "Name", tmrload.Tag & "ini.ini")
Label7shad = Label7
Label7.Left = (Screen.Width / 2) - (Label7.Width / 2)
Label7shad.Top = Label7.Top + 20
Label7shad.Left = Label7.Left + 20
Label7.Visible = True
Label7shad.Visible = True
frmmain.LoadMap tmrload.Tag, LoadBar
Unload Me
End Sub

