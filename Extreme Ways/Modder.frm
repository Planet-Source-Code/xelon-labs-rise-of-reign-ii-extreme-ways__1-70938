VERSION 5.00
Begin VB.Form Modder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modder"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   1905
   ClientWidth     =   16965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   16965
   Begin VB.PictureBox Picture6 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   3915
      TabIndex        =   78
      Top             =   480
      Width           =   3975
      Begin VB.Label Label28 
         Caption         =   "There are a very few bugs , if found one report me"
         Height          =   255
         Left            =   0
         TabIndex        =   79
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   8055
      Index           =   3
      Left            =   11280
      ScaleHeight     =   7995
      ScaleWidth      =   5595
      TabIndex        =   56
      Top             =   480
      Width           =   5655
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Modder.frx":0000
         Left            =   2400
         List            =   "Modder.frx":000A
         TabIndex        =   74
         Text            =   "Water"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   2400
         TabIndex        =   72
         Text            =   "Power"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.PictureBox Picture4 
         Height          =   3495
         Left            =   0
         ScaleHeight     =   229
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   341
         TabIndex        =   64
         Top             =   4440
         Width           =   5175
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   3000
            Max             =   20
            TabIndex        =   75
            Top             =   3120
            Width           =   2055
         End
         Begin VB.PictureBox tnkimage 
            AutoRedraw      =   -1  'True
            Height          =   3135
            Left            =   0
            ScaleHeight     =   3075
            ScaleWidth      =   3195
            TabIndex        =   76
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.TextBox Text23 
         Height          =   285
         Left            =   2400
         TabIndex        =   63
         Text            =   "Tech level"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox Text22 
         Height          =   285
         Left            =   2400
         TabIndex        =   62
         Text            =   "Cost"
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   2400
         TabIndex        =   61
         Text            =   "Speed"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   2400
         TabIndex        =   60
         Text            =   "Image"
         Top             =   600
         Width           =   2895
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save this Item"
         Height          =   255
         Left            =   4320
         TabIndex        =   58
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   360
         TabIndex        =   57
         Text            =   "Weapon"
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label27 
         Caption         =   "Water"
         Height          =   255
         Left            =   2280
         TabIndex        =   73
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label26 
         Caption         =   "Power"
         Height          =   255
         Left            =   2280
         TabIndex        =   71
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Tech level"
         Height          =   255
         Left            =   2160
         TabIndex        =   69
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label24 
         Caption         =   "Cost"
         Height          =   255
         Left            =   2280
         TabIndex        =   68
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Speed"
         Height          =   255
         Left            =   2280
         TabIndex        =   67
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "Image"
         Height          =   255
         Left            =   2280
         TabIndex        =   66
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Weapon"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   3600
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   7095
      Index           =   1
      Left            =   11160
      ScaleHeight     =   7035
      ScaleWidth      =   5595
      TabIndex        =   38
      Top             =   480
      Width           =   5655
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   2400
         TabIndex        =   51
         Text            =   "Weapon"
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save this Item"
         Height          =   255
         Left            =   4320
         TabIndex        =   45
         Top             =   0
         Width           =   1215
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text17 
         Height          =   285
         Left            =   2400
         TabIndex        =   43
         Text            =   "Image"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   2400
         TabIndex        =   42
         Text            =   "Speed"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   2400
         TabIndex        =   41
         Text            =   "Cost"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   2400
         TabIndex        =   40
         Text            =   "Tech level"
         Top             =   2400
         Width           =   2895
      End
      Begin VB.PictureBox Picture5 
         Height          =   3495
         Left            =   120
         ScaleHeight     =   229
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   341
         TabIndex        =   39
         Top             =   3480
         Width           =   5175
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            Left            =   3000
            Max             =   20
            TabIndex        =   70
            Top             =   3120
            Width           =   2055
         End
         Begin VB.PictureBox airimage 
            AutoRedraw      =   -1  'True
            Height          =   3135
            Left            =   0
            ScaleHeight     =   3075
            ScaleWidth      =   3195
            TabIndex        =   77
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.Label Label14 
         Caption         =   "Weapon"
         Height          =   255
         Left            =   2280
         TabIndex        =   50
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Image"
         Height          =   255
         Left            =   2280
         TabIndex        =   49
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Speed"
         Height          =   255
         Left            =   2280
         TabIndex        =   48
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Cost"
         Height          =   255
         Left            =   2280
         TabIndex        =   47
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Tech level"
         Height          =   255
         Left            =   2280
         TabIndex        =   46
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   8895
      Index           =   2
      Left            =   5640
      ScaleHeight     =   8835
      ScaleWidth      =   5595
      TabIndex        =   19
      Top             =   480
      Width           =   5655
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   3360
         TabIndex        =   86
         Text            =   "20"
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text27 
         Height          =   285
         Left            =   3360
         TabIndex        =   84
         Text            =   "50"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         Height          =   285
         Left            =   3360
         TabIndex        =   55
         Text            =   "Techlevel"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   3360
         TabIndex        =   53
         Text            =   "Cost"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.PictureBox Picture3 
         Height          =   735
         Left            =   120
         ScaleHeight     =   675
         ScaleWidth      =   5115
         TabIndex        =   33
         Top             =   8040
         Width           =   5175
         Begin VB.TextBox Text26 
            Height          =   285
            Left            =   840
            TabIndex        =   81
            Text            =   "BaseX"
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Text25 
            Height          =   285
            Left            =   3120
            TabIndex        =   80
            Text            =   "BaseY"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   3120
            TabIndex        =   37
            Text            =   "DocY"
            Top             =   0
            Width           =   1695
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   840
            TabIndex        =   34
            Text            =   "DocX"
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label30 
            Caption         =   "BaseX"
            Height          =   255
            Left            =   240
            TabIndex        =   83
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "BaseY"
            Height          =   255
            Left            =   2520
            TabIndex        =   82
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "DocY"
            Height          =   255
            Left            =   2520
            TabIndex        =   36
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "DocX"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   4095
         Left            =   120
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   341
         TabIndex        =   32
         Top             =   3840
         Width           =   5175
         Begin VB.CommandButton Command2 
            BackColor       =   &H0000FF00&
            Height          =   195
            Index           =   2
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "OffSet"
            Top             =   1800
            Width           =   135
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   "DocSet"
            Top             =   1800
            Width           =   135
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H000000FF&
            Height          =   195
            Index           =   1
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   88
            ToolTipText     =   "OffSet"
            Top             =   1800
            Width           =   135
         End
         Begin Project1.aicAlphaImage pvwbldng 
            Height          =   2295
            Left            =   0
            Top             =   0
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   4048
            Image           =   "Modder.frx":0014
            Scaler          =   3
         End
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   3360
         TabIndex        =   26
         Text            =   "OffY"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   3360
         TabIndex        =   25
         Text            =   "OffX"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3360
         TabIndex        =   24
         Text            =   "image"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3360
         TabIndex        =   23
         Text            =   "power"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   2955
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Modder.frx":002C
         Left            =   3360
         List            =   "Modder.frx":0039
         TabIndex        =   21
         Text            =   "Combo2"
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save this Item"
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "Space"
         Height          =   255
         Left            =   2280
         TabIndex        =   87
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "Mass"
         Height          =   255
         Left            =   2280
         TabIndex        =   85
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "Techlevel"
         Height          =   255
         Left            =   2280
         TabIndex        =   54
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Cost"
         Height          =   255
         Left            =   2280
         TabIndex        =   52
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "OffY"
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "OffX"
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Image"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Power"
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Type"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Index           =   0
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   5595
      TabIndex        =   4
      Top             =   480
      Width           =   5655
      Begin VB.CommandButton Command5 
         Caption         =   "Save this Item"
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Modder.frx":0057
         Left            =   2400
         List            =   "Modder.frx":0061
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   600
         Width           =   2895
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Text            =   "damage"
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Text            =   "interval"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Text            =   "range"
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Text            =   "distance"
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Text            =   "speedstep"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Type"
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Damage"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Interval"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Range"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Distance"
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Speedstep"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buildings"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tanks"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aircrafts"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Weapons"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Modder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim chng As Boolean
Dim chngB As Boolean
Dim chngT As Boolean
Dim chngA As Boolean
Dim mX As Integer
Dim mY As Integer
Dim c32 As New c32bppDIB

Private Sub Combo1_Click()
On Error Resume Next
chng = True
If Combo1.Text = "Bomb" Then
Label2 = "Interval"
Text2 = ""
Label5.Visible = True
Text5.Visible = True
ElseIf Combo1.Text = "Laser" Then
Label2 = "Color"
Text2 = ""
Label5.Visible = False
Text5.Visible = False
End If
End Sub

Private Sub Combo2_Click()
On Error Resume Next
chngB = True
If UCase(Combo2) = "ACC" Or UCase(Combo2) = "WARFACTORY" Then
Picture3.Visible = True
Command2(1).Visible = True
Else
Picture3.Visible = False
Command2(1).Visible = False
End If
End Sub

Private Sub Combo3_Click()
On Error Resume Next
chngT = True
End Sub

Private Sub Command1_Click(Index As Integer)
On Error Resume Next
Dim x As Integer
For x = 0 To Picture1.UBound
Picture1(x).Visible = False
Next
Picture1(Index).Visible = True
Picture1(Index).Left = 0
Me.Height = Picture1(Index).Height + Picture1(Index).Top + 400
End Sub

Private Sub Command2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then
mX = x
mY = y
End If
End Sub

Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then
Command2(Index).Move Command2(Index).Left + (x - mX) / 15, Command2(Index).Top + (y - mY) / 15
 If Index = 0 Then
 Text9 = Command2(Index).Top - (pvwbldng.Height / 2) + Command2(Index).Height / 2
 Text8 = Command2(Index).Left - (pvwbldng.Width / 2) + Command2(Index).Width / 2
ElseIf Index = 1 Then
 Text11 = Command2(Index).Top - (pvwbldng.Height / 2) + Command2(Index).Height / 2
 Text10 = Command2(Index).Left - (pvwbldng.Width / 2) + Command2(Index).Width / 2
ElseIf Index = 2 Then
 Text25 = Command2(Index).Top - Command2(Index).Height / 2
 Text26 = Command2(Index).Left - Command2(Index).Width / 2
End If
chngB = True
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
chngA = False
Dim pth As String
pth = App.Path & "\rules\aircrafts.ini"
WriteIni List3, "Image", Text17, pth
WriteIni List3, "Speed", Text16, pth
WriteIni List3, "Cost", Text15, pth
WriteIni List3, "techlevel", Text14, pth
WriteIni List3, "weapon", Text12, pth
chng = False
End Sub

Private Sub Command4_Click()
On Error Resume Next
chngT = True
Dim pth As String
pth = App.Path & "\rules\aircrafts.ini"
WriteIni List4, "Image", Text20, pth
WriteIni List4, "Speed", Text21, pth
WriteIni List4, "Power", Text24, pth
WriteIni List4, "Cost", Text22, pth
WriteIni List4, "techlevel", Text23, pth
WriteIni List4, "weapon", Text19, pth
WriteIni List4, "water", Combo3, pth
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim pth As String
pth = App.Path & "\rules\weapons.ini"
WriteIni List1, "Type", Combo1, pth
WriteIni List1, "Damage", Text1, pth
WriteIni List1, "Interval", Text2, pth
WriteIni List1, "Range", Text3, pth
WriteIni List1, "Distance", Text4, pth
WriteIni List1, "SpeedStep", Text5, pth
chng = False
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim pth As String
pth = App.Path & "\rules\Buildings.ini"
WriteIni List2, "Type", Combo2, pth
WriteIni List2, "Power", Text6, pth
WriteIni List2, "Image", Text7, pth
WriteIni List2, "OffX", Text8, pth
WriteIni List2, "OffY", Text9, pth
WriteIni List2, "DocX", Text10, pth
WriteIni List2, "DocY", Text11, pth
WriteIni List2, "BaseX", CInt(26), pth
WriteIni List2, "BaseY", CInt(25), pth
WriteIni List2, "Space", Text28, pth
WriteIni List2, "Mass", Text27, pth

chngB = False
End Sub

Private Sub Form_Load()
On Error Resume Next
List1.Clear
List2.Clear
List3.Clear
Dim x As Integer
For x = 1 To Val(GetFromIni("Main", "count", App.Path & "\rules\weapons.ini"))
List1.AddItem GetFromIni("Main", "W" & CStr(x), App.Path & "\rules\weapons.ini")
Next
For x = 1 To Val(GetFromIni("Main", "count", App.Path & "\rules\buildings.ini"))
List2.AddItem GetFromIni("Main", "b" & CStr(x), App.Path & "\rules\buildings.ini")
Next
For x = 1 To Val(GetFromIni("Main", "count", App.Path & "\rules\aircrafts.ini"))
List3.AddItem GetFromIni("Main", "a" & CStr(x), App.Path & "\rules\aircrafts.ini")
Next
For x = 1 To Val(GetFromIni("Main", "count", App.Path & "\rules\tanks.ini"))
List4.AddItem GetFromIni("Main", "t" & CStr(x), App.Path & "\rules\tanks.ini")
Next
Height = 4950
Width = 5745
List1.ListIndex = 0
List2.ListIndex = 0
List3.ListIndex = 0
List4.ListIndex = 0
chng = False
Set c32 = New c32bppDIB
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set c32 = Nothing
End Sub

Private Sub HScroll1_Change()
'airimage.LoadImage_FromFile App.Path & "\images\" & Text17 & "\" & Text17 & Val((HScroll1.Value)) & " copy.gif"
End Sub

Private Sub HScroll1_Scroll()
On Error Resume Next
HScroll1_Change
End Sub

Private Sub HScroll2_Change()
On Error Resume Next
'tnkimage.LoadImage_FromFile App.Path & "\images\" & Text20 & "\" & Text20 & (HScroll2.Value - 1) & " copy.gif"
End Sub

Private Sub HScroll2_Scroll()
On Error Resume Next
HScroll2_Change
End Sub

Private Sub List1_Click()
On Error Resume Next
If chng = True Then
Dim res As VbMsgBoxResult
res = MsgBox("Do you want to save changes in this item", vbYesNoCancel, "Save ??")
If res = vbYes Then
Command5_Click
GoTo y
ElseIf res = vbNo Then
y:
Combo1 = GetFromIni(List1, "type", App.Path & "\rules\weapons.ini")
If UCase(Combo1) = "LASER" Then
Label2 = "Color"
Text2 = GetFromIni(List1, "color", App.Path & "\rules\weapons.ini")
Label5.Visible = False
Text5.Visible = False
Else
Label2 = "Interval"
Text2 = GetFromIni(List1, "interval", App.Path & "\rules\weapons.ini")
Label5.Visible = True
Text5.Visible = True
End If
Text1 = GetFromIni(List1, "damage", App.Path & "\rules\weapons.ini")
Text3 = GetFromIni(List1, "range", App.Path & "\rules\weapons.ini")
Text4 = GetFromIni(List1, "distance", App.Path & "\rules\weapons.ini")
Text5 = GetFromIni(List1, "speedstep", App.Path & "\rules\weapons.ini")
chng = False
End If
Else
GoTo y
End If
End Sub

Private Sub List2_Click()
On Error Resume Next
If chngB = True Then
Dim res As VbMsgBoxResult
res = MsgBox("Do you want to save changes in this item", vbYesNoCancel, "Save ??")
If res = vbYes Then
Command6_Click
GoTo y
ElseIf res = vbNo Then
y:
Text7 = GetFromIni(List2, "Image", App.Path & "\rules\buildings.ini")
pvwbldng.AutoSize = True
Combo2 = GetFromIni(List2, "type", App.Path & "\rules\buildings.ini")
Text6 = GetFromIni(List2, "power", App.Path & "\rules\buildings.ini")
Text8 = GetFromIni(List2, "OffX", App.Path & "\rules\buildings.ini")
Text9 = GetFromIni(List2, "offY", App.Path & "\rules\buildings.ini")
Text26 = GetFromIni(List2, "BaseX", App.Path & "\rules\buildings.ini")
Text25 = GetFromIni(List2, "BaseY", App.Path & "\rules\buildings.ini")
Text10 = GetFromIni(List2, "DocX", App.Path & "\rules\buildings.ini")
Text11 = GetFromIni(List2, "DocY", App.Path & "\rules\buildings.ini")
Text18 = GetFromIni(List2, "techlevel", App.Path & "\rules\buildings.ini")

pvwbldng.LoadImage_FromFile App.Path & "\images\Buildings\" & Text7 & ".png"


If Text18 <> "-1" Then
Text13 = GetFromIni(List2, "Cost", App.Path & "\rules\buildings.ini")
Else
Text13 = ""
End If
chngB = False
If UCase(Combo2) = "ACC" Then
Label12.Visible = True
Label13.Visible = True
Text10.Visible = True
Text11.Visible = True
Command2(1).Visible = True
Else
Label12.Visible = False
Label13.Visible = False
Text10.Visible = False
Text11.Visible = False
Command2(1).Visible = False
End If
chngB = False
End If
Else
GoTo y
End If
End Sub

Private Sub List3_Click()
On Error Resume Next
If chngA = True Then
Dim res As VbMsgBoxResult
res = MsgBox(" Do you want to change item without saving ??", vbYesNo, "Save ??")
If res = vbYes Then
y:
Text17 = GetFromIni(List3, "Image", App.Path & "\rules\aircrafts.ini")
'airimage.LoadImage_FromFile App.Path & "\images\" & Text17 & "\" & Text17.text & "0 copy.gif"
airimage.AutoSize = True
Text17 = GetFromIni(List3, "Image", App.Path & "\rules\aircrafts.ini")
Text16 = GetFromIni(List3, "Speed", App.Path & "\rules\aircrafts.ini")
Text14 = GetFromIni(List3, "techlevel", App.Path & "\rules\aircrafts.ini")
Text12 = GetFromIni(List3, "weapon", App.Path & "\rules\aircrafts.ini")

Set airimage.Picture = Nothing
c32.InitializeDIB airimage.Width, airimage.Height
c32.LoadPicture_File App.Path & "\images\" & Text17 & "\" & Text17 & "13 copy.gif"
c32.Render airimage.hDC
c32.DestroyDIB
airimage.Picture = airimage.Image

If Text14 <> "-1" Then
Text15 = GetFromIni(List3, "Cost", App.Path & "\rules\aircrafts.ini")
Else
Text15 = ""
End If
chngA = False
ElseIf res = vbNo Then
End If
Else
GoTo y
chngA = False
End If
End Sub

Private Sub List4_Click()
Dim res As VbMsgBoxResult
If chngT = True Then
res = MsgBox(" Do you want to change item without saving ??", vbYesNo, "Save ??")
If res = vbYes Then
y:
Text20 = GetFromIni(List4, "Image", App.Path & "\rules\tanks.ini")
tnkimage.AutoSize = True
'tnkimage.LoadImage_FromFile App.Path & "\images\" & Text20 & "\" & Text20.text & "0 copy.gif"
Text21 = GetFromIni(List4, "Speed", App.Path & "\rules\tanks.ini")
Text24 = GetFromIni(List4, "power", App.Path & "\rules\tanks.ini")
Combo3 = GetFromIni(List4, "water", App.Path & "\rules\tanks.ini")
Text23 = GetFromIni(List4, "techlevel", App.Path & "\rules\tanks.ini")
Text19 = GetFromIni(List4, "weapon", App.Path & "\rules\tanks.ini")

Set tnkimage.Picture = Nothing
c32.InitializeDIB tnkimage.Width, tnkimage.Height
c32.LoadPicture_File App.Path & "\images\" & Text20 & "\" & Text20 & "13 copy.gif"
c32.Render tnkimage.hDC
c32.DestroyDIB
tnkimage.Picture = tnkimage.Image

If Text23 <> "-1" Then
Text22 = GetFromIni(List4, "cost", App.Path & "\rules\tanks.ini")
Else
Text22 = ""
End If
chngT = False
ElseIf res = vbNo Then
End If
Else
GoTo y
chngT = False
End If
End Sub

Private Sub Text1_Change()
chng = True
End Sub

Private Sub Text10_Change()
On Error Resume Next
Command2(1).Left = Val(Text10) + (pvwbldng.Width / 2) - Command2(1).Width / 2
chngB = True
End Sub

Private Sub Text11_Change()
On Error Resume Next
Command2(1).Top = Val(Text11) + (pvwbldng.Height / 2) - Command2(1).Height / 2
chngB = True
End Sub

Private Sub Text26_Change()
On Error Resume Next
Command2(2).Left = Val(Text26) - Command2(2).Width / 2
chngB = True
End Sub

Private Sub Text25_Change()
On Error Resume Next
Command2(2).Top = Val(Text25) - Command2(2).Height / 2
chngB = True
End Sub

Private Sub Text12_Change()
chngA = True
End Sub
Private Sub Text14_Change()
chngA = True
End Sub
Private Sub Text15_Change()
chngA = True
End Sub
Private Sub Text16_Change()
chngA = True
End Sub
Private Sub Text17_Change()
chngA = True
End Sub
Private Sub Text19_Change()
chngT = True
End Sub
Private Sub Text2_Change()
chng = True
End Sub
Private Sub Text20_Change()
chngT = True
End Sub
Private Sub Text21_Change()
chngT = True
End Sub
Private Sub Text22_Change()
chngT = True
End Sub
Private Sub Text23_Change()
chngT = True
End Sub
Private Sub Text24_Change()
chngT = True
End Sub
Private Sub Text3_Change()
chng = True
End Sub
Private Sub Text4_Change()
chng = True
End Sub
Private Sub Text5_Change()
chng = True
End Sub
Private Sub Text6_Change()
chngB = True
End Sub
Private Sub Text7_Change()
chngB = True
End Sub
Private Sub Text8_Change()
On Error Resume Next
Command2(0).Left = Val(Text8) + (pvwbldng.Width / 2) - Command2(0).Width / 2
chng = True
End Sub
Private Sub Text9_Change()
On Error Resume Next
Command2(0).Top = Val(Text9) + (pvwbldng.Height / 2) - Command2(0).Height / 2
chng = True
End Sub
