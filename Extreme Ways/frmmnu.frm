VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmmnu 
   BackColor       =   &H00200405&
   BorderStyle     =   0  'None
   Caption         =   "RoR-II"
   ClientHeight    =   10890
   ClientLeft      =   4680
   ClientTop       =   1725
   ClientWidth     =   10020
   Icon            =   "frmmnu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10890
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrCr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   1920
   End
   Begin RichTextLib.RichTextBox RMsc 
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmmnu.frx":27A2
   End
   Begin VB.PictureBox Load 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   460
      Left            =   6000
      ScaleHeight     =   435
      ScaleWidth      =   1845
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   1870
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading ...."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox PicSet 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   2520
      ScaleHeight     =   4905
      ScaleWidth      =   3465
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CheckBox Check4 
         BackColor       =   &H0059341C&
         Caption         =   "Enable Ground"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   2775
      End
      Begin VB.HScrollBar SR 
         CausesValidation=   0   'False
         Height          =   135
         LargeChange     =   5
         Left            =   360
         Max             =   50
         TabIndex        =   15
         Top             =   3960
         Width           =   2775
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0059341C&
         Caption         =   "Enable Minimap"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0059341C&
         Caption         =   "Enable Start Reveal FX"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0059341C&
         Caption         =   "Enable Music Tracks"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Game Speed  [Requires Resources]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   3720
         Width           =   2775
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   8160
      ScaleHeight     =   10665
      ScaleWidth      =   1830
      TabIndex        =   0
      Top             =   0
      Width           =   1860
      Begin RoR.button button3 
         Height          =   450
         Left            =   0
         TabIndex        =   4
         Top             =   9000
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button2 
         Height          =   450
         Left            =   0
         TabIndex        =   3
         Top             =   8520
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button1 
         Height          =   450
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   1
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   2
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   3
         Left            =   0
         TabIndex        =   7
         Top             =   1560
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   4
         Left            =   0
         TabIndex        =   8
         Top             =   2040
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button btnLvL 
         Height          =   450
         Index           =   5
         Left            =   0
         TabIndex        =   9
         Top             =   2520
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button4 
         Height          =   450
         Left            =   0
         TabIndex        =   10
         Top             =   8040
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin RoR.button button5 
         Height          =   450
         Left            =   0
         TabIndex        =   20
         Top             =   7560
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   794
      End
      Begin VB.Image imgSide 
         Height          =   135
         Left            =   0
         Top             =   4560
         Width           =   1695
      End
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   1920
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   3387
      _Version        =   393217
      BackColor       =   16644599
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmmnu.frx":2824
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00D48D54&
      Height          =   270
      Left            =   2055
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1080
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Line lnhoz 
      BorderColor     =   &H00F2E8E1&
      X1              =   -100
      X2              =   -85
      Y1              =   -100
      Y2              =   -85
   End
   Begin VB.Line lnvrt 
      BorderColor     =   &H00F2E2D9&
      X1              =   -100
      X2              =   -85
      Y1              =   -100
      Y2              =   -85
   End
   Begin VB.Image Image1 
      Height          =   4815
      Left            =   2040
      Top             =   3480
      Width           =   6015
   End
End
Attribute VB_Name = "frmmnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnLvL_click(Index As Integer, str As String)
If str <> "Not Available" Then
Load.Visible = True
GoOn
FrmLoad.Visible = True
FrmLoad.ZOrder 1
MakeTransparent FrmLoad.hwnd, 0
DoEvents
FrmLoad.tmrload = False
Dim i As Integer
For i = 0 To 255 Step 30
MakeTransparent FrmLoad.hwnd, i
FrmLoad.ZOrder 0
Next
MakeOpaque FrmLoad.hwnd
FrmLoad.tmrload = True
FrmLoad.tmrload.Tag = App.path & "\maps\Camp" & CStr(Index) & "\"
DoEvents
Unload Me
Else
rtb.Text = "Sir, You will have to complete the previous missions to proceed"
rtb.Visible = True
End If
End Sub

Private Sub btnLvL_Move(Index As Integer, str As String)
If str <> "Not Available" Then
rtb.Visible = True
rtb.Height = 1920
rtb.LoadFile App.path & "\Maps\Camp" & CStr(Index) & "\Description.txt"
DoEvents
Else
rtb.Visible = False
End If
Credit_Stop
Normalize
PicSet.Visible = False
SetXY Picture1.Left - 256, btnLvL(Index).Top + (btnLvL(Index).Height / 2)
End Sub

Private Sub button1_click(str As String)
Load.Visible = True
GoOn
FrmLoad.Visible = True
FrmLoad.ZOrder 1
MakeTransparent FrmLoad.hwnd, 0
DoEvents
FrmLoad.tmrload = False
Dim i As Integer
For i = 0 To 255 Step 30
MakeTransparent FrmLoad.hwnd, i
FrmLoad.ZOrder 0
Next
MakeOpaque FrmLoad.hwnd
FrmLoad.tmrload = True
FrmLoad.tmrload.Tag = App.path & "\maps\Camp" & GetFromIni("Main", "Progress", App.path & "\set.cfg") & "\"
DoEvents
Unload Me
End Sub

Private Sub button1_Move(str As String)
rtb.Visible = True
rtb.LoadFile App.path & "\Maps\Camp" & GetFromIni("Main", "Progress", App.path & "\set.cfg") & "\Description.txt"
DoEvents
SetXY Picture1.Left - 256, (button1.Height / 2)
PicSet.Visible = False
Normalize
Credit_Stop
End Sub

Private Sub button2_Move(str As String)
rtb.Visible = False
PicSet.Visible = False
Normalize
Credit_Start
SetXY Picture1.Left - 256, button2.Top + (button2.Height / 2)
End Sub

Private Sub button3_click(str As String)
End

End Sub

Private Sub button3_Move(str As String)
rtb.Visible = True
PicSet.Visible = False
rtb.Text = "Do you really want to leave the Battle Arena" & vbCrLf & " ?  &  ! "
SetXY Picture1.Left - 256, button3.Top + (button3.Height / 2)
Normalize
Credit_Stop
End Sub

Private Sub eves_MouseEnter(ctlEntered As Control)
If LCase(ctlEntered.Name) = "imgside" Then
LoadCursor "Pointer", ctlEntered.hwnd
End If
End Sub

Private Sub button4_Move(str As String)
rtb.Visible = True
Normalize
rtb.Text = "Set Settings according to your computer's capability" & vbCrLf & "More Visual FX will slow down the game."
PicSet.Visible = True
SetXY Picture1.Left - 256, button4.Top + (button4.Height / 2)
Credit_Stop
End Sub


Private Sub button5_Move(str As String)
On Error GoTo tnt
rtb.BackColor = vbWhite
rtb.Visible = True
PicSet.Visible = False
rtb.Height = Screen.Height / 1.4
rtb.LoadFile App.path & "\Manual.rtf"
SetXY Picture1.Left - 256, button5.Top + (button5.Height / 2)
Credit_Stop
Exit Sub
tnt:
MsgBox Err.Description
End Sub

Sub Normalize()
rtb.Height = 1960
rtb.Font.Name = "Arial"
rtb.Font.Bold = False
rtb.Font.Size = 11
rtb.Font.Italic = False
rtb.BackColor = &HFDF9F7
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteIni "Main", "Music", Check1.Value, App.path & "\set.cfg"
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteIni "Main", "Reveal", Check2.Value, App.path & "\set.cfg"
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteIni "Main", "Minimap", Check3.Value, App.path & "\set.cfg"
End Sub

Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
WriteIni "Main", "Ground", Check4.Value, App.path & "\set.cfg"
End Sub

Private Sub Form_Load()
Dim X As Integer
CheckState
Me.WindowState = 2
Set Image1.Picture = LoadPicture(App.path & "\images\buildbar\back.jpg")
Set PicSet.Picture = LoadPicture(App.path & "\Images\BuildBar\Menu.gif")
Set Load.Picture = LoadPicture(App.path & "\images\buildbar\button.gif")
button1.caption "Resume Campaign"
button1.image App.path & "\images\buildbar\button.gif"
button1.image_on App.path & "\images\buildbar\button0.gif"
button2.caption "Credits"
button2.image App.path & "\images\buildbar\button.gif"
button2.image_on App.path & "\images\buildbar\button0.gif"
button3.caption "Exit Game"
button3.image App.path & "\images\buildbar\button.gif"
button3.image_on App.path & "\images\buildbar\button0.gif"
button4.caption "Settings"
button4.image App.path & "\images\buildbar\button.gif"
button4.image_on App.path & "\images\buildbar\button0.gif"
button5.caption "Breifing"
button5.image App.path & "\images\buildbar\button.gif"
button5.image_on App.path & "\images\buildbar\button0.gif"
FormOnTop Me
RMsc.LoadFile App.path & "\Maps\Resource\Credits.txt"
Label3.caption = RMsc.Text
ReSet
imgSide.Left = 0
imgSide.Top = 0
Set imgSide.Picture = LoadPicture(App.path & "\images\Buildbar\Side.gif")
imgSide.Stretch = True
imgSide.Height = Screen.Height
LoadCursor "Select", hwnd
ComeOn
LoadCursor "Pointer", Picture1.hwnd
Normalize
Getini
End Sub

Sub CheckState()
Dim fso As New FileSystemObject
Dim pth As String
pth = App.path & "\" & App.EXEName & ".exe"
If fso.FileExists(pth) = False Then
MsgBox "  : You are Running RoR-II in VB-IDE Mode" & vbCrLf & "  : You Must Run it Compiled", vbCritical, "Error Code : 0x000021"
End
End If
Set fso = Nothing
End Sub

Sub ReSet()
For X = 1 To 5
If X <= Val(GetFromIni("Main", "Progress", App.path & "\set.cfg")) Then
btnLvL(X).caption GetFromIni("Main", "Name", App.path & "\maps\camp" & X & "\ini.ini")
Else
btnLvL(X).caption "Not Available"
End If
btnLvL(X).image App.path & "\images\buildbar\button.gif"
btnLvL(X).image_on App.path & "\images\buildbar\button0.gif"
Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetXY X, Y
End Sub

Private Sub Form_Resize()
Picture1.Left = Me.Width - Picture1.Width
Picture1.Height = Me.Height
Load.Left = (Me.Width / 2) - (Load.Width / 2) - Picture1.Width
Load.Top = (Height / 2) - (Load.Height / 2)
Image1.Left = (Me.Width / 2) - (Image1.Width / 2) - Picture1.Width
Image1.Top = (Height / 2) - (Image1.Height / 2)
Label3.Left = (Me.Width / 2) - (Label3.Width / 2) - Picture1.Width
Label4.Left = Label3.Left + 30
PicSet.Left = (Me.Width / 2) - (PicSet.Width / 2) - Picture1.Width
PicSet.Top = (Height / 2) - (PicSet.Height / 2)
button3.Top = Height - button3.Height
button4.Top = Height - (button4.Height * 3)
button5.Top = Height - (button4.Height * 4)
button2.Top = button3.Top - 480
End Sub

Sub SetNext()
Show
ComeOn
ReSet
End Sub

Sub SetLoose()
Show
Unload frmmain
button1.caption "Retry Mission"
rtb.Visible = True
rtb.Text = "You Lost" & vbCrLf & "Retry Again and Test your Skills"
End Sub

Sub ComeOn()
Dim k As Long
Dim i As Integer
For k = 1860 To 0 Step -3
button1.Left = k
button2.Left = k
button3.Left = k
button4.Left = k
button5.Left = k
For i = 1 To 5
btnLvL(i).Left = k
Next
Next
End Sub

Sub GoOn()
Dim k As Long
Dim i As Integer
k = 2
Y:
button1.Left = k
button2.Left = k
button3.Left = k
button4.Left = k
button5.Left = k
For i = 1 To 5
btnLvL(i).Left = k
Next
If k > 1860 Then Exit Sub
k = k * 1.3
GoTo Y
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetXY X + Image1.Left, Y + Image1.Top
End Sub

Sub SetXY(ByVal cX As Integer, ByVal cY As Integer)
lnvrt.X1 = cX
lnvrt.X2 = cX
lnvrt.Y1 = 0
lnvrt.Y2 = Height

lnhoz.Y1 = cY
lnhoz.Y2 = cY
lnhoz.X1 = 0
lnhoz.X2 = Width
End Sub

Private Sub PicSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetXY PicSet.Left, PicSet.Top
End Sub

Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetXY rtb.Width, rtb.Height
End Sub

Sub Getini()
Check1.Value = Val(GetFromIni("Main", "Music", App.path & "\set.cfg"))
Check2.Value = Val(GetFromIni("Main", "Reveal", App.path & "\set.cfg"))
Check3.Value = Val(GetFromIni("Main", "Minimap", App.path & "\set.cfg"))
Check4.Value = Val(GetFromIni("Main", "Ground", App.path & "\set.cfg"))
SR.Value = Val(GetFromIni("Main", "Speed", App.path & "\set.cfg"))
End Sub

Private Sub SR_Change()
WriteIni "Main", "Speed", SR.Value, App.path & "\set.cfg"
End Sub

Sub Credit_Start()
Label3.Top = Screen.Height
Label3.Visible = True
Label4.Visible = True
Label4.ZOrder 0
Label3.ZOrder 0
TmrCr = True
End Sub
Private Sub TmrCr_Timer()
If Label3.Top + Label3.Height < 0 Then
Label3.Top = Screen.Height
Else
Label3.Top = Label3.Top - 15
End If
Label4.Top = Label3.Top + 30
End Sub

Sub Credit_Stop()
TmrCr = False
Label3.Visible = False
Label4.Visible = False
End Sub
