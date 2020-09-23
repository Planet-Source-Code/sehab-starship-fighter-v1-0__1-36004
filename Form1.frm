VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5775
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   8970
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Paused 
      Caption         =   $"Form1.frx":0000
      Height          =   4095
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   3120
      Top             =   720
   End
   Begin VB.TextBox Spys 
      Height          =   975
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   48
      Text            =   "Form1.frx":03C4
      Top             =   2880
      Width           =   735
   End
   Begin VB.Timer Timer9 
      Interval        =   100
      Left            =   480
      Top             =   720
   End
   Begin VB.Timer Timer8 
      Interval        =   30000
      Left            =   3720
      Top             =   720
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Special"
      Height          =   255
      Left            =   1800
      TabIndex        =   46
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   2160
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar MyLife 
      Height          =   135
      Left            =   240
      TabIndex        =   43
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Max             =   200
      Scrolling       =   1
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7800
      Top             =   2760
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6480
      Top             =   3600
   End
   Begin VB.Timer Timer4 
      Interval        =   2000
      Left            =   1800
      Top             =   2400
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Fire"
      Height          =   255
      Left            =   4800
      TabIndex        =   34
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command18 
      Caption         =   "S-Bomb"
      Height          =   255
      Left            =   4800
      TabIndex        =   33
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Sorpedo"
      Height          =   255
      Left            =   4800
      TabIndex        =   32
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Ship--Seeker"
      Height          =   255
      Left            =   4800
      TabIndex        =   31
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Run-Away"
      Height          =   255
      Left            =   4800
      TabIndex        =   29
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2520
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   2160
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   2040
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00000000&
      Caption         =   "Make Agreement"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00000000&
      Caption         =   "Send Spy To"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Train Space Spy"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00000000&
      Caption         =   "Train Space Man"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00000000&
      Caption         =   "Train 2 StarShip Pilots"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00000000&
      Caption         =   "Train Merchant"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00000000&
      Caption         =   "Train Star Fighter Pilot"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "Form1.frx":03DC
      Left            =   3000
      List            =   "Form1.frx":03F2
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Text            =   "Planets"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00000000&
      Caption         =   "Create A Star Fighter"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Create A StarShip"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Create A Missle"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      Caption         =   "Sell To"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00000000&
      Caption         =   "Buy From"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "Trade With"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Attack Them"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar You 
      Height          =   135
      Left            =   4800
      TabIndex        =   30
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Max             =   200
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar Enemy 
      Height          =   135
      Left            =   7200
      TabIndex        =   35
      Top             =   4080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Max             =   200
      Scrolling       =   1
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enemy's Life:  "
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7200
      TabIndex        =   47
      Top             =   3840
      Width           =   1020
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Minimize"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   45
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Life Line: "
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   240
      TabIndex        =   44
      Top             =   5040
      Width           =   690
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "You Have Attacked:  "
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   42
      Top             =   1440
      Width           =   4095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "You:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   41
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Enemy:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7200
      TabIndex        =   40
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Label4"
      Height          =   255
      Left            =   6480
      TabIndex        =   39
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "Label4"
      Height          =   255
      Left            =   6480
      TabIndex        =   38
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "Label4"
      Height          =   255
      Left            =   6480
      TabIndex        =   37
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Label4"
      Height          =   255
      Left            =   6480
      TabIndex        =   36
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   17
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   18
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   19
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   20
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   21
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   22
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   23
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   24
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   25
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   26
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   27
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3000
      TabIndex        =   28
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   1440
      Picture         =   "Form1.frx":0423
      Top             =   360
      Width           =   1650
   End
   Begin VB.Menu mnumissle 
      Caption         =   "Missle"
      Begin VB.Menu mnusbomb 
         Caption         =   "S-Bomb -- 10,000 OreS"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnusorpedo 
         Caption         =   "Sorpedo -- 8,000 OreS"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnushipseeker 
         Caption         =   "Ship-Seeker -- 12,000 OreS"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnucreateammo 
         Caption         =   "Ammo -- 2,000 OreS"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnubuy 
      Caption         =   "Buy"
      Begin VB.Menu mnusb 
         Caption         =   "S-Bomb(1) -- 10,000"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuso 
         Caption         =   "Sorpedo(1) -- 12,000"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuss 
         Caption         =   "Ship-Seeker(1) -- 15,000"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuammo 
         Caption         =   "Ammo(50) -- 5,000"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnudecline 
         Caption         =   "Decline"
      End
   End
   Begin VB.Menu mnuhelp2 
      Caption         =   "Help"
      Begin VB.Menu mnuhelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'all of the dim's
Dim OreS As Long
Dim Spy As Integer
Dim Merchant As Integer
Dim StarShip As Integer
Dim SPilot As Integer
Dim FPilot As Integer
Dim StarFighter As Integer
Dim SpaceMan As Integer
Dim SBomb As Integer
Dim Sorpedo As Integer
Dim SS As Integer
Dim Ammo As Integer
Dim Life As Integer
Dim ELife As Integer
Dim MLife As Integer
Dim VLife As Integer
Dim SLife As Integer
Dim QLife As Integer
Dim PLife As Integer
Dim Less As Integer
Dim SpyFound As Integer
Dim Answer As VbMsgBoxResult
Dim Attacker As Integer
Dim Attack As String
Dim Pause As Boolean
Dim T1 As Boolean
Dim T2 As Boolean
Dim T3 As Boolean
Dim T4 As Boolean
Dim T5 As Boolean
Dim T6 As Boolean
Dim T7 As Boolean
Dim T8 As Boolean
Dim T9 As Boolean
Dim T10 As Boolean
'All The Dim's


Private Sub Command1_Click()
Me.Left = 200
If Text1.Text = "Planets" Then
Else
Timer6.Enabled = True
Me.Width = 9060
Timer4.Enabled = False
Timer3.Enabled = True
Label20.Caption = "You Have Attacked:  " & Text1.Text
If Text1.Text = "Earth" Then
Enemy.Value = ELife
Timer5.Enabled = True
End If
If Text1.Text = "Mars" Then
Enemy.Value = MLife
Timer5.Enabled = True
End If
If Text1.Text = "Venus" Then
Enemy.Value = VLife
Timer5.Enabled = True
End If
If Text1.Text = "QuadTri" Then
Enemy.Value = QLife
Timer5.Enabled = True
End If
If Text1.Text = "PeeQua" Then
Enemy.Value = PLife
Timer5.Enabled = True
End If
If Text1.Text = "Saturn" Then
Enemy.Value = SLife
Timer5.Enabled = True
End If
End If
End Sub

Private Sub Command10_Click()
OreS = OreS - 40
SPilot = SPilot + 2
End Sub

Private Sub Command11_Click()
OreS = OreS - 10
SpaceMan = SpaceMan + 1
End Sub

Private Sub Command12_Click()
OreS = OreS - 50
Spy = Spy + 1
End Sub

Private Sub Command13_Click()
If Text1.Text = "Planets" Then
Else
If Spy > 0 Then
Timer10.Enabled = True
Spys.Text = "Our Spys' Are At " & Text1.Text
Spy = 0
End If
End If
End Sub

Private Sub Command14_Click()
MsgBox "Wait for the next version for this to work."
End Sub

Private Sub Command20_Click()
Answer = MsgBox("Are You Sure -- 50,000 OreS for 50+ LifeLine", vbYesNo)
If Answer = vbYes Then
Life = Life + 50
OreS = OreS - 50000
Else
End If
End Sub

Private Sub Command3_Click()
PopupMenu mnubuy
End Sub

Private Sub Command4_Click()
On Error GoTo Error
OreS = OreS + 10000
Ammo = Ammo - 100
Error:
End Sub

Private Sub Command5_Click()
PopupMenu mnumissle
End Sub

Private Sub Command6_Click()
OreS = OreS - 100100
StarShip = StarShip + 1
End Sub

Private Sub Command7_Click()
StarFighter = StarFighter + 1
OreS = OreS - 100
End Sub

Private Sub Command8_Click()
OreS = OreS - 50
FPilot = FPilot + 1
End Sub

Private Sub Command9_Click()
OreS = OreS - 20
Merchant = Merchant + 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then Form2.Show
End Sub

Private Sub Form_Load()
OreS = 100000
Spy = 1
Merchant = 3
StarShip = 1
SPilot = 2
FPilot = 1
StarFighter = 3
SpaceMan = 100
SBomb = 0
Sorpedo = 0
SS = 0
Ammo = 10000
Life = 100
Me.Width = 4740
You.Value = Life
ELife = 100
PLife = 100
QLife = 100
SLife = 100
VLife = 100
MLife = 100
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Me.Move X, Y
End If
End Sub

Private Sub Image1_Click()
If Pause = False Then Pause = True Else Pause = False
If Pause = False Then
Paused.Visible = False
GetAllSettings App.Path, "Timers"
If T1 = True Then Timer1.Enabled = True Else Timer1.Enabled = False
If T2 = True Then Timer2.Enabled = True Else Timer2.Enabled = False
If T3 = True Then Timer3.Enabled = True Else Timer3.Enabled = False
If T4 = True Then Timer4.Enabled = True Else Timer4.Enabled = False
If T5 = True Then Timer5.Enabled = True Else Timer5.Enabled = False
If T6 = True Then Timer6.Enabled = True Else Timer6.Enabled = False
If T7 = True Then Timer7.Enabled = True Else Timer7.Enabled = False
If T8 = True Then Timer8.Enabled = True Else Timer8.Enabled = False
If T9 = True Then Timer9.Enabled = True Else Timer9.Enabled = False
If T10 = True Then Timer10.Enabled = True Else Timer10.Enabled = False
End If
If Pause = True Then
Paused.Visible = True
If Timer1.Enabled = True Then T1 = True Else T1 = False
If Timer2.Enabled = True Then T2 = True Else T2 = False
If Timer3.Enabled = True Then T3 = True Else T3 = False
If Timer4.Enabled = True Then T4 = True Else T4 = False
If Timer5.Enabled = True Then T5 = True Else T5 = False
If Timer6.Enabled = True Then T6 = True Else T6 = False
If Timer7.Enabled = True Then T7 = True Else T7 = False
If Timer8.Enabled = True Then T8 = True Else T8 = False
If Timer9.Enabled = True Then T9 = True Else T9 = False
If Timer10.Enabled = True Then T10 = True Else T10 = False
SaveSetting App.Path, "Timers", 1, T1
SaveSetting App.Path, "Timers", 2, T2
SaveSetting App.Path, "Timers", 3, T3
SaveSetting App.Path, "Timers", 4, T4
SaveSetting App.Path, "Timers", 5, T5
SaveSetting App.Path, "Timers", 6, T6
SaveSetting App.Path, "Timers", 7, T7
SaveSetting App.Path, "Timers", 8, T8
SaveSetting App.Path, "Timers", 9, T9
SaveSetting App.Path, "Timers", 10, T10
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Timer9.Enabled = False
Timer10.Enabled = False
End If

End Sub

Private Sub Image1_DblClick()
Form2.Show
End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub Label22_Click()
Me.WindowState = 1
End Sub

Private Sub List1_Click()
Text1.Text = List1
End Sub

Private Sub mnuammo_Click()
OreS = OreS - 5000
Ammo = Ammo + 50
End Sub

Private Sub mnucreateammo_Click()
On Error GoTo Error
Ammo = Ammo + 100
OreS = OreS - 2000
Error:
End Sub

Private Sub mnuhelp_Click()
Form2.Show
End Sub

Private Sub mnusb_Click()
OreS = OreS - 10000
SBomb = SBomb + 1
End Sub

Private Sub mnusbomb_Click()
OreS = OreS - 10000
SBomb = SBomb + 1
End Sub

Private Sub mnushipseeker_Click()
OreS = OreS - 12000
SS = SS + 1
End Sub

Private Sub mnuso_Click()
OreS = OreS - 12000
Sorpedo = Sorpedo + 1
End Sub

Private Sub mnusorpedo_Click()
OreS = OreS - 8000
Sorpedo = Sorpedo + 1
End Sub

Private Sub mnuss_Click()
OreS = OreS - 15000
SS = SS + 1
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "OreS = " & OreS
Label3.Caption = "StarShip = " & StarShip
Label4.Caption = "StarFighter = " & StarFighter
Label5.Caption = "S-Bomb = " & SBomb
Label6.Caption = "Sorpedo = " & Sorpedo
Label7.Caption = "Ship-Seeker = " & SS
Label8.Caption = "Ammo = " & Ammo
Label9.Caption = "Merchant = " & Merchant
Label10.Caption = "SpaceMen = " & SpaceMan
Label11.Caption = "SpaceSpy = " & Spy
Label12.Caption = "SPilot = " & SPilot
Label13.Caption = "FPilot = " & FPilot
End Sub

Private Sub Timer10_Timer()
SpyFound = Int(Rnd * 10)
If SpyFound = 1 Then
MsgBox Spys.Text & " and have found out that they have 3 SBombs."
End If
If SpyFound = 1 Then
MsgBox Spys.Text & " and have found out that they have 3 SBombs."
End If
If SpyFound = 2 Then
MsgBox Spys.Text & " and have found out that they have 3 SBombs."
End If
If SpyFound = 3 Then
MsgBox Spys.Text & " and have found out that they have 3 SBombs."
End If
If SpyFound = 4 Then
MsgBox Spys.Text & " and have found out that they have 500000 of Ammo."
End If
If SpyFound = 5 Then
MsgBox Spys.Text & " and have found out that they have 2 Sorpedos."
End If
If SpyFound = 6 Then
MsgBox Spys.Text & " and have found out that they have 3 Ship-Seekers."
End If
If SpyFound = 7 Then
MsgBox Spys.Text & " and have found out that they have 5 SBombs."
End If
If SpyFound = 8 Then
MsgBox Spys.Text & " have found out that they have 100 of Ammo."
End If
If SpyFound = 9 Then
MsgBox Spys.Text & " have found out that they have 6 Ship-Seekers."
End If
If SpyFound = 10 Then
MsgBox Spys.Text & " have found out that they have 6 Sorpedos."
End If
End Sub

Private Sub Timer2_Timer()
Label21.Caption = "Life Line: " & MyLife.Value & " : Max = 200"
MyLife.Value = You.Value
End Sub

Private Sub Command19_Click()
On Error GoTo Win
Enemy.Value = Enemy.Value - 1
Ammo = Ammo - 2
Win:
End Sub

Private Sub Command18_Click()
On Error GoTo Win
Enemy.Value = Enemy.Value - 30
You.Value = You.Value - 10
SBomb = SBomb - 1
Win:
End Sub

Private Sub Command17_Click()
On Error GoTo Win
Enemy.Value = Enemy.Value - 15
Sorpedo = Sorpedo - 1
Win:
End Sub

Private Sub Command16_Click()
On Error GoTo Win
Enemy.Value = Enemy.Value - 25
SS = SS - 1
Win:
End Sub

Private Sub Command15_Click()
Life = You.Value
Me.Width = 4740
Timer6.Enabled = False
Timer5.Enabled = False
Timer4.Enabled = True
End Sub

Private Sub Timer3_Timer()
Label17.Caption = Ammo
Label16.Caption = SBomb
Label15.Caption = Sorpedo
Label14.Caption = SS
If Ammo <= 0 Then
Command19.Enabled = False
Else
Command19.Enabled = True
End If
If SBomb <= 0 Then
Command18.Enabled = False
Else
Command18.Enabled = True
End If
If Sorpedo <= 0 Then
Command17.Enabled = False
Else
Command17.Enabled = True
End If
If SS <= 0 Then
Command16.Enabled = False
Else: Command16.Enabled = True
If You.Value <= 0 Then
MsgBox "You Have Lost To:  " & Text1.Text
End
End If
End If
End Sub


Private Sub Timer4_Timer()
If Life >= 200 Then
Life = 200
Else
Life = Life + 1
End If
You.Value = Life
If SLife = 0 Then
Else
SLife = SLife + 2
End If
If MLife = 0 Then
Else
MLife = MLife + 2
End If
If ELife = 0 Then
Else
ELife = ELife + 2
End If
If VLife = 0 Then
Else
VLife = VLife + 2
End If
If QLife = 0 Then
Else
QLife = QLife + 2
End If
If PLife = 0 Then
Else
PLife = PLife + 2
End If
If Life >= 199 Then Life = 200
If SLife >= 199 Then SLife = 200
If MLife >= 199 Then MLife = 200
If ELife >= 199 Then ELife = 200
If VLife >= 199 Then VLife = 200
If QLife >= 199 Then QLife = 200
If PLife >= 199 Then PLife = 200
If Life <= 0 Then
MsgBox "You Have Lost"
End
End If
End Sub

Private Sub Timer5_Timer()
On Error GoTo Error:
You.Value = You.Value - Int(Rnd * 12)
If You.Value <= 1 Then
MsgBox "You Have Lost To " & Text1.Text
StarShip = StarShip - 1
Life = 100
Timer4.Enabled = True
Timer5.Enabled = False
Timer6.Enabled = False
End If
If Enemy.Value <= 1 Then
MsgBox Text1.Text & " has lost."
SpaceMan = SpaceMan - 5
StarFighter = StarFighter - 2
FPilot = FPilot - 2
Merchant = Merchant - 1
Timer4.Enabled = True
Timer5.Enabled = False
Timer6.Enabled = False
Me.Width = 4740
Life = You.Value
End If
Less = You.Value - Enemy.Value
If Enemy.Value < You.Value And Less = 20 Then
MsgBox Text1.Text & " Has Gave Up!!"
Me.Width = 4740
Timer4.Enabled = True
Timer5.Enabled = False
Timer6.Enabled = False
Life = You.Value
End If
Label23.Caption = "Enemy's Life: " & Enemy.Value
Error:

End Sub

Private Sub Timer6_Timer()
If Text1.Text = "Earth" Then
ELife = Enemy.Value
End If
If Text1.Text = "Mars" Then
MLife = Enemy.Value
End If
If Text1.Text = "Venus" Then
VLife = Enemy.Value
End If
If Text1.Text = "Saturn" Then
SLife = Enemy.Value
End If
If Text1.Text = "QuadTri" Then
QLife = Enemy.Value
End If
If Text1.Text = "PeeQua" Then
PLife = Enemy.Value
End If
End Sub

Private Sub Timer7_Timer()
If MLife = 0 And ELife = 0 And PLife = 0 And QLife = 0 And VLife = 0 And SLife = 0 Then
MsgBox "You Have Beat The Game."
End
End If
End Sub

Private Sub Timer8_Timer()
Attacker = Int(Rnd * 8)
If Attacker = 0 Then
Attack = "Earth"
Enemy.Value = ELife
End If
If Attacker = 1 Then
Attack = "Mars"
Enemy.Value = MLife
End If
If Attacker = 2 Then
Attack = "Saturn"
Enemy.Value = SLife
End If
If Attacker = 3 Then
Attack = "PeeQua"
Enemy.Value = PLife
End If
If Attacker = 4 Then
Attack = "Venus"
Enemy.Value = VLife
End If
If Attacker = 5 Then
Attack = "QuadTri"
Enemy.Value = QLife
End If
MsgBox "You Have Been Attacked By " & Attack
If Attack = "Earth" And Spys.Text = "Our Spys' Are At:  Earth" Then
MsgBox "We have found a spy!!!!!!!!!!!!!!"
End If
If Attack = "Mars" And Spys.Text = "Our Spys' Are At:  Mars" Then
MsgBox "We have found a spy!!!!!!!!!!!!!!"
End If
If Attack = "Venus" And Spys.Text = "Our Spys' Are At:  Venus" Then
MsgBox "We have found a spy!!!!!!!!!!!!!!"
End If
If Attack = "Saturn" And Spys.Text = "Our Spys' Are At:  Saturn" Then
MsgBox "We have found a spy!!!!!!!!!!!!!!"
End If
If Attack = "QuadTri" And Spys.Text = "Our Spys' Are At:  QuadTri" Then
MsgBox "We have found a spy!!!!!!!!!!!!!!"
End If
If Attack = "PeeQua" And Spys.Text = "Our Spys' Are At:  PeeQua" Then
MsgBox "We have found a spy!!!!!!!!!!!!!!"
End If
Me.Width = 9060
Text1.Text = Attack
Me.Left = 200
Label20.Caption = "You have been attacked by " & Attack
Timer5.Enabled = True
Timer6.Enabled = True
Timer4.Enabled = False

Timer3.Enabled = True
End Sub

Private Sub Timer9_Timer()
If OreS <= -500000 Then
MsgBox "You have gone into too much debt.  !Sorry You Lose!"
End
End If
If SpaceMan <= -200 Then
MsgBox "You have lost too many people."
End
End If
If StarShip <= 0 Then
MsgBox "You have lost the last StarShip."
End
End If
If Ammo <= -32000 Then
MsgBox "You have sold too much ammo.  Sorry"
End
End If
End Sub
