VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "So You Need Help"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Help Me"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Stragety Help"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Winning"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Loosing"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   2895
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
List1.AddItem "How Do I Lose"
End Sub

Private Sub Command2_Click()
List1.AddItem "How Do I Win"
End Sub

Private Sub Command3_Click()
List1.AddItem "Some Stragety"
End Sub

Private Sub Command4_Click()
If Text1.Text = "How Do I Lose" Then
Label1.Caption = "You can lose four ways.  One way is to lose all of your StarShips, another is if you go into too much debt.  Meaning that your OreS are less then -499999.  Another is if you loose too many Space Men.  Almost Impossible though.  And if you loose too much ammo less then  - 31999"
End If
If Text1.Text = "How Do I Win" Then
Label1.Caption = "You can only win one way.  That is if you beat all of the planets and Ships.  Their life must be at 0."
End If
If Text1.Text = "Some Stragety" Then
Label1.Caption = "I will give you 2 stageties.  First is to create 2 starships.  Second is to click the special whenever your life is below 20."
End If
End Sub

Private Sub List1_Click()
Text1.Text = List1
End Sub
