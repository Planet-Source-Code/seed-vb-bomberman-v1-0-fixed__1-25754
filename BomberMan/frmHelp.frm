VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1988
      TabIndex        =   10
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   548
      TabIndex        =   11
      Top             =   4200
      Width           =   4455
   End
   Begin VB.Image EnemyFreeze 
      Height          =   495
      Left            =   2880
      Picture         =   "frmHelp.frx":00EF
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Enhanced Bomb"
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Life"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Enemy Freezer"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Temporary Invincibility"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Hmm.. I don't know."
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Temporary Slow Mode"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "1 Extra Bomb"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed Up"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   2880
      Picture         =   "frmHelp.frx":0FA1
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   2880
      Picture         =   "frmHelp.frx":3DD3
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   360
      Picture         =   "frmHelp.frx":5BC5
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   360
      Picture         =   "frmHelp.frx":6EC7
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   360
      Picture         =   "frmHelp.frx":81C9
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2880
      Picture         =   "frmHelp.frx":8BCB
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   360
      Picture         =   "frmHelp.frx":9ECD
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Powerups:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":B1CF
      Height          =   855
      Left            =   615
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
frmMain.Show
frmMain.EnemyMover.Enabled = True
End Sub

