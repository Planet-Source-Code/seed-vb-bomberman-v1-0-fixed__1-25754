VERSION 5.00
Begin VB.Form frmHighs 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores:"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   3960
      Picture         =   "frmHighs.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   720
      Picture         =   "frmHighs.frx":37C2
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Score1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   11
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Score1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2640
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Score1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   9
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Score1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   8
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Score1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Score1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Name1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Name1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Name1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Name1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Name1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Name1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmHighs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If JustLostGame = True Then
    response = MsgBox("Play again, champ?", vbYesNo + vbInformation, "Replay?")
        If response = vbYes Then
            frmMain.Bman.Picture = frmMain.BmanDown.Picture 'default
            frmMain.BManRemains.Visible = False 'clean up his bodily remains
            Hovering = False
            Slowed = False
            frmMain.DelayTimer.interval = 120
            NumFlames = 8
            frmMain.tmrStartingOff.Enabled = True
            frmMain.Bman.Top = 360
            frmMain.Bman.Left = 360
            CurrLives = 5
            BombCount = 1
            PlayerScore = 0
            JustLostGame = False
            frmMain.FreqScan.interval = 20
            frmMain.FreqScan.Enabled = True
            LoadLevel 1
            frmMain.EnemyMover.interval = 600
            frmMain.Show
            For tr = 0 To NumberOfEnemies
                frmMain.Enemy(tr).Tag = ""
            Next tr
            frmMain.FreqScan.Enabled = True
        Else
            Unload Me
            End
        End If
End If
Me.Hide
GotHighScore = False
End Sub

Private Sub Command2_Click()
response = MsgBox("Are you sure you want to clear all the high scores?", vbQuestion + vbYesNo + vbSystemModal, "Really?")
If response = vbYes Then
'clear .dat file and clear list
Text1.Text = ""
For pp = 1 To 6
Text1.Text = Text1.Text + "Noneë0" + vbCrLf
Next pp
Open App.Path + "\Highscores.dat" For Output As #1
Print #1, Text1.Text
Close #1
LoadScores frmHighs
'For uu = 0 To 5
'Name1(uu).Caption = "None"
'Score1(uu).Caption = "0"
'Next uu
End If
End Sub

Private Sub Form_Load()
LoadScores frmHighs
CurrLives = 5
For i = 0 To 5
    Name1(i).BackStyle = 0: Score1(i).BackStyle = 0
Next i
End Sub

