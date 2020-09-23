VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H0005CB05&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   0
      TabIndex        =   11
      Top             =   5760
      Width           =   6375
      Begin VB.Label lblPU 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Image PicDisp 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   2880
         Stretch         =   -1  'True
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   20
         ToolTipText     =   "Exit Bomberman"
         Top             =   480
         Width           =   855
      End
      Begin VB.Image Image8 
         Height          =   495
         Left            =   5280
         Picture         =   "frmMain.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Move me!"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   19
         ToolTipText     =   "Help"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image Image11 
         Height          =   615
         Left            =   0
         Picture         =   "frmMain.frx":030A
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblEnsLeft 
         BackStyle       =   0  'Transparent
         Caption         =   "X 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Score:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblLevel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   840
         TabIndex        =   14
         Top             =   0
         Width           =   495
      End
      Begin VB.Image Image10 
         Height          =   615
         Left            =   0
         Picture         =   "frmMain.frx":1C74
         Stretch         =   -1  'True
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblLives 
         BackStyle       =   0  'Transparent
         Caption         =   "X 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Pick-Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer tmrStartingOff 
      Interval        =   444
      Left            =   7800
      Top             =   0
   End
   Begin VB.PictureBox PRight 
      Height          =   375
      Left            =   7080
      Picture         =   "frmMain.frx":5436
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PLeft 
      Height          =   375
      Left            =   7440
      Picture         =   "frmMain.frx":87B8
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PUp 
      Height          =   375
      Left            =   6360
      Picture         =   "frmMain.frx":BB3A
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PDown 
      Height          =   375
      Left            =   6720
      Picture         =   "frmMain.frx":EEBC
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Seven 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   7800
      Top             =   1080
   End
   Begin VB.Timer BombDelay3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7800
      Top             =   1440
   End
   Begin VB.Timer BombDelay2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7440
      Top             =   0
   End
   Begin VB.Timer tmrEnGo 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   0
   End
   Begin VB.Timer Slug 
      Enabled         =   0   'False
      Interval        =   9000
      Left            =   6360
      Top             =   0
   End
   Begin VB.Timer EightTimer 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   7800
      Top             =   360
   End
   Begin VB.PictureBox BManRight 
      Height          =   375
      Left            =   7080
      Picture         =   "frmMain.frx":1223E
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox BManLeft 
      Height          =   375
      Left            =   6720
      Picture         =   "frmMain.frx":155C0
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox BManUp 
      Height          =   375
      Left            =   6360
      Picture         =   "frmMain.frx":18942
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox BmanDown 
      Height          =   375
      Left            =   7440
      Picture         =   "frmMain.frx":1BCC4
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer BombDelay 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7800
      Top             =   720
   End
   Begin VB.Timer FreqScan 
      Interval        =   20
      Left            =   7080
      Top             =   0
   End
   Begin VB.Timer EnemyMover 
      Interval        =   400
      Left            =   7440
      Top             =   360
   End
   Begin VB.Timer DelayTimer 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   7440
      Top             =   720
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Image Stairs 
      Height          =   375
      Left            =   360
      Picture         =   "frmMain.frx":1F046
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image TI 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":20348
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Bomb3 
      Height          =   375
      Left            =   6480
      Picture         =   "frmMain.frx":2164A
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Bomb2 
      Height          =   375
      Left            =   6480
      Picture         =   "frmMain.frx":2294C
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image ExtraBomb 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":23C4E
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   47
      Left            =   360
      Picture         =   "frmMain.frx":24F50
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   46
      Left            =   8280
      Picture         =   "frmMain.frx":268BA
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   45
      Left            =   5040
      Picture         =   "frmMain.frx":28224
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   44
      Left            =   8280
      Picture         =   "frmMain.frx":29B8E
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   43
      Left            =   2160
      Picture         =   "frmMain.frx":2B4F8
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   50
      Left            =   5040
      Picture         =   "frmMain.frx":2CE62
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   49
      Left            =   2520
      Picture         =   "frmMain.frx":2E7CC
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   48
      Left            =   360
      Picture         =   "frmMain.frx":30136
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   47
      Left            =   3960
      Picture         =   "frmMain.frx":31AA0
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   46
      Left            =   5400
      Picture         =   "frmMain.frx":3340A
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   45
      Left            =   720
      Picture         =   "frmMain.frx":34D74
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   42
      Left            =   1440
      Picture         =   "frmMain.frx":366DE
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   41
      Left            =   3600
      Picture         =   "frmMain.frx":38048
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   40
      Left            =   5400
      Picture         =   "frmMain.frx":399B2
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   39
      Left            =   5040
      Picture         =   "frmMain.frx":3B31C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   38
      Left            =   1440
      Picture         =   "frmMain.frx":3CC86
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   37
      Left            =   8280
      Picture         =   "frmMain.frx":3E5F0
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   36
      Left            =   7920
      Picture         =   "frmMain.frx":3FF5A
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   35
      Left            =   7560
      Picture         =   "frmMain.frx":418C4
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   44
      Left            =   3240
      Picture         =   "frmMain.frx":4322E
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   43
      Left            =   1080
      Picture         =   "frmMain.frx":44B98
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   42
      Left            =   3240
      Picture         =   "frmMain.frx":46502
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   41
      Left            =   2880
      Picture         =   "frmMain.frx":47E6C
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   40
      Left            =   3240
      Picture         =   "frmMain.frx":497D6
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   39
      Left            =   360
      Picture         =   "frmMain.frx":4B140
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   38
      Left            =   2520
      Picture         =   "frmMain.frx":4CAAA
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   37
      Left            =   2520
      Picture         =   "frmMain.frx":4E414
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   36
      Left            =   3600
      Picture         =   "frmMain.frx":4FD7E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   35
      Left            =   1800
      Picture         =   "frmMain.frx":516E8
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   34
      Left            =   3600
      Picture         =   "frmMain.frx":53052
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   33
      Left            =   4320
      Picture         =   "frmMain.frx":549BC
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   32
      Left            =   3600
      Picture         =   "frmMain.frx":56326
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Enemy 
      Height          =   375
      Index           =   7
      Left            =   360
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Enemy 
      Height          =   375
      Index           =   6
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   34
      Left            =   1080
      Picture         =   "frmMain.frx":57C90
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   33
      Left            =   3240
      Picture         =   "frmMain.frx":595FA
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   32
      Left            =   4680
      Picture         =   "frmMain.frx":5AF64
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   31
      Left            =   1080
      Picture         =   "frmMain.frx":5C8CE
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   31
      Left            =   2880
      Picture         =   "frmMain.frx":5E238
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image Heart 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":5FBA2
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   360
      Picture         =   "frmMain.frx":61994
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   5415
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   360
      Picture         =   "frmMain.frx":77446
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5415
   End
   Begin VB.Image Image2 
      Height          =   5775
      Left            =   5760
      Picture         =   "frmMain.frx":8CEF8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   5775
      Left            =   0
      Picture         =   "frmMain.frx":A60A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Hover 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":BF24C
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Slow 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":C2556
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image DeathPills 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":C3858
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   20
      Left            =   6120
      Picture         =   "frmMain.frx":C425A
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   19
      Left            =   6120
      Picture         =   "frmMain.frx":C555C
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   18
      Left            =   6120
      Picture         =   "frmMain.frx":C685E
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   17
      Left            =   6120
      Picture         =   "frmMain.frx":C7B60
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   16
      Left            =   6120
      Picture         =   "frmMain.frx":C8E62
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   15
      Left            =   6120
      Picture         =   "frmMain.frx":CA164
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   14
      Left            =   6120
      Picture         =   "frmMain.frx":CB466
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   13
      Left            =   6120
      Picture         =   "frmMain.frx":CC768
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image BiggerBomb 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":CDA6A
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   12
      Left            =   6120
      Picture         =   "frmMain.frx":CED6C
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   11
      Left            =   6120
      Picture         =   "frmMain.frx":D006E
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   10
      Left            =   6120
      Picture         =   "frmMain.frx":D1370
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   9
      Left            =   6120
      Picture         =   "frmMain.frx":D2672
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image EnemyFreeze 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":D3974
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   8
      Left            =   6120
      Picture         =   "frmMain.frx":D4826
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   7
      Left            =   6120
      Picture         =   "frmMain.frx":D5B28
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   6
      Left            =   6120
      Picture         =   "frmMain.frx":D6E2A
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   5
      Left            =   6120
      Picture         =   "frmMain.frx":D812C
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   4
      Left            =   6120
      Picture         =   "frmMain.frx":D942E
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   3
      Left            =   6120
      Picture         =   "frmMain.frx":DA730
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   2
      Left            =   6120
      Picture         =   "frmMain.frx":DBA32
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   1
      Left            =   6120
      Picture         =   "frmMain.frx":DCD34
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Flame 
      Height          =   375
      Index           =   0
      Left            =   6120
      Picture         =   "frmMain.frx":DE036
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Bomb1 
      Height          =   375
      Left            =   6480
      Picture         =   "frmMain.frx":DF338
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image En1 
      Height          =   375
      Left            =   6480
      Picture         =   "frmMain.frx":E063A
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   30
      Left            =   4320
      Picture         =   "frmMain.frx":E1FA4
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   29
      Left            =   4320
      Picture         =   "frmMain.frx":E390E
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   28
      Left            =   1800
      Picture         =   "frmMain.frx":E5278
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   27
      Left            =   5040
      Picture         =   "frmMain.frx":E6BE2
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   26
      Left            =   2880
      Picture         =   "frmMain.frx":E854C
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   25
      Left            =   2520
      Picture         =   "frmMain.frx":E9EB6
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   24
      Left            =   4680
      Picture         =   "frmMain.frx":EB820
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   23
      Left            =   2520
      Picture         =   "frmMain.frx":ED18A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   22
      Left            =   720
      Picture         =   "frmMain.frx":EEAF4
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   21
      Left            =   4320
      Picture         =   "frmMain.frx":F045E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   20
      Left            =   2160
      Picture         =   "frmMain.frx":F1DC8
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   19
      Left            =   3240
      Picture         =   "frmMain.frx":F3732
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   18
      Left            =   720
      Picture         =   "frmMain.frx":F509C
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   17
      Left            =   720
      Picture         =   "frmMain.frx":F6A06
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   16
      Left            =   4320
      Picture         =   "frmMain.frx":F8370
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   15
      Left            =   3960
      Picture         =   "frmMain.frx":F9CDA
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   14
      Left            =   1440
      Picture         =   "frmMain.frx":FB644
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   13
      Left            =   3600
      Picture         =   "frmMain.frx":FCFAE
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   12
      Left            =   720
      Picture         =   "frmMain.frx":FE918
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   11
      Left            =   1080
      Picture         =   "frmMain.frx":100282
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   10
      Left            =   4320
      Picture         =   "frmMain.frx":101BEC
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   9
      Left            =   1800
      Picture         =   "frmMain.frx":103556
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   8
      Left            =   5040
      Picture         =   "frmMain.frx":104EC0
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   7
      Left            =   2160
      Picture         =   "frmMain.frx":10682A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   6
      Left            =   5400
      Picture         =   "frmMain.frx":108194
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   5
      Left            =   4320
      Picture         =   "frmMain.frx":109AFE
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   4
      Left            =   1080
      Picture         =   "frmMain.frx":10B468
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   3
      Left            =   720
      Picture         =   "frmMain.frx":10CDD2
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   2
      Left            =   360
      Picture         =   "frmMain.frx":10E73C
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   1
      Left            =   1440
      Picture         =   "frmMain.frx":1100A6
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Rock 
      Height          =   375
      Index           =   0
      Left            =   360
      Picture         =   "frmMain.frx":111A10
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblBonus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblBonus"
      BeginProperty Font 
         Name            =   "Tribune"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   30
      Left            =   7560
      Picture         =   "frmMain.frx":11337A
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   29
      Left            =   1080
      Picture         =   "frmMain.frx":114CE4
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   28
      Left            =   2880
      Picture         =   "frmMain.frx":11664E
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   27
      Left            =   3240
      Picture         =   "frmMain.frx":117FB8
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   26
      Left            =   4680
      Picture         =   "frmMain.frx":119922
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   25
      Left            =   3240
      Picture         =   "frmMain.frx":11B28C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   24
      Left            =   7920
      Picture         =   "frmMain.frx":11CBF6
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   23
      Left            =   5400
      Picture         =   "frmMain.frx":11E560
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   22
      Left            =   5040
      Picture         =   "frmMain.frx":11FECA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   21
      Left            =   3960
      Picture         =   "frmMain.frx":121834
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   20
      Left            =   2160
      Picture         =   "frmMain.frx":12319E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   19
      Left            =   4680
      Picture         =   "frmMain.frx":124B08
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   18
      Left            =   8280
      Picture         =   "frmMain.frx":126472
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   17
      Left            =   3240
      Picture         =   "frmMain.frx":127DDC
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   16
      Left            =   360
      Picture         =   "frmMain.frx":129746
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   15
      Left            =   7560
      Picture         =   "frmMain.frx":12B0B0
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   14
      Left            =   7920
      Picture         =   "frmMain.frx":12CA1A
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   13
      Left            =   2160
      Picture         =   "frmMain.frx":12E384
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   12
      Left            =   1080
      Picture         =   "frmMain.frx":12FCEE
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   11
      Left            =   2520
      Picture         =   "frmMain.frx":131658
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   10
      Left            =   7920
      Picture         =   "frmMain.frx":132FC2
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   9
      Left            =   7920
      Picture         =   "frmMain.frx":13492C
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   8
      Left            =   8280
      Picture         =   "frmMain.frx":136296
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   7
      Left            =   1800
      Picture         =   "frmMain.frx":137C00
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   6
      Left            =   2880
      Picture         =   "frmMain.frx":13956A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   5
      Left            =   3600
      Picture         =   "frmMain.frx":13AED4
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   4
      Left            =   7920
      Picture         =   "frmMain.frx":13C83E
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   3
      Left            =   4680
      Picture         =   "frmMain.frx":13E1A8
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   2
      Left            =   720
      Picture         =   "frmMain.frx":13FB12
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   1
      Left            =   7560
      Picture         =   "frmMain.frx":14147C
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Weak 
      Height          =   375
      Index           =   0
      Left            =   5400
      Picture         =   "frmMain.frx":142DE6
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image SpeedUp 
      Height          =   375
      Left            =   6120
      Picture         =   "frmMain.frx":144750
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Enemy 
      Height          =   375
      Index           =   5
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image Enemy 
      Height          =   375
      Index           =   4
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Enemy 
      Height          =   375
      Index           =   3
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Enemy 
      Height          =   375
      Index           =   2
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Enemy 
      Height          =   375
      Index           =   1
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image Enemy 
      Height          =   375
      Index           =   0
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Bman 
      Height          =   375
      Left            =   360
      Stretch         =   -1  'True
      Top             =   360
      Width           =   375
   End
   Begin VB.OLE EnemyDie 
      Class           =   "SoundRec"
      Height          =   375
      Left            =   7080
      OleObjectBlob   =   "frmMain.frx":147582
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OLE BManDie 
      BackColor       =   &H0005CB05&
      Class           =   "SoundRec"
      Height          =   375
      Left            =   6840
      OleObjectBlob   =   "frmMain.frx":14AF9A
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image BManRemains 
      Height          =   375
      Left            =   6840
      Picture         =   "frmMain.frx":14F5B2
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Visible         =   0   'False
      Begin VB.Menu mnuHTP 
         Caption         =   "How To Play"
      End
      Begin VB.Menu mnULINE 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IMMEDIATE
'--need to add 4 more levels

Dim NoUp As Boolean
Dim NoDown As Boolean
Dim NoLeft As Boolean
Dim NoRight As Boolean
Dim Release As Boolean
Dim MaxedOut4Bombs As Boolean

Dim Direc As String
Dim EnDirec As String
Dim LastDir(15) As String

Dim NumberOfWeakBlocks As Integer
Dim NumberOfRocks As Integer
Dim GreaterVal As Integer
Dim RandNum1 As Integer
Dim RandNum2 As Integer
Dim Choices As Integer
Dim StartLeft As Integer
Dim StartTop As Integer
Dim Ched As Integer
Dim Chex As Integer
Dim Chev As Integer
Dim Chey As Integer
Dim EFCount As Integer
Dim MaxLevels As Integer

Sub Pause(inter)
Current = Timer
Do While Timer - Current < Val(inter)
DoEvents
Loop
End Sub

Private Sub BombDelay_Timer()
Bomb1.Visible = False
Flame(0).Left = Bomb1.Left: Flame(0).Top = Bomb1.Top: Flame(0).Visible = True
'put the default set of flames in place:
Flame(1).Left = Bomb1.Left - 360: Flame(1).Top = Bomb1.Top
Flame(2).Top = Bomb1.Top - 360: Flame(2).Left = Bomb1.Left
Flame(3).Left = Bomb1.Left + 360: Flame(3).Top = Bomb1.Top
Flame(4).Top = Bomb1.Top + 360: Flame(4).Left = Bomb1.Left
Flame(5).Left = Bomb1.Left - 720: Flame(5).Top = Bomb1.Top
Flame(6).Top = Bomb1.Top - 720: Flame(6).Left = Bomb1.Left
Flame(7).Left = Bomb1.Left + 720: Flame(7).Top = Bomb1.Top
Flame(8).Top = Bomb1.Top + 720: Flame(8).Left = Bomb1.Left
If NumFlames >= 12 Then 'if bomb is powered up by 1.
Flame(9).Left = Bomb1.Left - 1080: Flame(9).Top = Bomb1.Top
Flame(10).Top = Bomb1.Top - 1080: Flame(10).Left = Bomb1.Left
Flame(11).Left = Bomb1.Left + 1080: Flame(11).Top = Bomb1.Top
Flame(12).Top = Bomb1.Top + 1080: Flame(12).Left = Bomb1.Left
End If
If NumFlames >= 16 Then 'if bomb has been pumped 2 times:
Flame(13).Left = Bomb1.Left - 1440: Flame(13).Top = Bomb1.Top
Flame(14).Top = Bomb1.Top - 1440: Flame(14).Left = Bomb1.Left
Flame(15).Left = Bomb1.Left + 1440: Flame(15).Top = Bomb1.Top
Flame(16).Top = Bomb1.Top + 1440: Flame(16).Left = Bomb1.Left
End If
If NumFlames = 20 Then 'if bomb is FULLY pumped (watch out!)
Flame(17).Left = Bomb1.Left - 1800: Flame(17).Top = Bomb1.Top
Flame(18).Top = Bomb1.Top - 1800: Flame(18).Left = Bomb1.Left
Flame(19).Left = Bomb1.Left + 1800: Flame(19).Top = Bomb1.Top
Flame(20).Top = Bomb1.Top + 1800: Flame(20).Left = Bomb1.Left
End If
'now see if a collision occured:
'(this is probably the part that will make the game slow down for people with slower computers...especially if you have the bomb charged a lot)
For i = 0 To NumFlames 'for each flame
 For j = 0 To NumberOfWeakBlocks ' for each weak block
  For l = 0 To NumberOfRocks
   If Flame(i).Tag <> "X" And Flame(i).Left = Rock(l).Left And Flame(i).Top = Rock(l).Top Then
    Flame(i).Tag = "X" '-will not appear
    On Error Resume Next
    Flame(i + 4).Tag = "X" 'next one won't either
    Flame(i + 8).Tag = "X" 'the next one too (if applicable)
    Flame(i + 12).Tag = "X" '''
    Flame(i + 16).Tag = "X" '''
   End If
   If Flame(i).Tag <> "X" And Flame(i).Left = Weak(j).Left And Flame(i).Top = Weak(j).Top Then
    Flame(i).Tag = "X" '-will not appear!
    Weak(j).Tag = "X"
    On Error Resume Next
    Flame(i + 4).Tag = "X" '-make sure that outermore flames also don't appear
    Flame(i + 8).Tag = "X" 'even more outer ones dont appear or have any effect
    Flame(i + 12).Tag = "X" '''
    Flame(i + 16).Tag = "X" '''
   End If
  Next l
 Next j
Next i
'set the appropriate flames as visible (finally!):
For n = 0 To NumFlames
 If Flame(n).Tag = "" Then Flame(n).Visible = True
Next n
'dissolve all the hit weak blocks and MAYBE give a powerup (or powerdown.. hehe)
For p = 0 To NumberOfWeakBlocks
 If Weak(p).Tag = "X" Then
  RandNum2 = Int(Rnd * 40)
   If RandNum2 = 0 Or RandNum2 = 1 Or RandNum2 = 2 Then
    If Slowed = True Or SpeedUp.Top = -1000 Or SpeedUp.Visible = True Then GoTo AlreadyGotIt  'already got it once or it's already there.
     SpeedUp.Top = Weak(p).Top
     SpeedUp.Left = Weak(p).Left
     SpeedUp.Visible = True
   ElseIf RandNum2 = 3 Or RandNum2 = 4 Then
    If EnemyFreeze.Top = -1000 Or EnemyFreeze.Visible = True Then GoTo AlreadyGotIt
     EnemyFreeze.Top = Weak(p).Top
     EnemyFreeze.Left = Weak(p).Left
     EnemyFreeze.Visible = True
   ElseIf RandNum2 = 5 Or RandNum2 = 6 Or RandNum2 = 7 Then
    If BiggerBomb.Visible = True Or BiggerBomb.Top = -2000 Then GoTo AlreadyGotIt
     If BiggerBomb.Top = -1000 And EFCount <= 4 Then
      If EFCount = 3 Then
      GoTo AlreadyGotIt
      End If
      BiggerBomb.Top = Weak(p).Top
      BiggerBomb.Left = Weak(p).Left
      BiggerBomb.Visible = True
     End If
   ElseIf RandNum2 = 8 Then
    If DeathPills.Top = -1000 Or DeathPills.Visible = True Then GoTo AlreadyGotIt
    DeathPills.Top = Weak(p).Top
    DeathPills.Left = Weak(p).Left
    DeathPills.Visible = True
   ElseIf RandNum2 = 9 Or RandNum2 = 10 Then
    If Slow.Top = -1000 Or Slow.Visible = True Then GoTo AlreadyGotIt
    Slow.Top = Weak(p).Top
    Slow.Left = Weak(p).Left
    Slow.Visible = True
    'HOVER BOOTS are a HUGE advantage.  they may be too good.  if u want to add them just add use the next 5 lines of code in each sections!
   'ElseIf RandNum2 = 11 Then
    'If Hover.Top = -1000 Or Hover.Visible = True Then GoTo AlreadyGotIt
    'Hover.Top = Weak(p).Top
    'Hover.Left = Weak(p).Left
    'Hover.Visible = True
   ElseIf RandNum2 = 12 Then
    If Heart.Top = -1000 Or Heart.Visible = True Then GoTo AlreadyGotIt
    Heart.Top = Weak(p).Top
    Heart.Left = Weak(p).Left
    Heart.Visible = True
   ElseIf RandNum2 > 12 And RandNum2 < 16 Then 'half the time nothing happens at all...
    If MaxedOut4Bombs = True Or ExtraBomb.Visible = True Then GoTo AlreadyGotIt
    ExtraBomb.Top = Weak(p).Top
    ExtraBomb.Left = Weak(p).Left
    ExtraBomb.Visible = True
   ElseIf RandNum2 > 15 And RandNum2 < 18 Then
    If TI.Top = -1000 Or TI.Visible = True Then GoTo AlreadyGotIt
    TI.Top = Weak(p).Top
    TI.Left = Weak(p).Left
    TI.Visible = True
   Else
   GoTo AlreadyGotIt 'so u only get something less than half the time
   End If
AlreadyGotIt: 'you can only get one of some of them once each level.
  Weak(p).Top = -1000 'see ya!
  PlayerScore = PlayerScore + 25
End If
Next p
Leave:
'clear tags so it works more than once:
If NumberOfEnemies > NumberOfWeakBlocks Then GreaterVal = NumberOfEnemies
If NumberOfEnemies < NumberOfWeakBlocks Then GreaterVal = NumberOfWeakBlocks
For q = 0 To GreaterVal
 On Error Resume Next
 Weak(q).Tag = ""
 Flame(q).Tag = ""
Next q
'set flame delay - so it looks like a flash of flame:
PlayWav App.Path + "\Sounds\bombexplode3.wav"
Pause 0.6 'momentary pause so u can at least see them for a lil' while. (on a slow pc, the bomb will disappear then a sec later the flames will show.  this is just the comp thinking thru all those arduous for next statements above...)
For i = 0 To 20
'plz notice that i want to go the full 20, even though all of them may not even be activated.  the reason is cause sometimes when you killed yourself w/ big bombs, some of the flames would remain onscreen because when bman dies it resets the flamecount.  k.
 Flame(i).Visible = False
Next i
BombDelay.Enabled = False
End Sub

Private Sub BombDelay2_Timer()
Bomb2.Visible = False
Flame(0).Left = Bomb2.Left: Flame(0).Top = Bomb2.Top: Flame(0).Visible = True
'put the default set of flames in place:
Flame(1).Left = Bomb2.Left - 360: Flame(1).Top = Bomb2.Top
Flame(2).Top = Bomb2.Top - 360: Flame(2).Left = Bomb2.Left
Flame(3).Left = Bomb2.Left + 360: Flame(3).Top = Bomb2.Top
Flame(4).Top = Bomb2.Top + 360: Flame(4).Left = Bomb2.Left
Flame(5).Left = Bomb2.Left - 720: Flame(5).Top = Bomb2.Top
Flame(6).Top = Bomb2.Top - 720: Flame(6).Left = Bomb2.Left
Flame(7).Left = Bomb2.Left + 720: Flame(7).Top = Bomb2.Top
Flame(8).Top = Bomb2.Top + 720: Flame(8).Left = Bomb2.Left
If NumFlames >= 12 Then 'if bomb is powered up by 1.
Flame(9).Left = Bomb2.Left - 1080: Flame(9).Top = Bomb2.Top
Flame(10).Top = Bomb2.Top - 1080: Flame(10).Left = Bomb2.Left
Flame(11).Left = Bomb2.Left + 1080: Flame(11).Top = Bomb2.Top
Flame(12).Top = Bomb2.Top + 1080: Flame(12).Left = Bomb2.Left
End If
If NumFlames >= 16 Then 'if bomb has been pumped 2 times:
Flame(13).Left = Bomb2.Left - 1440: Flame(13).Top = Bomb2.Top
Flame(14).Top = Bomb2.Top - 1440: Flame(14).Left = Bomb2.Left
Flame(15).Left = Bomb2.Left + 1440: Flame(15).Top = Bomb2.Top
Flame(16).Top = Bomb2.Top + 1440: Flame(16).Left = Bomb2.Left
End If
If NumFlames = 20 Then 'if bomb is FULLY pumped (watch out!)
Flame(17).Left = Bomb2.Left - 1800: Flame(17).Top = Bomb2.Top
Flame(18).Top = Bomb2.Top - 1800: Flame(18).Left = Bomb2.Left
Flame(19).Left = Bomb2.Left + 1800: Flame(19).Top = Bomb2.Top
Flame(20).Top = Bomb2.Top + 1800: Flame(20).Left = Bomb2.Left
End If
'now see if a collision occured:
'(this is probably the part that will make the game slow down for people with slower computers...especially if you have the bomb charged a lot)
For i = 0 To NumFlames 'for each flame
 For j = 0 To NumberOfWeakBlocks ' for each weak block
  For l = 0 To NumberOfRocks
   If Flame(i).Tag <> "X" And Flame(i).Left = Rock(l).Left And Flame(i).Top = Rock(l).Top Then
    Flame(i).Tag = "X" '-will not appear
    On Error Resume Next
    Flame(i + 4).Tag = "X" 'next one won't either
    Flame(i + 8).Tag = "X" 'the next one too (if applicable)
    Flame(i + 12).Tag = "X" '''
    Flame(i + 16).Tag = "X" '''
   End If
   If Not Flame(i).Tag = "X" And Flame(i).Left = Weak(j).Left And Flame(i).Top = Weak(j).Top Then
    Flame(i).Tag = "X" '-will not appear!
    Weak(j).Tag = "X"
    On Error Resume Next
    Flame(i + 4).Tag = "X" '-make sure that outermore flames also don't appear
    Flame(i + 8).Tag = "X" 'even more outer ones dont appear or have any effect
    Flame(i + 12).Tag = "X" '''
    Flame(i + 16).Tag = "X" '''
   End If
  Next l
 Next j
Next i
'set the appropriate flames as visible (finally!):
For n = 0 To NumFlames
 If Flame(n).Tag = "" Then Flame(n).Visible = True
Next n
'dissolve all the hit weak blocks and MAYBE give a powerup (or powerdown.. hehe)
For p = 0 To NumberOfWeakBlocks
 If Weak(p).Tag = "X" Then
  RandNum2 = Int(Rnd * 40)
   If RandNum2 = 0 Or RandNum2 = 1 Or RandNum2 = 2 Then
    If Slowed = True Or SpeedUp.Top = -1000 Or SpeedUp.Visible = True Then GoTo AlreadyGotIt  'already got it once or it's already there.
     SpeedUp.Top = Weak(p).Top
     SpeedUp.Left = Weak(p).Left
     SpeedUp.Visible = True
   ElseIf RandNum2 = 3 Or RandNum2 = 4 Then
    If EnemyFreeze.Top = -1000 Or EnemyFreeze.Visible = True Then GoTo AlreadyGotIt
     EnemyFreeze.Top = Weak(p).Top
     EnemyFreeze.Left = Weak(p).Left
     EnemyFreeze.Visible = True
   ElseIf RandNum2 = 5 Or RandNum2 = 6 Or RandNum2 = 7 Then
    If BiggerBomb.Visible = True Or BiggerBomb.Top = -2000 Then GoTo AlreadyGotIt
     If BiggerBomb.Top = -1000 And EFCount <= 4 Then
      If EFCount = 3 Then
      GoTo AlreadyGotIt
      End If
      BiggerBomb.Top = Weak(p).Top
      BiggerBomb.Left = Weak(p).Left
      BiggerBomb.Visible = True
     End If
   ElseIf RandNum2 = 8 Then
    If DeathPills.Top = -1000 Or DeathPills.Visible = True Then GoTo AlreadyGotIt
    DeathPills.Top = Weak(p).Top
    DeathPills.Left = Weak(p).Left
    DeathPills.Visible = True
   ElseIf RandNum2 = 9 Or RandNum2 = 10 Then
    If Slow.Top = -1000 Or Slow.Visible = True Then GoTo AlreadyGotIt
    Slow.Top = Weak(p).Top
    Slow.Left = Weak(p).Left
    Slow.Visible = True
    'HOVER BOOTS are a HUGE advantage.  they may be too good.  if u want to add them just add use the next 5 lines of code in each sections!
   'ElseIf RandNum2 = 11 Then
    'If Hover.Top = -1000 Or Hover.Visible = True Then GoTo AlreadyGotIt
    'Hover.Top = Weak(p).Top
    'Hover.Left = Weak(p).Left
    'Hover.Visible = True
   ElseIf RandNum2 = 12 Then
    If Heart.Top = -1000 Or Heart.Visible = True Then GoTo AlreadyGotIt
    Heart.Top = Weak(p).Top
    Heart.Left = Weak(p).Left
    Heart.Visible = True
   ElseIf RandNum2 > 12 And RandNum2 < 16 Then 'half the time nothing happens at all...
    If MaxedOut4Bombs = True Or ExtraBomb.Visible = True Then GoTo AlreadyGotIt
    ExtraBomb.Top = Weak(p).Top
    ExtraBomb.Left = Weak(p).Left
    ExtraBomb.Visible = True
   ElseIf RandNum2 > 15 And RandNum2 < 18 Then
    If TI.Top = -1000 Or TI.Visible = True Then GoTo AlreadyGotIt
    TI.Top = Weak(p).Top
    TI.Left = Weak(p).Left
    TI.Visible = True
   Else
   GoTo AlreadyGotIt 'so u only get something less than half the time
   End If
AlreadyGotIt: 'you can only get one of some of them once each level.
  Weak(p).Top = -1000 'see ya!
  PlayerScore = PlayerScore + 25
End If
Next p
Leave:
'clear tags so it works more than once:
If NumberOfEnemies > NumberOfWeakBlocks Then GreaterVal = NumberOfEnemies
If NumberOfEnemies < NumberOfWeakBlocks Then GreaterVal = NumberOfWeakBlocks
For q = 0 To GreaterVal
 On Error Resume Next
 Weak(q).Tag = ""
 Flame(q).Tag = ""
Next q
'set flame delay - so it looks like a flash of flame:
PlayWav App.Path + "\Sounds\bombexplode3.wav"
Pause 0.6 'momentary pause so u can at least see them for a lil' while. (on a slow pc, the bomb will disappear then a sec later the flames will show.  this is just the comp thinking thru all those arduous for next statements above...)
For i = 0 To 20 'rid flames
 Flame(i).Visible = False
Next i
BombDelay2.Enabled = False
End Sub

Private Sub BombDelay3_Timer()
Bomb3.Visible = False
Flame(0).Left = Bomb3.Left: Flame(0).Top = Bomb3.Top: Flame(0).Visible = True
'put the default set of flames in place:
Flame(1).Left = Bomb3.Left - 360: Flame(1).Top = Bomb3.Top
Flame(2).Top = Bomb3.Top - 360: Flame(2).Left = Bomb3.Left
Flame(3).Left = Bomb3.Left + 360: Flame(3).Top = Bomb3.Top
Flame(4).Top = Bomb3.Top + 360: Flame(4).Left = Bomb3.Left
Flame(5).Left = Bomb3.Left - 720: Flame(5).Top = Bomb3.Top
Flame(6).Top = Bomb3.Top - 720: Flame(6).Left = Bomb3.Left
Flame(7).Left = Bomb3.Left + 720: Flame(7).Top = Bomb3.Top
Flame(8).Top = Bomb3.Top + 720: Flame(8).Left = Bomb3.Left
If NumFlames >= 12 Then 'if bomb is powered up by 1.
Flame(9).Left = Bomb3.Left - 1080: Flame(9).Top = Bomb3.Top
Flame(10).Top = Bomb3.Top - 1080: Flame(10).Left = Bomb3.Left
Flame(11).Left = Bomb3.Left + 1080: Flame(11).Top = Bomb3.Top
Flame(12).Top = Bomb3.Top + 1080: Flame(12).Left = Bomb3.Left
End If
If NumFlames >= 16 Then 'if bomb has been pumped 2 times:
Flame(13).Left = Bomb3.Left - 1440: Flame(13).Top = Bomb3.Top
Flame(14).Top = Bomb3.Top - 1440: Flame(14).Left = Bomb3.Left
Flame(15).Left = Bomb3.Left + 1440: Flame(15).Top = Bomb3.Top
Flame(16).Top = Bomb3.Top + 1440: Flame(16).Left = Bomb3.Left
End If
If NumFlames = 20 Then 'if bomb is FULLY pumped (watch out!)
Flame(17).Left = Bomb3.Left - 1800: Flame(17).Top = Bomb3.Top
Flame(18).Top = Bomb3.Top - 1800: Flame(18).Left = Bomb3.Left
Flame(19).Left = Bomb3.Left + 1800: Flame(19).Top = Bomb3.Top
Flame(20).Top = Bomb3.Top + 1800: Flame(20).Left = Bomb3.Left
End If
'now see if a collision occured:
'(this is probably the part that will make the game slow down for people with slower computers...especially if you have the bomb charged a lot)
For i = 0 To NumFlames 'for each flame
 For j = 0 To NumberOfWeakBlocks ' for each weak block
  For l = 0 To NumberOfRocks
   If Flame(i).Tag <> "X" And Flame(i).Left = Rock(l).Left And Flame(i).Top = Rock(l).Top Then
    Flame(i).Tag = "X" '-will not appear
    On Error Resume Next
    Flame(i + 4).Tag = "X" 'next one won't either
    Flame(i + 8).Tag = "X" 'the next one too (if applicable)
    Flame(i + 12).Tag = "X" '''
    Flame(i + 16).Tag = "X" '''
   End If
   If Not Flame(i).Tag = "X" And Flame(i).Left = Weak(j).Left And Flame(i).Top = Weak(j).Top Then
    Flame(i).Tag = "X" '-will not appear!
    Weak(j).Tag = "X"
    On Error Resume Next
    Flame(i + 4).Tag = "X" '-make sure that outermore flames also don't appear
    Flame(i + 8).Tag = "X" 'even more outer ones dont appear or have any effect
    Flame(i + 12).Tag = "X" '''
    Flame(i + 16).Tag = "X" '''
   End If
  Next l
 Next j
Next i
'set the appropriate flames as visible (finally!):
For n = 0 To NumFlames
 If Flame(n).Tag = "" Then Flame(n).Visible = True
Next n
'dissolve all the hit weak blocks and MAYBE give a powerup (or powerdown.. hehe)
For p = 0 To NumberOfWeakBlocks
 If Weak(p).Tag = "X" Then
  RandNum2 = Int(Rnd * 40)
   If RandNum2 = 0 Or RandNum2 = 1 Or RandNum2 = 2 Then
    If Slowed = True Or SpeedUp.Top = -1000 Or SpeedUp.Visible = True Then GoTo AlreadyGotIt  'already got it once or it's already there.
     SpeedUp.Top = Weak(p).Top
     SpeedUp.Left = Weak(p).Left
     SpeedUp.Visible = True
   ElseIf RandNum2 = 3 Or RandNum2 = 4 Then
    If EnemyFreeze.Top = -1000 Or EnemyFreeze.Visible = True Then GoTo AlreadyGotIt
     EnemyFreeze.Top = Weak(p).Top
     EnemyFreeze.Left = Weak(p).Left
     EnemyFreeze.Visible = True
   ElseIf RandNum2 = 5 Or RandNum2 = 6 Or RandNum2 = 7 Then
    If BiggerBomb.Visible = True Or BiggerBomb.Top = -2000 Then GoTo AlreadyGotIt
     If BiggerBomb.Top = -1000 And EFCount <= 4 Then
      If EFCount = 3 Then
      GoTo AlreadyGotIt
      End If
      BiggerBomb.Top = Weak(p).Top
      BiggerBomb.Left = Weak(p).Left
      BiggerBomb.Visible = True
     End If
   ElseIf RandNum2 = 8 Then
    If DeathPills.Top = -1000 Or DeathPills.Visible = True Then GoTo AlreadyGotIt
    DeathPills.Top = Weak(p).Top
    DeathPills.Left = Weak(p).Left
    DeathPills.Visible = True
   ElseIf RandNum2 = 9 Or RandNum2 = 10 Then
    If Slow.Top = -1000 Or Slow.Visible = True Then GoTo AlreadyGotIt
    Slow.Top = Weak(p).Top
    Slow.Left = Weak(p).Left
    Slow.Visible = True
    'HOVER BOOTS are a HUGE advantage.  they may be too good.  if u want to add them just add use the next 5 lines of code in each sections!
   'ElseIf RandNum2 = 11 Then
    'If Hover.Top = -1000 Or Hover.Visible = True Then GoTo AlreadyGotIt
    'Hover.Top = Weak(p).Top
    'Hover.Left = Weak(p).Left
    'Hover.Visible = True
   ElseIf RandNum2 = 12 Then
    If Heart.Top = -1000 Or Heart.Visible = True Then GoTo AlreadyGotIt
    Heart.Top = Weak(p).Top
    Heart.Left = Weak(p).Left
    Heart.Visible = True
   ElseIf RandNum2 > 12 And RandNum2 < 16 Then 'half the time nothing happens at all...
    If MaxedOut4Bombs = True Or ExtraBomb.Visible = True Then GoTo AlreadyGotIt
    ExtraBomb.Top = Weak(p).Top
    ExtraBomb.Left = Weak(p).Left
    ExtraBomb.Visible = True
   ElseIf RandNum2 > 15 And RandNum2 < 18 Then
    If TI.Top = -1000 Or TI.Visible = True Then GoTo AlreadyGotIt
    TI.Top = Weak(p).Top
    TI.Left = Weak(p).Left
    TI.Visible = True
   Else
   GoTo AlreadyGotIt 'so u only get something less than half the time
   End If
AlreadyGotIt: 'you can only get one of some of them once each level.
  Weak(p).Top = -1000 'see ya!
  PlayerScore = PlayerScore + 25
End If
Next p
Leave:
'clear tags so it works more than once:
If NumberOfEnemies > NumberOfWeakBlocks Then GreaterVal = NumberOfEnemies
If NumberOfEnemies < NumberOfWeakBlocks Then GreaterVal = NumberOfWeakBlocks
For q = 0 To GreaterVal
 On Error Resume Next
 Weak(q).Tag = ""
 Flame(q).Tag = ""
Next q
'set flame delay - so it looks like a flash of flame:
PlayWav App.Path + "\Sounds\bombexplode3.wav"
Pause 0.6 'momentary pause so u can at least see them for a lil' while. (on a slow pc, the bomb will disappear then a sec later the flames will show.  this is just the comp thinking thru all those arduous for next statements above...)
For i = 0 To 20 'rid flames
 Flame(i).Visible = False
Next i
BombDelay3.Enabled = False
End Sub

Private Sub DelayTimer_Timer()
DelayTimer.Enabled = False
End Sub

Private Sub EightTimer_Timer()
Ched = Ched + 1
If Ched = 2 Then
EnemyMover.Enabled = True
Ched = 0
EightTimer.Enabled = False
Else: Exit Sub
End If
End Sub

Private Sub EnemyMover_Timer()
'The basic enemy AI (actually, their movements aren't smart
'as they don't seek bomberman.)  They sort of just drift around.
'Since, in BomberMan, the danger is not only from the enemies
'but also from the bombs, the enemies aren't supposed
'to home for bomberman.  Otherwise it would be too hard!
For k = 0 To NumberOfEnemies
If Enemy(k).Tag = "Dead" Then GoTo DeadSoSkip
 Choices = 4 'reset main variable
 Enemy(k).Tag = "" 'reset co-variable(s)
 For l = 0 To NumberOfRocks
  If Enemy(k).Top - Rock(l).Top = 360 And Enemy(k).Left = Rock(l).Left Then Enemy(k).Tag = Enemy(k).Tag + "NoUp"
  If Rock(l).Top - Enemy(k).Top = 360 And Enemy(k).Left = Rock(l).Left Then Enemy(k).Tag = Enemy(k).Tag + "NoDown"
  If Enemy(k).Left - Rock(l).Left = 360 And Enemy(k).Top = Rock(l).Top Then Enemy(k).Tag = Enemy(k).Tag + "NoLeft"
  If Rock(l).Left - Enemy(k).Left = 360 And Enemy(k).Top = Rock(l).Top Then Enemy(k).Tag = Enemy(k).Tag + "NoRight"
 Next l
 For m = 0 To NumberOfWeakBlocks
  If Enemy(k).Top - Weak(m).Top = 360 And Enemy(k).Left = Weak(m).Left Then Enemy(k).Tag = Enemy(k).Tag + "NoUp"
  If Weak(m).Top - Enemy(k).Top = 360 And Enemy(k).Left = Weak(m).Left Then Enemy(k).Tag = Enemy(k).Tag + "NoDown"
  If Enemy(k).Left - Weak(m).Left = 360 And Enemy(k).Top = Weak(m).Top Then Enemy(k).Tag = Enemy(k).Tag + "NoLeft"
  If Weak(m).Left - Enemy(k).Left = 360 And Enemy(k).Top = Weak(m).Top Then Enemy(k).Tag = Enemy(k).Tag + "NoRight"
 Next m
 If Enemy(k).Left = 360 Then Enemy(k).Tag = Enemy(k).Tag + "NoLeft"
 If Enemy(k).Top = 360 Then Enemy(k).Tag = Enemy(k).Tag + "NoUp"
 If Enemy(k).Left = 5400 Then Enemy(k).Tag = Enemy(k).Tag + "NoRight"
 If Enemy(k).Top = 5040 Then Enemy(k).Tag = Enemy(k).Tag + "NoDown"
 'enemy/bomb collision detection:
 If Bomb1.Visible = True And Enemy(k).Top - Bomb1.Top = 360 And Enemy(k).Left = Bomb1.Left Then Enemy(k).Tag = Enemy(k).Tag + "NoUp"
 If Bomb1.Visible = True And Bomb1.Top - Enemy(k).Top = 360 And Enemy(k).Left = Bomb1.Left Then Enemy(k).Tag = Enemy(k).Tag + "NoDown"
 If Bomb1.Visible = True And Enemy(k).Left - Bomb1.Left = 360 And Enemy(k).Top = Bomb1.Top Then Enemy(k).Tag = Enemy(k).Tag + "NoLeft"
 If Bomb1.Visible = True And Bomb1.Left - Enemy(k).Left = 360 And Enemy(k).Top = Bomb1.Top Then Enemy(k).Tag = Enemy(k).Tag + "NoRight"
 If Bomb2.Visible = True And Enemy(k).Top - Bomb2.Top = 360 And Enemy(k).Left = Bomb2.Left Then Enemy(k).Tag = Enemy(k).Tag + "NoUp"
 If Bomb2.Visible = True And Bomb2.Top - Enemy(k).Top = 360 And Enemy(k).Left = Bomb2.Left Then Enemy(k).Tag = Enemy(k).Tag + "NoDown"
 If Bomb2.Visible = True And Enemy(k).Left - Bomb2.Left = 360 And Enemy(k).Top = Bomb2.Top Then Enemy(k).Tag = Enemy(k).Tag + "NoLeft"
 If Bomb2.Visible = True And Bomb2.Left - Enemy(k).Left = 360 And Enemy(k).Top = Bomb2.Top Then Enemy(k).Tag = Enemy(k).Tag + "NoRight"
 If Bomb3.Visible = True And Enemy(k).Top - Bomb3.Top = 360 And Enemy(k).Left = Bomb3.Left Then Enemy(k).Tag = Enemy(k).Tag + "NoUp"
 If Bomb3.Visible = True And Bomb3.Top - Enemy(k).Top = 360 And Enemy(k).Left = Bomb3.Left Then Enemy(k).Tag = Enemy(k).Tag + "NoDown"
 If Bomb3.Visible = True And Enemy(k).Left - Bomb3.Left = 360 And Enemy(k).Top = Bomb3.Top Then Enemy(k).Tag = Enemy(k).Tag + "NoLeft"
 If Bomb3.Visible = True And Bomb3.Left - Enemy(k).Left = 360 And Enemy(k).Top = Bomb3.Top Then Enemy(k).Tag = Enemy(k).Tag + "NoRight"
 'Deduct a choice(s)
  If InStr(Enemy(k).Tag, "NoRight") > 0 Then Choices = Choices - 1
  If InStr(Enemy(k).Tag, "NoUp") > 0 Then Choices = Choices - 1
  If InStr(Enemy(k).Tag, "NoLeft") > 0 Then Choices = Choices - 1
  If InStr(Enemy(k).Tag, "NoDown") > 0 Then Choices = Choices - 1
 'if the direction the enemy was just going in is still open then keep going that way
 If LastDir(k) = "Right" And InStr(Enemy(k).Tag, "NoRight") = 0 Then
 Enemy(k).Tag = Enemy(k).Tag + "R.ight"
 GoTo DeadSoSkip
 End If
 If LastDir(k) = "Left" And InStr(Enemy(k).Tag, "NoLeft") = 0 Then
 Enemy(k).Tag = Enemy(k).Tag + "L.eft"
 GoTo DeadSoSkip
  End If
 If LastDir(k) = "Up" And InStr(Enemy(k).Tag, "NoUp") = 0 Then
 Enemy(k).Tag = Enemy(k).Tag + "U.p"
 GoTo DeadSoSkip
 End If
 If LastDir(k) = "Down" And InStr(Enemy(k).Tag, "NoDown") = 0 Then
 Enemy(k).Tag = Enemy(k).Tag + "D.own"
 GoTo DeadSoSkip
 End If
 'otherwise assign a new direction
 If Choices = 2 Then '(only 2 bordering)
  If InStr(Enemy(k).Tag, "NoDown") > 0 And InStr(Enemy(k).Tag, "NoRight") > 0 Then
  'in lower right corner
  RandNum1 = Int(Rnd * 2)
  If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "L.eft"
  If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "U.p"
  End If
  If InStr(Enemy(k).Tag, "NoDown") > 0 And InStr(Enemy(k).Tag, "NoLeft") > 0 Then
  'in lower left corner
  RandNum1 = Int(Rnd * 2)
  If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "R.ight"
  If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "U.p"
  End If
  If InStr(Enemy(k).Tag, "NoUp") > 0 And InStr(Enemy(k).Tag, "NoRight") > 0 Then
  'in upper right corner
  RandNum1 = Int(Rnd * 2)
  If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "L.eft"
  If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "D.own"
  End If
  If InStr(Enemy(k).Tag, "NoUp") > 0 And InStr(Enemy(k).Tag, "NoLeft") > 0 Then
  'in upper left corner
  RandNum1 = Int(Rnd * 2)
  If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "D.own"
  If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "R.ight"
  End If
  'if 1 on each side horiz:
  If InStr(Enemy(k).Tag, "NoRight") > 0 And InStr(Enemy(k).Tag, "NoLeft") > 0 Then
  RandNum1 = Int(Rnd * 2)
  If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "U.p"
  If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "D.own"
  End If
  'if 1 on each side vertical:
  If InStr(Enemy(k).Tag, "NoUp") > 0 And InStr(Enemy(k).Tag, "NoDown") > 0 Then
  RandNum1 = Int(Rnd * 2)
  If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "L.eft"
  If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "R.ight"
  End If
 End If
'<< end 2 choices
 If Choices = 1 Then 'cornered by 3 things, get out!
  If InStr(Enemy(k).Tag, "NoUp") > 0 And InStr(Enemy(k).Tag, "NoLeft") > 0 And InStr(Enemy(k).Tag, "NoRight") > 0 Then Enemy(k).Tag = Enemy(k).Tag + "D.own"
  If InStr(Enemy(k).Tag, "NoUp") > 0 And InStr(Enemy(k).Tag, "NoLeft") > 0 And InStr(Enemy(k).Tag, "NoDown") > 0 Then Enemy(k).Tag = Enemy(k).Tag + "R.ight"
  If InStr(Enemy(k).Tag, "NoUp") > 0 And InStr(Enemy(k).Tag, "NoRight") > 0 And InStr(Enemy(k).Tag, "NoDown") > 0 Then Enemy(k).Tag = Enemy(k).Tag + "L.eft"
  If InStr(Enemy(k).Tag, "NoDown") > 0 And InStr(Enemy(k).Tag, "NoLeft") > 0 And InStr(Enemy(k).Tag, "NoRight") > 0 Then Enemy(k).Tag = Enemy(k).Tag + "U.p"
 End If
'<< end 1 choice
 If Choices = 3 Then 'only 1 on one side
  If InStr(Enemy(k).Tag, "NoRight") > 0 Then
  'just can't go right
  RandNum1 = Int(Rnd * 3) '0,1, or 2
   If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "L.eft"
   If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "U.p"
   If RandNum1 = 2 Then Enemy(k).Tag = Enemy(k).Tag + "D.own"
  End If
  If InStr(Enemy(k).Tag, "NoLeft") > 0 Then
  'just can't go left
  RandNum1 = Int(Rnd * 3) '0,1, or 2
   If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "R.ight"
   If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "U.p"
   If RandNum1 = 2 Then Enemy(k).Tag = Enemy(k).Tag + "D.own"
  End If
  If InStr(Enemy(k).Tag, "NoUp") > 0 Then
  'just can't go up
  RandNum1 = Int(Rnd * 3) '0,1, or 2
   If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "L.eft"
   If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "D.own"
   If RandNum1 = 2 Then Enemy(k).Tag = Enemy(k).Tag + "R.ight"
  End If
  If InStr(Enemy(k).Tag, "NoDown") > 0 Then
  'just can't go down
  RandNum1 = Int(Rnd * 3) '0,1, or 2
   If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "L.eft"
   If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "U.p"
   If RandNum1 = 2 Then Enemy(k).Tag = Enemy(k).Tag + "R.ight"
  End If
 End If
 '< end 3 choices
 If Choices = 4 Then
  RandNum1 = Int(Rnd * 4) '0,1,2, or 3
  If RandNum1 = 0 Then Enemy(k).Tag = Enemy(k).Tag + "R.ight"
  If RandNum1 = 1 Then Enemy(k).Tag = Enemy(k).Tag + "L.eft"
  If RandNum1 = 2 Then Enemy(k).Tag = Enemy(k).Tag + "D.own"
  If RandNum1 = 3 Then Enemy(k).Tag = Enemy(k).Tag + "U.p"
 End If
 '< end 4 choices
DeadSoSkip:
 Next k
 EnemyMover.Enabled = False
 tmrEnGo.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
Release = False
'
If KeyCode = 32 Then
    If Direc = "L" Or Direc = "U" Or Direc = "R" Or Direc = "D" Then
        If BombDelay3.Enabled = False And BombDelay2.Enabled = True And BombDelay.Enabled = True And BombCount > 2 Then
            Bomb3.Left = Bman.Left
            Bomb3.Top = Bman.Top
            Bomb3.Visible = True
            BombDelay3.Enabled = True
        End If
        If BombDelay2.Enabled = False And BombDelay.Enabled = True And BombCount > 1 Then
            Bomb2.Left = Bman.Left
            Bomb2.Top = Bman.Top
            Bomb2.Visible = True
            BombDelay2.Enabled = True
        End If
        If BombDelay.Enabled = False Then
            Bomb1.Left = Bman.Left
            Bomb1.Top = Bman.Top
            Bomb1.Visible = True
            BombDelay.Enabled = True
        End If
    Else
        If BombDelay3.Enabled = False And BombDelay2.Enabled = True And BombDelay.Enabled = True And BombCount > 2 Then
            Bomb3.Left = Bman.Left
            Bomb3.Top = Bman.Top
            Bomb3.Visible = True
            BombDelay3.Enabled = True
        End If
        If BombDelay2.Enabled = False And BombDelay.Enabled = True And BombCount > 1 Then
            Bomb2.Left = Bman.Left
            Bomb2.Top = Bman.Top
            Bomb2.Visible = True
            BombDelay2.Enabled = True
        End If
        If BombDelay.Enabled = False Then
            Bomb1.Left = Bman.Left
            Bomb1.Top = Bman.Top
            Bomb1.Visible = True
            BombDelay.Enabled = True
        End If
    End If
End If
'
If Direc = "L" Or Direc = "U" Or Direc = "R" Or Direc = "D" Then Exit Sub
'
If KeyCode = vbKeyRight Then
    If Protected = False Then Bman.Picture = BManRight.Picture
    If Protected = True Then Bman.Picture = PRight.Picture
If Direc <> "R" Then
Direc = "R"
 Do While Release = False
  DoEvents
   If Bman.Left = 5400 Then Exit Sub
   If DelayTimer.Enabled = True Then GoTo sk1
  If Hovering = False Then PreDetectCollision
   If NoRight = True Then Exit Sub
  Bman.Left = Bman.Left + 360
  NoRight = False
  DelayTimer.Enabled = True
sk1:
 Loop
End If
End If
'
If KeyCode = vbKeyLeft Then
    If Protected = False Then Bman.Picture = BManLeft.Picture
    If Protected = True Then Bman.Picture = PLeft.Picture
If Direc <> "L" Then
Direc = "L"
 Do While Release = False
  DoEvents
   If Bman.Left = 360 Then Exit Sub
   If DelayTimer.Enabled = True Then GoTo sk2
  If Hovering = False Then PreDetectCollision
   If NoLeft = True Then Exit Sub
  Bman.Left = Bman.Left - 360
  NoLeft = False
  DelayTimer.Enabled = True
sk2:
 Loop
End If
End If
'
If KeyCode = vbKeyUp Then
    If Protected = False Then Bman.Picture = BManUp.Picture
    If Protected = True Then Bman.Picture = PUp.Picture
If Direc <> "U" Then
Direc = "U"
 Do While Release = False
  DoEvents
   If Bman.Top = 360 Then Exit Sub
   If DelayTimer.Enabled = True Then GoTo sk3
  If Hovering = False Then PreDetectCollision
   If NoUp = True Then Exit Sub
  Bman.Top = Bman.Top - 360
  NoUp = False
  DelayTimer.Enabled = True
sk3:
 Loop
End If
End If
'
If KeyCode = vbKeyDown Then
    If Protected = False Then Bman.Picture = BmanDown.Picture
    If Protected = True Then Bman.Picture = PDown.Picture
If Direc <> "D" Then
Direc = "D"
 Do While Release = False
  DoEvents
   If Bman.Top = 5040 Then Exit Sub
   If DelayTimer.Enabled = True Then GoTo sk4
  If Hovering = False Then PreDetectCollision
   If NoDown = True Then Exit Sub
  Bman.Top = Bman.Top + 360
  NoDown = False
  DelayTimer.Enabled = True
sk4:
 Loop
End If
End If
'
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then Exit Sub
Direc = ""
Release = True
NoRight = False
NoLeft = False
NoUp = False
NoDown = False
End Sub

Private Sub Form_Load()
Randomize
'
Me.Width = 6210
'make sure no direction is prohibited (at first):
NoRight = False
NoLeft = False
NoUp = False
NoDown = False
Release = False
Slowed = False
Protected = False
JustLostGame = False
'//////if u add levels change this number!:
MaxLevels = 10
'
Choices = 4 'to do w/ the enemies movements
NumberOfWeakBlocks = 47 'remember, the array begins at 0 so always use one less than the ACTUAL number of weak blocks!!!!!
NumberOfEnemies = 7 'see weakblocks
NumberOfRocks = 50 'see weakblocks and enemies
StartLeft = 360: StartTop = 360 'bman starts at default 360, 360
NumFlames = 8 'at first there are 8 flame puffs (this can grow!)
EFCount = 0
Chex = 0    '\
Chev = 0    ' - used in various delay timers
Ched = 0    '/
Chey = 0
BombCount = 1 'at first u have 1 bomb max onscreen
BiggerBomb.Top = -1000 'stores it in a "safe place" offscreen so it can be retrieved but only if you CAN get more bomb expansions!
CurrLevel = 0
PlayerScore = 0
CurrLives = 5
EnemyCount = NumberOfEnemies + 1
'load embedded default pics:
Bman.Picture = BmanDown.Picture 'the default for Bman
'
GotHighScore = False
'
Bman.Left = StartLeft
Bman.Top = StartTop
'
For i = 0 To 8
Flame(i).ZOrder 'makes sure flames "consume all" (visibly)
Next i
For z = 0 To NumberOfEnemies
Enemy(z).ZOrder 'makes sure enemies "consume all" (visibly)
Next z
Bman.ZOrder
'
For j = 0 To NumberOfEnemies
Enemy(j).Picture = En1.Picture
Next j
'so flames won't appear to jump over the boundary stones:
Image1.ZOrder: Image2.ZOrder: Image3.ZOrder: Image4.ZOrder
Bomb1.ZOrder: Bomb2.ZOrder: Bomb3.ZOrder
'set weaks w/zorder so the stairs are hidden:
For Y = 0 To NumberOfWeakBlocks
Weak(Y).ZOrder
Next Y
Frame1.ZOrder
'load the first level presets:
NextLevel
'so i can see the stop button in vb on my pc:
Top = (Screen.Height - Height) \ 2
Left = (Screen.Width - Width) \ 2
frmMain.Top = frmMain.Top + 100
'start midi in background (if possible)
'On Error Resume Next
MediaPlayer1.FileName = App.Path + "\Sounds\Music.mid"
MediaPlayer1.Play
'
End Sub

Sub PreDetectCollision()
'see if bman would hit a bomb, rock, or weakrock and if so prevent that collision
If Bomb1.Visible = True And Bman.Left - Bomb1.Left = 360 And Bman.Top = Bomb1.Top Then NoLeft = True
If Bomb1.Visible = True And Bomb1.Left - Bman.Left = 360 And Bman.Top = Bomb1.Top Then NoRight = True
If Bomb1.Visible = True And Bman.Top - Bomb1.Top = 360 And Bman.Left = Bomb1.Left Then NoUp = True
If Bomb1.Visible = True And Bomb1.Top - Bman.Top = 360 And Bman.Left = Bomb1.Left Then NoDown = True
If Bomb2.Visible = True And Bman.Left - Bomb2.Left = 360 And Bman.Top = Bomb2.Top Then NoLeft = True
If Bomb2.Visible = True And Bomb2.Left - Bman.Left = 360 And Bman.Top = Bomb2.Top Then NoRight = True
If Bomb2.Visible = True And Bman.Top - Bomb2.Top = 360 And Bman.Left = Bomb2.Left Then NoUp = True
If Bomb2.Visible = True And Bomb2.Top - Bman.Top = 360 And Bman.Left = Bomb2.Left Then NoDown = True
If Bomb3.Visible = True And Bman.Left - Bomb3.Left = 360 And Bman.Top = Bomb3.Top Then NoLeft = True
If Bomb3.Visible = True And Bomb3.Left - Bman.Left = 360 And Bman.Top = Bomb3.Top Then NoRight = True
If Bomb3.Visible = True And Bman.Top - Bomb3.Top = 360 And Bman.Left = Bomb3.Left Then NoUp = True
If Bomb3.Visible = True And Bomb3.Top - Bman.Top = 360 And Bman.Left = Bomb3.Left Then NoDown = True
For i = 0 To NumberOfRocks
If Rock(i).Top - Bman.Top = 360 And Bman.Left = Rock(i).Left Then NoDown = True
If Bman.Top - Rock(i).Top = 360 And Bman.Left = Rock(i).Left Then NoUp = True
If Rock(i).Left - Bman.Left = 360 And Bman.Top = Rock(i).Top Then NoRight = True
If Bman.Left - Rock(i).Left = 360 And Bman.Top = Rock(i).Top Then NoLeft = True
Next i
For j = 0 To NumberOfWeakBlocks
If Weak(j).Top - Bman.Top = 360 And Bman.Left = Weak(j).Left Then NoDown = True
If Bman.Top - Weak(j).Top = 360 And Bman.Left = Weak(j).Left Then NoUp = True
If Weak(j).Left - Bman.Left = 360 And Bman.Top = Weak(j).Top Then NoRight = True
If Bman.Left - Weak(j).Left = 360 And Bman.Top = Weak(j).Top Then NoLeft = True
Next j
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbBlack
Label4.ForeColor = vbBlack
End Sub

Private Sub FreqScan_Timer()
On Error Resume Next
'this timer frequently scans for many misc. things and handles them
'if midi is done, restart it:
If MediaPlayer1.PlayState = mpStopped Then MediaPlayer1.Play
'if all lives are lost, restart game if user wants
If CurrLives < 0 Then
'
PlayerScore = PlayerScore + 4 'prevents glitches in the highscores table as much as possible.  i know the highscores table is kind of primitive as are the methods i use, but it works pretty well and i had to go thru TONS of trial and error to get it that way!
'
Hide
'
CheckIfHigh 6, PlayerScore, frmHighs
If GotHighScore = True Then
    frmHighs.Show
    JustLostGame = True
    FreqScan.Enabled = False
    Exit Sub
End If
'
Show
If JustLostGame = True Then Exit Sub
    JustLostGame = True
    response = MsgBox("You lost all your lives!  Play again?", vbYesNo + vbInformation, "Ouch...")
        If response = vbYes Then
            frmMain.Bman.Picture = frmMain.BmanDown.Picture 'default
            frmMain.BManRemains.Visible = False 'clean up his bodily remains
            Hovering = False
            Slowed = False
            frmMain.DelayTimer.interval = 120
            NumFlames = 8
            frmMain.tmrStartingOff.Enabled = True
            frmMain.EnemyMover.interval = 600
            frmMain.Bman.Top = 360
            frmMain.Bman.Left = 360
            CurrLives = 5
            BombCount = 1
            PlayerScore = 0
            JustLostGame = True
            LoadLevel 1
            Me.Show
            For tr = 0 To NumberOfEnemies
                Enemy(tr).Tag = ""
            Next tr
        Else
            lblExit_Click
        End If
End If
'if game beaten:
If CurrLevel = MaxLevels + 1 Then GameWin
'update score/lives left/level/and enemy label:
lblLevel.Caption = CurrLevel
lblScore.Caption = PlayerScore
lblLives = "X " & CurrLives
lblEnsLeft = "X " & EnemyCount
'check if bman goes in stairs:
If Stairs.Visible = True And Bman.Top = Stairs.Top And Bman.Left = Stairs.Left Then
    PlayerScore = PlayerScore + 250 'level bonus
    For po = 0 To NumberOfEnemies
        frmMain.Enemy(po).Tag = "" 'clear tags so the enemies move again
    Next po
    NextLevel
End If
'check for enemy hitting bman
If Protected = False And StartingOffClean = False Then
For i = 0 To NumberOfEnemies
If Enemy(i).Tag <> "Dead" And Bman.Top = Enemy(i).Top And Bman.Left = Enemy(i).Left Then
If tmrStartingOff.Enabled = True Then Exit Sub
Bman.Visible = False
BManRemains.Left = Bman.Left: BManRemains.Top = Bman.Top
BManRemains.Visible = True
BManDie.DoVerb
Pause 2
Bman.Left = StartLeft 'put him at start point
Bman.Top = StartTop '''
Bman.Visible = True
CurrLives = CurrLives - 1
Bman.Picture = BmanDown.Picture 'default
BManRemains.Visible = False 'clean up his bodily remains
Hovering = False
BombCount = 1
Slowed = False
DelayTimer.interval = 120
NumFlames = 8
tmrStartingOff.Enabled = True
End If
Next i
End If
'check for bman getting burned
If Protected = False And StartingOffClean = False Then
For j = 0 To NumFlames
If Flame(j).Visible = True And Bman.Top = Flame(j).Top And Bman.Left = Flame(j).Left Then
Bman.Visible = False
'BManRemains.Left = Bman.Left: BManRemains.Top = Bman.Top
'BManRemains.Visible = True
BManDie.DoVerb
    For X = 0 To 8
    If StartingOffClean = True Then Exit For
    If StartTop = Flame(X).Top And StartLeft = Flame(X).Left Then Pause 2 'if and only if he would start in a flame, then pause for 2 and prevent a double life loss.
    Next X
Bman.Left = StartLeft 'put him at start point
Bman.Top = StartTop '''
Bman.Visible = True
CurrLives = CurrLives - 1
'BManRemains.Visible = False
Hovering = False
BombCount = 1
Slowed = False
DelayTimer.interval = 120
NumFlames = 8
tmrStartingOff.Enabled = True
End If
Next j
End If
'if an enemy moves into the line of bomb fire then the bomb engine above will not detect it so do it here!:
For l = 0 To 20
For k = 0 To NumberOfEnemies
If Flame(l).Visible = True And Enemy(k).Top = Flame(l).Top And Enemy(k).Left = Flame(l).Left Then
Enemy(k).Top = -1000
Enemy(k).Tag = "Dead"
EnemyDie.DoVerb
PlayerScore = PlayerScore + 100
EnemyCount = EnemyCount - 1
If EnemyCount = 0 Then Stairs.Visible = True
End If
Next k
Next l
'if powerup gotten:
'speed up:
If SpeedUp.Visible = True And Bman.Top = SpeedUp.Top And Bman.Left = SpeedUp.Left Then
 SpeedUp.Visible = False
 PlayWav App.Path + "\Sounds\PickUp.wav"
 DoBonusSeq "Speed Up!", vbRed, SpeedUp.Picture, "Speed Up"
 DelayTimer.interval = 66 'cuts moving delay in half!
 SpeedUp.Top = -1000
End If
'Enemy freezer:
If EnemyFreeze.Visible = True And Bman.Top = EnemyFreeze.Top And Bman.Left = EnemyFreeze.Left Then
 EnemyFreeze.Visible = False
 PlayWav App.Path + "\Sounds\PickUp.wav"
 DoBonusSeq "Enemy Freeze!", vbBlue, EnemyFreeze.Picture, "Enemy Freeze"
 EnemyFreeze.Top = -1000
 EnemyMover.Enabled = False
 EightTimer.Enabled = True
 EnemyMover.Enabled = False
End If
'Powered Up bomb:
If BiggerBomb.Visible = True And BiggerBomb.Left = Bman.Left And BiggerBomb.Top = Bman.Top Then
 EFCount = EFCount + 1 ' if this number reaches 4, no more bomb expansions for you!!
 If EFCount = 4 Then BiggerBomb.Top = -2000 'this is kind of an inconvienient way to do this, but oh well.  when you've gotten your bombs bigger twice already then u can't anymore.
 BiggerBomb.Visible = False
 PlayWav App.Path + "\Sounds\PickUp.wav"
 DoBonusSeq "Bigger Bombs!", vbRed, BiggerBomb.Picture, "Bigger Explosions!"
 BiggerBomb.Top = -1000 'may reuse
 NumFlames = NumFlames + 4 'makes the explosion 1 longer in each direction! Cool!
End If
'Death pills eaten (ohno!)
If Protected = False Then
If DeathPills.Visible = True And Bman.Left = DeathPills.Left And Bman.Top = DeathPills.Top Then
 DeathPills.Visible = False
 PlayWav App.Path + "\Sounds\bmandie.wav"
 DoBonusSeq "Death Pills!", vbRed, DeathPills.Picture, "Suicide Pills"
 DeathPills.Top = -1000
 Bman.Top = StartTop
 Bman.Left = StartLeft
 Bman.Picture = BmanDown.Picture
 Hovering = False
BombCount = 1
Slowed = False
DelayTimer.interval = 120
NumFlames = 8
CurrLives = CurrLives - 1
tmrStartingOff.Enabled = True
End If
End If
'Slow Down pick up:
If Slow.Visible = True And Bman.Left = Slow.Left And Bman.Top = Slow.Top Then
 Slow.Visible = False
 PlayWav App.Path + "\Sounds\PickUp.wav"
 DoBonusSeq "Sluggish!", vbBlack, Slow.Picture, "Sluggish Bomberman"
 Slow.Top = -1000
 DelayTimer.interval = 1500 'make bman a slug.
 Slowed = True 'flag so u cant get speed up powerup while slowed
 Slug.Enabled = True
End If
'Hover Boots PickUp (bman can (until end of level) walk over all bricks
If Hover.Visible = True And Bman.Left = Hover.Left And Bman.Top = Hover.Top Then
 Hover.Visible = False
 PlayWav App.Path + "\Sounds\PickUp.wav"
 DoBonusSeq "Hover Boots!", vbYellow, Hover.Picture, "Hover Boots!"
 Hover.Top = -1000
 Hovering = True
 'so u can see bombs on bricks if u have hoverboots:
 Bomb1.ZOrder
 Bomb2.ZOrder
 Bomb3.ZOrder
End If
'Extra Life Gotten
If Heart.Visible = True And Bman.Left = Heart.Left And Bman.Top = Heart.Top Then
 Heart.Visible = False
 PlayWav App.Path + "\Sounds\PickUp.wav"
 DoBonusSeq "Extra Life!", vbRed, Heart.Picture, "Extra Life!"
 Heart.Top = -1000
 CurrLives = CurrLives + 1
End If
'extra bomb
If ExtraBomb.Visible = True And ExtraBomb.Left = Bman.Left And ExtraBomb.Top = Bman.Top Then
 ExtraBomb.Visible = False
 PlayWav App.Path + "\Sounds\PickUp.wav"
 DoBonusSeq "Extra Bomb!", vbBlue, ExtraBomb.Picture, "Extra Bomb"
 ExtraBomb.Top = -1000
 BombCount = BombCount + 1
 If BombCount = 3 Then MaxedOut4Bombs = True
End If
'temporary invincibilty
If TI.Visible = True And TI.Top = Bman.Top And TI.Left = Bman.Left Then
 TI.Visible = False
 TI.Top = -1000
 PlayWav App.Path + "\Sounds\PickUp.wav"
 DoBonusSeq "Invincibility!", vbYellow, TI.Picture, "Invincibility!"
 Protected = True
 Seven.Enabled = True
  If Bman.Picture = BManUp.Picture Then Bman.Picture = PUp.Picture
  If Bman.Picture = BManRight.Picture Then Bman.Picture = PRight.Picture
  If Bman.Picture = BManLeft.Picture Then Bman.Picture = PLeft.Picture
  If Bman.Picture = BmanDown.Picture Then Bman.Picture = PDown.Picture
End If
End Sub

Sub DoBonusSeq(WhatToSay As String, ForeColor As String, PicToLoad As Picture, WhatPickUp As String)
'when a bonus is picked up it flashes the words.  just looks cool.
PicDisp.Picture = PicToLoad
lblPU.Caption = WhatPickUp
lblBonus.ZOrder
lblBonus.Left = 0
lblBonus.Top = 2040
lblBonus.Width = 6135
lblBonus.Height = 1335
lblBonus.Font.Size = 48
lblBonus.Caption = WhatToSay
lblBonus.ForeColor = ForeColor
lblBonus.Visible = True
Pause 0.6
lblBonus.Visible = False
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbBlack
Label4.ForeColor = vbBlack
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbBlack
Label4.ForeColor = vbBlack
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbBlack
Label4.ForeColor = vbBlack
End Sub

Private Sub Label4_Click()
PopupMenu mnufile, , 5520, 6000
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlack
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlue
lblExit.ForeColor = vbBlack
End Sub

Private Sub lblExit_Click()
response = MsgBox("Are you sure you want to quit VBBomberman?", vbYesNo + vbQuestion, "Quit?")
If response = vbYes Then
Unload Me
End
Else
End If
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbBlack
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.ForeColor = vbBlue
Label4.ForeColor = vbBlack
End Sub

Private Sub mnuAbout_Click()
'EnemyMover.Enabled = False 'just so that when ur not lookin no baddy tries to get u!  (O:
Shell "NotePad " & App.Path & "\README.txt", vbNormalFocus
End Sub

Private Sub mnuHTP_Click()
frmHelp.Show
End Sub

Private Sub PicDisp_Click()
'frmGetData.Show
End Sub


Private Sub Seven_Timer()
Chev = Chev + 1
If Chev = 2 Then
Protected = False
If Bman.Picture = PUp.Picture Then Bman.Picture = BManUp.Picture
If Bman.Picture = PDown.Picture Then Bman.Picture = BmanDown.Picture
If Bman.Picture = PRight.Picture Then Bman.Picture = BManRight.Picture
If Bman.Picture = PLeft.Picture Then Bman.Picture = BManLeft.Picture
Chev = 0
Seven.Enabled = False
End If
End Sub

Private Sub Slug_Timer()
Chex = Chex + 1
If Chex = 2 Then
Slug.Enabled = False
 Slowed = False
 Chex = 0
 DelayTimer.interval = 120
End If
End Sub

Private Sub tmrEnGo_Timer()
For i = 0 To NumberOfEnemies
If InStr(Enemy(i).Tag, "R.ight") > 0 Then 'as strange as it may seem, yes, there is supposed to be a period after the first character in each of the directions.  this helps the instr function properly!
Enemy(i).Left = Enemy(i).Left + 360
LastDir(i) = "Right"
End If
If InStr(Enemy(i).Tag, "L.eft") > 0 Then
Enemy(i).Left = Enemy(i).Left - 360
LastDir(i) = "Left"
End If
If InStr(Enemy(i).Tag, "U.p") > 0 Then
Enemy(i).Top = Enemy(i).Top - 360
LastDir(i) = "Up"
End If
If InStr(Enemy(i).Tag, "D.own") > 0 Then
Enemy(i).Top = Enemy(i).Top + 360
LastDir(i) = "Down"
End If
Next i
'when all the movements are finished, :
tmrEnGo.Enabled = False
EnemyMover.Enabled = True
End Sub

Sub NextLevel()
CurrLevel = CurrLevel + 1
frmMain.Stairs.Visible = False
LoadLevel CurrLevel
End Sub

Private Sub tmrStartingOff_Timer()
StartingOffClean = True 'this makes sure that when bman starts off, no enemies cheaply kill him.
Chey = Chey + 1
If Bman.Visible = True Then Bman.Visible = False Else Bman.Visible = True
If Chey = 8 Then
tmrStartingOff.Enabled = False
StartingOffClean = False
Chey = 0
End If
End Sub

Sub GameWin()
'if you get past the last level (very hard, btw) you start at easy level and get to really advance your score!  congrats!
CurrLevel = 0: NextLevel
End Sub

