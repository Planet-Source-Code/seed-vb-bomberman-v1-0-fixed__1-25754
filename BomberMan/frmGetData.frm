VERSION 5.00
Begin VB.Form frmGetData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Form"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Get Data"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'hello there.
'this form retrieves all the data of frmmains objects positions
'i used this to create new levels and load them with ease.
'if u want to add/edit levels, then simply rearrange the stuff
'on frmmain to your liking and then get this form to load
'somehow.  click the button and it will get all the data.  (O:
'paste that data into Levels.bas under either a new level
'or an existing one.  (but u must also change the MaxLevel
'in the form_load sub to the new high or whatever.)
'i think this method is quite convenient! (oh yeah, be sure that
'the enemy timer (EnemyMover) is disabled when making/editing
'levels otherwise u'll get invalid or unwanted data.
For i = 0 To 50
Text1.Text = Text1.Text & "frmmain.Rock(" & i & ").top = " & frmMain.Rock(i).Top & vbCrLf
Text1.Text = Text1.Text & "frmmain.Rock(" & i & ").left = " & frmMain.Rock(i).Left & vbCrLf
Next i
For j = 0 To 47
Text1.Text = Text1.Text & "frmmain.Weak(" & j & ").top = " & frmMain.Weak(j).Top & vbCrLf
Text1.Text = Text1.Text & "frmmain.Weak(" & j & ").left = " & frmMain.Weak(j).Left & vbCrLf
Next j
For k = 0 To 7
Text1.Text = Text1.Text & "frmmain.Enemy(" & k & ").top = " & frmMain.Enemy(k).Top & vbCrLf
Text1.Text = Text1.Text & "frmmain.Enemy(" & k & ").left = " & frmMain.Enemy(k).Left & vbCrLf
Next k
Text1.Text = Text1.Text & "frmmain.stairs.left = " & frmMain.Stairs.Left & vbCrLf
Text1.Text = Text1.Text & "frmmain.stairs.top = " & frmMain.Stairs.Top & vbCrLf
End Sub
