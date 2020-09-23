Attribute VB_Name = "Highs"
Public GotHighScore As Boolean
Public JustLostGame As Boolean

Public Function LoadScores(ZeForm As Form)
On Error Resume Next
z# = -1
Open App.Path + "\Highscores.dat" For Input As #1
Do Until EOF(1)
    Input #1, NowStr$
    For i = 1 To Len(NowStr$)
        X$ = Mid$(NowStr$, i, 1)
        If X$ = "ë" Then Named$ = Mid$(NowStr$, 1, i - 1)
    Next i
    z# = z# + 1
    ZeForm.Name1(z#).Caption = Named$
Loop
Close #1
Open App.Path + "\Highscores.dat" For Input As #1
    z# = -1
Do Until EOF(1)
    Input #1, NowStr$
    For j = 1 To Len(NowStr$)
        X$ = Mid$(NowStr$, j, 1)
        If X$ = "ë" Then Scr$ = Mid$(NowStr$, j + 1, Len(NowStr$) - j)
    Next j
    z# = z# + 1
    ZeForm.Score1(z#).Caption = Scr$
Loop
Close #1
End Function

Public Function CheckIfHigh(NumberOfHighs As Integer, CurrentScore As Integer, TheFormX As Form)
'see if the score is greater than any of the highs:
If JustLostGame = True Then Exit Function
For k = 0 To NumberOfHighs - 1
    If CurrentScore > TheFormX.Score1(k).Caption Then
    GotHighScore = True
    TheFormX.Score1(k).Tag = "X"
Retry:
    pname$ = InputBox("Congratulations!  You got a high score!  Enter your name! (1-12 characters)", "High Score!")
    If InStr(pname$, "ë") > 0 Or Len(pname$) > 12 Or pname$ = "" Then MsgBox "Please enter a valid name, 1 to 12 characters in length!", vbOKOnly + vbExclamation, "Error:": GoTo Retry
    GoTo ks
    End If
Next k
ks:
'if so, then change the highscores table
For X = 0 To 5
    If TheFormX.Score1(X).Tag = "X" Then
        'move the beaten score down 1
        TheFormX.Score1(X + 1).Caption = TheFormX.Score1(X).Caption
        TheFormX.Name1(X + 1).Caption = TheFormX.Name1(X).Caption
        '
        TheFormX.Score1(X).Caption = CurrentScore
        TheFormX.Name1(X).Caption = pname$
    End If
Next X
'now send the updated highs back to the .dat file:
For z = 0 To 5
        TheFormX.Text1.Text = TheFormX.Text1.Text + TheFormX.Name1(z).Caption + "ë"
        TheFormX.Text1.Text = TheFormX.Text1.Text + TheFormX.Score1(z).Caption + vbCrLf
Next z
Open App.Path + "\Highscores.dat" For Output As #1
    Print #1, TheFormX.Text1.Text
Close #1
'clear tags:
For q = 0 To 5
    TheFormX.Score1(q).Tag = ""
Next q
End Function

