Option Compare Database
Option Explicit

'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    Dim objAD As Object
    Dim objUser As Object
    Dim userName As String

    Set objAD = CreateObject("AdSystemInfo")
    Set objUser = GetObject("LDAP://" & objAD.userName)

    userName = objUser.displayName
    Me.lblCurrentUser.Caption = "Logged in as: " & userName

    DoCmd.SetWarnings False
    DoCmd.ShowToolbar "Ribbon", acToolbarNo


Form_Load_Exit:
    Exit Sub

Form_Load_Err:
    MsgBox Error$
    Resume Form_Load_Exit

End Sub

'------------------------------------------------------------
' btnStartQuiz_Click
'
'------------------------------------------------------------
Private Sub btnStartQuiz_Click()
On Error GoTo btnStartQuiz_Click_Err

    Dim vbOkay As VbMsgBoxStyle
    Dim clearAnswerSheet As String

    clearAnswerSheet = "DELETE * FROM tblAnswerSheet"

                If MsgBox("You are about to begin your quiz. Good luck!", vbOKCancel + vbQuestion) = vbCancel Then
        Exit Sub
    Else:
        TempVars.RemoveAll
        DoCmd.RunSQL (clearAnswerSheet)
        DoCmd.OpenForm "frmRandomizedQuiz"
        DoCmd.Close acForm, Me.Name
    End If


btnStartQuiz_Click_Exit:
    Exit Sub
btnStartQuiz_Click_Err:
    MsgBox Error$
    Resume btnStartQuiz_Click_Exit

End Sub

'------------------------------------------------------------
' btnExitQuiz_Click
'
'------------------------------------------------------------
Private Sub btnExitQuiz_Click()
On Error GoTo btnExitQuiz_Click_Err

    DoCmd.Quit acQuitPrompt


btnExitQuiz_Click_Exit:
    Exit Sub

btnExitQuiz_Click_Err:
    MsgBox Error$
    Resume btnExitQuiz_Click_Exit

End Sub
