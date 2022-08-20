Option Compare Database
Option Explicit
Public questionNo As Integer
Public rsQuestionBank As Recordset
Public rsAnswerSheet As Recordset2
Public vbOkay As VbMsgBoxStyle

'------------------------------------------------------------
' Form_Load
'
'------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo Form_Load_Err

    Set rsAnswerSheet = CurrentDb.OpenRecordset("tblAnswerSheet", dbOpenSnapshot, dbReadOnly)
    Set rsQuestionBank = CurrentDb.OpenRecordset("tblQuestionBank", dbOpenSnapshot, dbReadOnly)

    questionNo = 1

    Me.lblQuestionNo.Caption = questionNo & " of " & DCount("*", "qryRandomizedQuiz")

    Me.radOptionA.Value = False
    Me.radOptionB.Value = False
    Me.radOptionC.Value = False
    Me.radOptionD.Value = False

    If Me.QuestionType = "Multiple Choice" Then
        Me.radOptionA.Visible = True
        Me.radOptionB.Visible = True
        Me.radOptionC.Visible = True
        Me.radOptionD.Visible = True
    ElseIf Me.QuestionType = "True/False" Then
        Me.radOptionA.Visible = True
        Me.radOptionB.Visible = True
        Me.radOptionC.Visible = False
        Me.radOptionD.Visible = False
    End If


Form_Load_Exit:
    Exit Sub

Form_Load_Err:
    MsgBox Error$
    Resume Form_Load_Exit

End Sub

'------------------------------------------------------------
' btnExitQuiz_Click
'
'------------------------------------------------------------
Private Sub btnExitQuiz_Click()
On Error GoTo btnExitQuiz_Click_Err

    If MsgBox("Are you sure you want to exit your quiz?", vbYesNo + vbQuestion) = vbYes Then
        Set rsAnswerSheet = Nothing
        Set rsQuestionBank = Nothing
        TempVars.RemoveAll
        DoCmd.Quit acPrompt
    Else:
        Exit Sub
    End If


btnExitQuiz_Click_Exit:
    Exit Sub

btnExitQuiz_Click_Err:
    MsgBox Error$
    Resume btnExitQuiz_Click_Exit

End Sub

'------------------------------------------------------------
' radOptionA_Click
'
'------------------------------------------------------------
Private Sub radOptionA_Click()
On Error GoTo radOptionA_Click_Err

    If Me.QuestionType = "Multiple Choice" Then
        TempVars("SelectedOption") = "A"
        TempVars("FinalAnswer") = "A - " & Me.OptionA
    ElseIf Me.QuestionType = "True/False" Then
        TempVars("SelectedOption") = "True"
        TempVars("FinalAnswer") = "True"
    End If

    Me.radOptionA.Value = True
    Me.radOptionB.Value = False
    Me.radOptionC.Value = False
    Me.radOptionD.Value = False


radOptionA_Click_Exit:
    Exit Sub

radOptionA_Click_Err:
    MsgBox Error$
    Resume radOptionA_Click_Exit

End Sub

'------------------------------------------------------------
' radOptionB_Click
'
'------------------------------------------------------------
Private Sub radOptionB_Click()
On Error GoTo radOptionB_Click_Err

    If Me.QuestionType = "Multiple Choice" Then
        TempVars("SelectedOption") = "B"
        TempVars("FinalAnswer") = "B - " & Me.OptionB
    ElseIf Me.QuestionType = "True/False" Then
        TempVars("SelectedOption") = "False"
        TempVars("FinalAnswer") = "False"
    End If

    Me.radOptionA.Value = False
    Me.radOptionB.Value = True
    Me.radOptionC.Value = False
    Me.radOptionD.Value = False


radOptionB_Click_Exit:
    Exit Sub

radOptionB_Click_Err:
    MsgBox Error$
    Resume radOptionB_Click_Exit

End Sub

'------------------------------------------------------------
' radOptionC_Click
'
'------------------------------------------------------------
Private Sub radOptionC_Click()
On Error GoTo radOptionC_Click_Err

    TempVars("SelectedOption") = "C"
    TempVars("FinalAnswer") = "C - " & Me.OptionC

    Me.radOptionA.Value = False
    Me.radOptionB.Value = False
    Me.radOptionC.Value = True
    Me.radOptionD.Value = False


radOptionC_Click_Exit:
    Exit Sub

radOptionC_Click_Err:
    MsgBox Error$
    Resume radOptionC_Click_Exit

End Sub

'------------------------------------------------------------
' radOptionD_Click
'
'------------------------------------------------------------
Private Sub radOptionD_Click()
On Error GoTo radOptionD_Click_Err

    TempVars("SelectedOption") = "D"
    TempVars("FinalAnswer") = "D - " & Me.OptionD

    Me.radOptionA.Value = False
    Me.radOptionB.Value = False
    Me.radOptionC.Value = False
    Me.radOptionD.Value = True


radOptionD_Click_Exit:
    Exit Sub

radOptionD_Click_Err:
    MsgBox Error$
    Resume radOptionD_Click_Exit

End Sub

'------------------------------------------------------------
' btnPreviousQuestion_Click
'
'------------------------------------------------------------
Private Sub btnPreviousQuestion_Click()
On Error GoTo btnPreviousQuestion_Click_Err

    If questionNo = 1 Then
        Exit Sub
    Else:
        questionNo = questionNo - 1
        Me.lblQuestionNo.Caption = questionNo & " of " & DCount("*", "qryRandomizedQuiz")
    End If

    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

    If Me.QuestionType = "Multiple Choice" Then
        Me.radOptionA.Visible = True
        Me.radOptionB.Visible = True
        Me.radOptionC.Visible = True
        Me.radOptionD.Visible = True
    ElseIf Me.QuestionType = "True/False" Then
        Me.radOptionA.Visible = True
        Me.radOptionB.Visible = True
        Me.radOptionC.Visible = False
        Me.radOptionD.Visible = False
    End If

    rsAnswerSheet.FindFirst "QuestionText='" & Me.txtQuestionText & "'"

    If rsAnswerSheet!SelectedOption = "A" Then
        TempVars("SelectedOption") = "A"
        TempVars("FinalAnswer") = "A - " & Me.OptionA
        Me.radOptionA.Value = True
        Me.radOptionB.Value = False
        Me.radOptionC.Value = False
        Me.radOptionD.Value = False
    ElseIf rsAnswerSheet!SelectedOption = "B" Then
        TempVars("SelectedOption") = "B"
        TempVars("FinalAnswer") = "B - " & Me.OptionB
        Me.radOptionA.Value = False
        Me.radOptionB.Value = True
        Me.radOptionC.Value = False
        Me.radOptionD.Value = False
    ElseIf rsAnswerSheet!SelectedOption = "C" Then
        TempVars("SelectedOption") = "C"
        TempVars("FinalAnswer") = "C - " & Me.OptionC
        Me.radOptionA.Value = False
        Me.radOptionB.Value = False
        Me.radOptionC.Value = True
        Me.radOptionD.Value = False
    ElseIf rsAnswerSheet!SelectedOption = "D" Then
        TempVars("SelectedOption") = "D"
        TempVars("FinalAnswer") = "D - " & Me.OptionD
        Me.radOptionA.Value = False
        Me.radOptionB.Value = False
        Me.radOptionC.Value = False
        Me.radOptionD.Value = True
    ElseIf rsAnswerSheet!SelectedOption = "True" Then
        TempVars("SelectedOption") = "True"
        TempVars("FinalAnswer") = "True"
        Me.radOptionA.Value = True
        Me.radOptionB.Value = False
    ElseIf rsAnswerSheet!SelectedOption = "False" Then
        TempVars("SelectedOption") = "False"
        TempVars("FinalAnswer") = "False"
        Me.radOptionA.Value = False
        Me.radOptionB.Value = True
    End If


btnPreviousQuestion_Click_Exit:
    Exit Sub

btnPreviousQuestion_Click_Err:
    MsgBox Error$
    Resume btnPreviousQuestion_Click_Exit

End Sub

'------------------------------------------------------------
' btnNextQuestion_Click
'
'------------------------------------------------------------
Private Sub btnNextQuestion_Click()
On Error GoTo btnNextQuestion_Click_Err

    Dim updateSelectedOption As String
    Dim updateFinalAnswer As String
    Dim submitFinalAnswer As String
    Dim gradeQuiz As String

    updateSelectedOption = "UPDATE tblAnswerSheet SET SelectedOption = '" & TempVars("SelectedOption") & "' WHERE QuestionText='" & Me.txtQuestionText & "'"
    updateFinalAnswer = "UPDATE tblAnswerSheet SET FinalAnswer = '" & TempVars("FinalAnswer") & "' WHERE QuestionText='" & Me.txtQuestionText & "'"
    gradeQuiz = "UPDATE tblAnswerSheet SET IsCorrect = True WHERE FinalAnswer = CorrectAnswer"

    If IsNull(TempVars("SelectedOption")) Then
        MsgBox ("Uh-oh. It's blank!") _
        & vbCrLf & "" _
        & vbCrLf & "Please make a selection before moving to the next question.", vbOkay
        Exit Sub
    End If

    rsAnswerSheet.FindFirst "QuestionText='" & Me.txtQuestionText & "'"

    If rsAnswerSheet.NoMatch = False Then
        DoCmd.RunSQL (updateSelectedOption)
        DoCmd.RunSQL (updateFinalAnswer)
    Else:
        rsQuestionBank.FindFirst "QuestionText='" & Me.txtQuestionText & "'"
        If rsQuestionBank!CorrectOption = "A" Then
            TempVars("CorrectAnswer") = "A - " & Me.OptionA
        ElseIf rsQuestionBank!CorrectOption = "B" Then
            TempVars("CorrectAnswer") = "B - " & Me.OptionB
        ElseIf rsQuestionBank!CorrectOption = "C" Then
            TempVars("CorrectAnswer") = "C - " & Me.OptionC
        ElseIf rsQuestionBank!CorrectOption = "D" Then
            TempVars("CorrectAnswer") = "D - " & Me.OptionD
        ElseIf rsQuestionBank!CorrectOption = "True" Then
            TempVars("CorrectAnswer") = "True"
        ElseIf rsQuestionBank!CorrectOption = "False" Then
            TempVars("CorrectAnswer") = "False"
        End If
        submitFinalAnswer = "INSERT INTO tblAnswerSheet (QuestionNo, QuestionID, QuestionText, SelectedOption, FinalAnswer, CorrectAnswer) VALUES" & _
            "('" & questionNo & "', '" & Me.ID & "', '" & Me.QuestionText & "','" & TempVars("SelectedOption") & "', '" & TempVars("FinalAnswer") & "', '" & TempVars("CorrectAnswer") & "')"
        DoCmd.RunSQL (submitFinalAnswer)
    End If

nextQuestion:

    If questionNo = DCount("*", "tblQuestionBank") Then
        If MsgBox("Are you sure you want to submit your quiz?", vbYesNo + vbQuestion) = vbYes Then
            DoCmd.RunSQL (gradeQuiz)
            DoCmd.OpenReport "rptReportCard, acViewReport, "", "", acNormal"
            DoCmd.Close acForm, Me.Name
        Else:
            Exit Sub
        End If
        Exit Sub
    Else:
        questionNo = questionNo + 1
        Me.lblQuestionNo.Caption = questionNo & " of " & DCount("*", "qryRandomizedQuiz")
    End If

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

    Me.radOptionA.Value = False
    Me.radOptionB.Value = False
    Me.radOptionC.Value = False
    Me.radOptionD.Value = False

    If Me.QuestionType = "Multiple Choice" Then
        Me.radOptionA.Visible = True
        Me.radOptionB.Visible = True
        Me.radOptionC.Visible = True
        Me.radOptionD.Visible = True
    ElseIf Me.QuestionType = "True/False" Then
        Me.radOptionA.Visible = True
        Me.radOptionB.Visible = True
        Me.radOptionC.Visible = False
        Me.radOptionD.Visible = False
    End If

    Set rsAnswerSheet = CurrentDb.OpenRecordset("tblAnswerSheet", dbOpenSnapshot, dbReadOnly)

    rsAnswerSheet.FindFirst "QuestionText='" & Me.txtQuestionText & "'"

    If rsAnswerSheet.NoMatch = False Then
        If rsAnswerSheet!SelectedOption = "A" Then
            TempVars("FinalAnswer") = "A - " & Me.OptionA
            TempVars("SelectedOption") = "A"
            Me.radOptionA.Value = True
            Me.radOptionB.Value = False
            Me.radOptionC.Value = False
            Me.radOptionD.Value = False
        ElseIf rsAnswerSheet!SelectedOption = "B" Then
            TempVars("FinalAnswer") = "B - " & Me.OptionB
            TempVars("SelectedOption") = "B"
            Me.radOptionA.Value = False
            Me.radOptionB.Value = True
            Me.radOptionC.Value = False
            Me.radOptionD.Value = False
        ElseIf rsAnswerSheet!SelectedOption = "C" Then
            TempVars("FinalAnswer") = "C - " & Me.OptionC
            TempVars("SelectedOption") = "C"
            Me.radOptionA.Value = False
            Me.radOptionB.Value = False
            Me.radOptionC.Value = True
            Me.radOptionD.Value = False
        ElseIf rsAnswerSheet!SelectedOption = "D" Then
            TempVars("FinalAnswer") = "D - " & Me.OptionD
            TempVars("SelectedOption") = "D"
            Me.radOptionA.Value = False
            Me.radOptionB.Value = False
            Me.radOptionC.Value = False
            Me.radOptionD.Value = True
        ElseIf rsAnswerSheet!SelectedOption = "True" Then
            TempVars("FinalAnswer") = "True"
            TempVars("SelectedOption") = "True"
            Me.radOptionA.Value = True
            Me.radOptionB.Value = False
        ElseIf rsAnswerSheet!SelectedOption = "False" Then
            TempVars("FinalAnswer") = "False"
            TempVars("SelectedOption") = "False"
            Me.radOptionA.Value = False
            Me.radOptionB.Value = True
        End If
        Exit Sub
    End If

    'TempVars("FinalAnswer") = Null
    'TempVars("SelectedOption") = Null
    'TempVars("CorrectAnswer") = Null


btnNextQuestion_Click_Exit:
    Exit Sub

btnNextQuestion_Click_Err:
    MsgBox Error$
    Resume btnNextQuestion_Click_Exit

End Sub
