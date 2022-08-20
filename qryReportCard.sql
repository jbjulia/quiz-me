SELECT tblAnswerSheet.ID, tblAnswerSheet.QuestionNo, tblAnswerSheet.QuestionText, tblAnswerSheet.FinalAnswer, tblAnswerSheet.[CorrectAnswer], tblAnswerSheet.IsCorrect
FROM tblAnswerSheet
WHERE (((tblAnswerSheet.IsCorrect)=False))
ORDER BY tblAnswerSheet.QuestionNo;
