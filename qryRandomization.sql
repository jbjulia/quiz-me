SELECT tblQuestionBank.ID, tblQuestionBank.QuestionText, tblQuestionBank.OptionA, tblQuestionBank.OptionB, tblQuestionBank.OptionC, tblQuestionBank.OptionD, tblQuestionBank.CorrectOption, tblQuestionBank.QuestionType, tblQuestionBank.QuestionReference
FROM tblQuestionBank
ORDER BY Rnd(([ID]+1)-([ID]-1));
