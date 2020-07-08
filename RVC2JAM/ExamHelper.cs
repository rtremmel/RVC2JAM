using System.Data;
using VectorSolutions;

namespace RVC2JAM
{
    internal class ExamHelper
    {
        public static void GenerateExam(Course course)
        {
            RLTLIB2.Log("Creating exam images folder...");
            RLTLIB2.CreateFolderWithDelete(course.ExamImages);

            // Load questions and answers
            DataTable exam = Course.GetCourseExam(course);

            // Determine number of questions
            int lastQuestionSeq = 0;
            int questionCount = 0;
            foreach (DataRow row in exam.Rows)
                if ((int) row["question_seq"] > lastQuestionSeq)
                {
                    questionCount++;
                    lastQuestionSeq = (int) row["question_seq"];
                }

            RLTLIB2.Log($"Loaded {RLTLIB2.Pluralize(exam.Rows.Count, "assessment record")} with {RLTLIB2.Pluralize(questionCount, "question")}");

            foreach (DataRow row in exam.Rows)
            {
                ExamImageHelpers.LocalizeExamImages(
                    course,
                    $"{course.CatalogItemId} question",
                    row["assessment_question_id"].ToString(),
                    row["question_distractor_id"].ToString(),
                    row["question_material"].ToString());

                ExamImageHelpers.LocalizeExamImages(course,
                    course.CatalogItemId + " answer",
                    row["assessment_question_id"].ToString(),
                    row["question_distractor_id"].ToString(),
                    row["answer_material"].ToString());
            }
        }
    }
}