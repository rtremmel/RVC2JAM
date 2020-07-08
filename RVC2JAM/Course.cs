using System;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using VectorSolutions;

// ReSharper disable StringIndexOfIsCultureSpecific.2

namespace RVC2JAM
{
    public class Course
    {
        public Course(Guid catalogItemId, string rvSku, string title, int status, decimal hours, string descr, string objectives, string launchUrl)
        {
            CatalogItemId = catalogItemId;
            RvSku = rvSku.ToUpper();
            Title = title;
            Status = status;
            Hours = hours;
            Descr = RLTLIB2.StripHtml(descr);
            Objectives = RLTLIB2.StripHtml(objectives).Replace(".", ". ");
            LaunchUrl = launchUrl.Replace("\\", "/").Replace("~/", "/");
            WorkingDirectoryPath = ConfigurationManager.AppSettings["WorkingDirectory"];
            FinalDirectoryPath = ConfigurationManager.AppSettings["FinalDirectory"];
        }

        public Guid CatalogItemId { get; set; }
        public string RvSku { get; set; }
        public string Title { get; set; }
        public int Status { get; set; }
        public decimal Hours { get; set; }
        public decimal DefMasteryScore { get; set; }
        public bool ReturnsComplete { get; set; }
        public string Descr { get; set; }
        public string Objectives { get; set; }
        public DateTime LaunchDate { get; set; }
        public string CourseType { get; set; }
        public DateTime LastModDate { get; set; }
        public DateTime LastFileDate { get; set; }
        public string LaunchUrl { get; set; }
        public string RelativeStartPath { get; set; }
        public string LaunchFileName { get; set; }

        // ReSharper disable once UnusedMember.Global
        public int Depth { get; set; }

        public string ProductionContentPath
        {
            get
            {
                if (string.IsNullOrEmpty(LaunchUrl))
                    return null;

                var content = Path.GetDirectoryName(LaunchUrl);
                if (content == null)
                    return null;

                // Hack to handle LaunchInNewWindow for GlobalEtraining
                content = content.Replace(@"\LaunchInNewWindow?c=\", @"\");

                // Handle deeply embedded CiNet launchUrls
                // Example: /extra/CINet/ACPLF00-1010CEN/ACI/ACI_Modules/Player.html
                if (content.ToLower().StartsWith(@"\extra\cinet\safety\"))
                {
                    var nextBackslash = content.ToLower().IndexOf(@"\", @"\extra\cinet\safety\".Length);
                    if (nextBackslash > 0)
                        content = content.Substring(0, nextBackslash);
                }
                else if (content.ToLower().StartsWith(@"\extra\cinet\")
                         && !content.ToLower().StartsWith(@"\extra\cinet\client\luc"))
                {
                    var nextBackslash = content.ToLower().IndexOf(@"\", @"\extra\cinet\".Length);
                    if (nextBackslash > 0)
                        content = content.Substring(0, nextBackslash);
                }

                content = content.Replace(@"~\extra", "");
                content = content.Replace(@"~\Extra", "");
                content = content.Replace(@"\extra", "");
                content = content.Replace(@"\Extra", "");

                if (!content.StartsWith(@"\"))
                    content = @"\" + content;
                content = ConfigurationManager.AppSettings["ContentSourceDirectory"] + content;
                return content;
            }
        }

        public string WorkingDirectoryPath { get; set; }
        public string FinalDirectoryPath { get; set; }
        public string ProductionPreview { get; set; }
        public string WorkingDirectoryPreview { get; set; }
        public string FinalDirectoryPreview { get; set; }
        public string FinalDirectoryScormFilePath { get; set; }
        public long FinalDirectorySizeInByes { get; set; }
        public string WorkingScormDirectoryPath { get; set; }
        public string FinalScormDirectoryPath { get; set; }
        public DateTime ProcessedDate { get; set; }
        public bool PreviouslyProcessed { get; set; }
        public string DatabaseError { get; set; }
        public string ExamImages { get; set; }
        public bool HasLessonPdf { get; set; }
        public Guid LessonUnitId { get; set; }
        public string LessonPdfTargetPath { get; set; }
        public string LessonPdfSourcePath { get; set; }

        public static Course GetCourseInfo(string rvSku)
        {
            Course course = new Course(Guid.Empty, rvSku, "", 0, 0, "", "", "");
            string sql = "SELECT TOP 1 catalog_item_id, status FROM dbo.rv_cat_catalog_item WITH (NOLOCK) ";
            sql += "WHERE rv_sku = '" + rvSku + "' ";
            sql += "AND (status IN (" + ConfigurationManager.AppSettings["CatalogStatusToProcess"] + ")) ";
            //sql += "AND (is_visible = 1) ";
            DataTable catalog = RLTLIB2.ExecuteQuery(sql, out _);

            if (catalog.Rows.Count == 0)
            {
                sql = "SELECT TOP 1 catalog_item_id, status FROM dbo.rv_cat_catalog_item WITH (NOLOCK) ";
                sql += "WHERE rv_sku = '" + rvSku + "' ";
                catalog = RLTLIB2.ExecuteQuery(sql, out _);
                if (catalog.Rows.Count > 0)
                {
                    course.DatabaseError = $"Course {rvSku} - Catalog has status of {Status2Text(int.Parse(catalog.Rows[0]["status"].ToString()))}";
                    RLTLIB2.LogError(course.DatabaseError);
                }
                else
                {
                    course.DatabaseError = $"Course {rvSku} - Catalog record not found";
                    RLTLIB2.LogError(course.DatabaseError);
                }

                return course;
            }

            //if (catalog.Rows.Count > 1)
            //{
            //    course.DatabaseError = $"Course {rvSku} - Multiple catalog entries found with a status of {ConfigurationManager.AppSettings["CatalogStatusToProcess"]}";
            //    RLTLIB2.LogError(course.DatabaseError);
            //    return course;
            //}

            Guid catalogItemId = new Guid(catalog.Rows[0]["catalog_item_id"].ToString());
            string status = catalog.Rows[0]["status"].ToString();
            course.CatalogItemId = catalogItemId;
            RLTLIB2.Log($"Course {rvSku} is valid, cid={catalogItemId}, status={status} ");

            sql = "SELECT TOP 1 ci.catalog_item_id, ";
            sql += "        ci.title, ";
            sql += "        ci.rv_sku, ";
            sql += "        ci.status, ";
            sql += "        ci.hours, ";
            sql += "        ci.descr, ";
            sql += "        ci.objectives, ";
            sql += "        cu.course_unit_id, ";
            sql += "        cu.def_mastery_score, ";
            sql += "        cu.course_unit_id, ";
            sql += "        cu.launch_url, ";
            sql += "        cu.is_pdf_available, ";
            sql += "        cu.title AS lesson_title, ";
            sql += "        ci.launch_date, ";
            sql += "        ci.last_mod_date, ";
            sql += "        CASE WHEN cu.import_schema_id = 'EF31DCC7-7461-4197-8DA7-B0571A732966' THEN 1 ";
            sql += "             ELSE 0 END AS returns_complete, ";
            sql += "        CASE WHEN ci.account_id LIKE '15%' ";
            sql += "                  AND (cu.launch_url LIKE '%index_lms.html%' ";
            sql += "                       OR cu.launch_url LIKE '%index_lms.html%' ";
            sql += "                      ) THEN '[CiNet GRAY]' ";
            sql += "             WHEN ci.account_id LIKE '15%' ";
            sql += "                  AND cu.launch_url LIKE '%a001index.html%' THEN '[CiNet NEW ]' ";
            sql += "             WHEN ci.account_id LIKE '15%' THEN '[CiNet BLUE]' ";
            sql += "             WHEN ci.rv_sku LIKE 'ST-%' THEN '[SMARTTEAM ]' ";
            sql += "             ELSE '[REDVECTOR ]' ";
            sql += "        END AS course_type ";
            sql += "FROM    dbo.rv_cat_catalog_item AS ci WITH (NOLOCK) ";
            sql += "        INNER JOIN dbo.uvw_rvm_cat_course_list AS ccl WITH (NOLOCK) ON ccl.catalog_item_id = ci.catalog_item_id ";
            sql += "        INNER JOIN dbo.rv_crs_course_unit AS cu WITH (NOLOCK) ON ccl.course_unit_id = cu.course_unit_id ";
            sql += "WHERE   ci.catalog_item_id = '" + catalogItemId + "' ";
            sql += "        AND (ci.status IN (" + ConfigurationManager.AppSettings["CatalogStatusToProcess"] + ")) ";
            sql += "        AND (cu.launch_url IS NOT NULL AND cu.launch_url <> '') ";
            //sql += "        AND (ci.is_visible = 1) ";
            sql += "        AND (cu.status = 1) ";
            sql += "        AND ((ccl.course_type = 3 AND ci.rv_sku NOT LIKE 'RVTN-%') ";
            sql += "        OR (ccl.course_type IN (2,4) AND ci.rv_sku LIKE 'RVTN-%') ";
            sql += "        OR (ccl.course_type IN (4) AND ci.rv_sku LIKE 'RV-11416CA') ";
            sql += "        OR (ccl.course_type IN (4) AND ci.rv_sku LIKE 'RV-11443CA') ";
            sql += "        OR (cu.title LIKE 'lesson%' AND ci.rv_sku LIKE 'RVI-%')) ";
            var dt = RLTLIB2.ExecuteQuery(sql, out _);
            if (dt.Rows.Count == 0)
            {
                course.DatabaseError = $"Course {rvSku} - No course unit information found for catalogItemId={catalogItemId}";
                RLTLIB2.LogError(course.DatabaseError);
                return course;
            }

            //if (dt.Rows.Count > 1)
            //{
            //    course.DatabaseError = $"Course {rvSku} - Multiple course unit entries found for catalogItemId={catalogItemId}";
            //    RLTLIB2.LogError(course.DatabaseError);
            //    return course;
            //}

            course.Title = dt.Rows[0]["title"].ToString();
            course.Status = int.Parse(dt.Rows[0]["status"].ToString());
            course.Hours = decimal.Parse(dt.Rows[0]["hours"].ToString());
            course.DefMasteryScore = decimal.Parse(dt.Rows[0]["def_mastery_score"].ToString());
            if (course.DefMasteryScore == 0) course.DefMasteryScore = 80;
            course.ReturnsComplete = dt.Rows[0]["returns_complete"].ToString() == "1";
            course.Descr = RLTLIB2.StripHtml(dt.Rows[0]["descr"].ToString());
            course.Objectives = RLTLIB2.StripHtml(dt.Rows[0]["objectives"].ToString()).Replace(".", ". ");
            course.LaunchUrl = dt.Rows[0]["launch_url"].ToString();
            course.LessonUnitId = new Guid(dt.Rows[0]["course_unit_id"].ToString());
            course.CourseType = dt.Rows[0]["course_type"].ToString();
            if (!string.IsNullOrEmpty(dt.Rows[0]["launch_date"].ToString()))
                course.LaunchDate = (DateTime) dt.Rows[0]["launch_date"];
            if (!string.IsNullOrEmpty(dt.Rows[0]["last_mod_date"].ToString()))
                course.LaunchDate = (DateTime) dt.Rows[0]["last_mod_date"];

            course.HasLessonPdf = bool.Parse(dt.Rows[0]["is_pdf_available"].ToString());
            course.LessonPdfSourcePath = $@"{ConfigurationManager.AppSettings["LessonPdfDirectory"]}\{course.LessonUnitId}.pdf";
            course.LessonPdfTargetPath =
                $@"{ConfigurationManager.AppSettings["FinalDirectory"]}\_PDF\{course.RvSku}\{new Regex("[^a-zA-Z0-9]").Replace(dt.Rows[0]["lesson_title"].ToString(), "")}.pdf";

            // Handle LearnSmart courses
            if (course.RvSku.StartsWith("RVLS-"))
            {
                course.LaunchUrl = course.LaunchUrl.Replace("https://www.redvector.com", "");
                course.LaunchUrl = course.LaunchUrl.Replace("?ls=1", "");
            }

            // Do not process third party courses
            if (course.LaunchUrl.ToLower().Contains("http") || course.LaunchUrl.ToLower().Contains("gbes?"))
            {
                course.DatabaseError = $"Course {rvSku} - Third party course with launchUrl={course.LaunchUrl}";
                RLTLIB2.LogError(course.DatabaseError);
                return course;
            }

            // Do not process legacy HTML courses
            if (course.LaunchUrl.ToLower().Contains("launchpad/generator/course/html"))
            {
                course.DatabaseError = $"Course {rvSku} - Legacy HTML course with launchUrl={course.LaunchUrl}";
                RLTLIB2.LogError(course.DatabaseError);
                return course;
            }

            // Do not process assessment courses
            if (course.LaunchUrl.ToLower().Contains("launchpad/generator/assessment"))
            {
                course.DatabaseError = $"Course {rvSku} - Assessment-only course with launchUrl={course.LaunchUrl}";
                RLTLIB2.LogError(course.DatabaseError);
                return course;
            }

            return course;
        }

        public static void LoadCourseContent(ref Course course)
        {
            course.WorkingScormDirectoryPath = $@"{ConfigurationManager.AppSettings["WorkingDirectory"]}\_SCORM";
            course.FinalScormDirectoryPath = $@"{ConfigurationManager.AppSettings["FinalDirectory"]}\_SCORM";
            course.WorkingDirectoryPath = $@"{ConfigurationManager.AppSettings["WorkingDirectory"]}\{course.RvSku}";
            course.FinalDirectoryPath = $@"{ConfigurationManager.AppSettings["FinalDirectory"]}\{course.RvSku}";
            course.FinalDirectoryScormFilePath = $@"{course.FinalDirectoryPath}\{course.RvSku}.zip";
            course.LaunchFileName = Path.GetFileName(course.LaunchUrl);
            course.ExamImages = Path.Combine(course.WorkingDirectoryPath, "exam_images");

            // Handle deeply embedded CiNet launchUrls
            // Example: /extra/CINet/ACPLF00-1010CEN/ACI/ACI_Modules/Player.html
            Debug.Assert(course.LaunchUrl != null, "course.LaunchUrl != null");
            if (course.LaunchUrl.ToLower().StartsWith("/extra/cinet/safety/"))
            {
                course.LaunchFileName = course.LaunchUrl.Substring("/extra/cinet/safety/".Length);
                course.LaunchFileName = course.LaunchFileName.Substring(course.LaunchFileName.IndexOf("/", 1) + 1);
            }
            else if (course.LaunchUrl.ToLower().StartsWith("/extra/cinet/") &&
                     !course.LaunchUrl.ToLower().StartsWith("/extra/cinet/client/luc"))
            {
                course.LaunchFileName = course.LaunchUrl.Substring("/extra/cinet/".Length);
                course.LaunchFileName = course.LaunchFileName.Substring(course.LaunchFileName.IndexOf("/", 1) + 1);
            }

            course.RelativeStartPath = course.LaunchFileName;
            if (course.LaunchFileName != null)
            {
                var depth = course.LaunchFileName.Count(ch => ch == '/');
                if (depth > 0)
                    for (var i = 0; i <= depth; i++)
                        course.RelativeStartPath = "../" + course.RelativeStartPath;
            }

            //// Yet another weird CiNet launch method
            //if (course.launchFileName != null)
            //    if (course.launchFileName.ToLower().StartsWith("industrial/modules") && course.launchFileName.ToLower().EndsWith("module.html"))
            //        course.relativeStartPath = "../../../LhccCommon/Start.html";

            RLTLIB2.Log($"\tcourse.LaunchUrl={course.LaunchUrl}");
            RLTLIB2.Log($"\tcourse.LaunchFileName={course.LaunchFileName}");
            RLTLIB2.Log($"\tcourse.RelativeStartPath={course.RelativeStartPath}");

            string productionPreviewPath = course.LaunchUrl.ToLower().Replace("/extra", "rv_extra");
            string workingDirectoryPreviewPath = $"{RLTLIB2.AppAbbr}/working/{course.RvSku}/{course.LaunchFileName}";
            string finalDirectoryPreviewPath = $"{RLTLIB2.AppAbbr}/final/{course.RvSku}/{course.LaunchFileName}";

            course.ProductionPreview = $"{ConfigurationManager.AppSettings["ProductionPreviewUrl"]}?path={productionPreviewPath}";
            course.WorkingDirectoryPreview = $"{ConfigurationManager.AppSettings["LocalPreviewUrl"]}?height=1000&&path={workingDirectoryPreviewPath}";
            course.FinalDirectoryPreview = $"{ConfigurationManager.AppSettings["ProductionPreviewUrl"]}?height=1000&path={finalDirectoryPreviewPath}";
            RLTLIB2.Log("Course content: " + course.ProductionContentPath);

            if (Directory.Exists(course.FinalDirectoryPath) && !bool.Parse(ConfigurationManager.AppSettings["ReprocessAllCourses"]))
            {
                course.ProcessedDate = new DirectoryInfo(course.FinalDirectoryPath).CreationTime;
                course.LastFileDate = RLTLIB2.DirMaxDate(new DirectoryInfo(course.FinalDirectoryPath));
                course.WorkingDirectoryPath = "";
                course.WorkingDirectoryPreview = "";
                course.FinalDirectorySizeInByes = RLTLIB2.DirSize(new DirectoryInfo(course.FinalDirectoryPath));
                RLTLIB2.Log("Course previously processed on " + course.ProcessedDate.ToString("M/d/yyyy 'at' h:mm tt"));
                course.PreviouslyProcessed = true;
            }
            else
            {
                course.FinalDirectorySizeInByes = 0;
                course.ProcessedDate = DateTime.MinValue;

                RLTLIB2.CreateFolderWithDelete(course.WorkingDirectoryPath);

                if (!Directory.Exists(course.ProductionContentPath))
                {
                    course.DatabaseError = $"Course {course.RvSku} - Content does not exist at {course.ProductionContentPath}";
                    RLTLIB2.LogError(course.DatabaseError);
                    Debug.Assert(course.LaunchFileName != null, "course.LaunchFileName != null");
                    RLTLIB2.WriteTextFile(Path.Combine(course.WorkingDirectoryPath, course.LaunchFileName), course.DatabaseError);
                }
                else if (!bool.Parse(ConfigurationManager.AppSettings["DisableCourseProcessing"]))
                {
                    RLTLIB2.Log("Copying content to " + course.WorkingDirectoryPath);
                    RLTLIB2.CopyFolderTree(course.ProductionContentPath, course.WorkingDirectoryPath, "*.*", "*.bak;*.zip",
                        RLTLIB2.CopyFolderTreeLoggingLevel.Hundred);
                    course.LastFileDate = RLTLIB2.DirMaxDate(new DirectoryInfo(course.WorkingDirectoryPath));
                }
            }
        }

        public static DataTable GetCourseExam(Course course)
        {
            var sql = "";
            sql += "SELECT  aqt.code_ref AS question_type, ";
            sql += "        aq.seq AS question_seq, ";
            sql += "        aq.assessment_question_id, ";
            sql += "        aq.material AS question_material, ";
            sql += "        aqd.seq AS answer_seq, ";
            sql += "        aqd.question_distractor_id, ";
            sql += "        aqd.material AS answer_material, ";
            sql += "        aqd.answer ";
            sql += "FROM    dbo.rv_cat_catalog_item AS ci WITH (NOLOCK) ";
            sql += "        INNER JOIN dbo.uvw_rvm_cat_course_list AS ccl WITH (NOLOCK) ON ccl.catalog_item_id = ci.catalog_item_id ";
            sql += "        INNER JOIN dbo.rv_crs_course_unit AS cu WITH (NOLOCK) ON ccl.course_unit_id = cu.course_unit_id ";
            sql += "        INNER JOIN dbo.rv_crs_assessment_objective AS ao WITH (NOLOCK) ON cu.course_unit_id = ao.course_unit_id ";
            sql += "        INNER JOIN dbo.rv_crs_assessment_question AS aq WITH (NOLOCK) ON aq.assessment_objective_id = ao.assessment_objective_id ";
            sql +=
                "        INNER JOIN dbo.rv_crs_assessment_question_distractor AS aqd WITH (NOLOCK) ON aq.assessment_question_id = aqd.assessment_question_id ";
            sql += "        INNER JOIN dbo.rv_crs_assessment_question_type AS aqt WITH (NOLOCK) ON aq.question_type_id = aqt.question_type_id ";
            sql += "WHERE   ci.catalog_item_id = '" + course.CatalogItemId + "' ";
            sql += "        AND ao.status = 1 ";
            sql += "        AND aq.status = 1 ";
            sql += "        AND ci.status = " + course.Status + " ";
            //sql += "        AND ci.is_visible = 1 ";
            sql += "        AND ccl.course_type IN (2,4) ";
            sql += "ORDER BY question_seq, ";
            sql += "        answer_seq ";
            // MULTI_CHOICE, TRUE_FALSE, SINGLE_CHOICE_COMM, SINGLE_CHOICE

            var dt = RLTLIB2.ExecuteQuery(sql, out _);

            if (dt.Rows.Count == 0)
                RLTLIB2.LogWarning($"Course {course.RvSku} - Does not have any exam questions");

            return dt;
        }

        public static string Status2Text(int status)
        {
            switch (status)
            {
                case 0: return "0 (Disabled)";
                case 1: return "1 (Active)";
                case 2: return "2 (Archived)";
                case 3: return "3 (ReportOnly";
                case 4: return "4 (AssignOnly)";
                case 5: return "5 (InDev)";
                default: return "Unknown";
            }
        }
    }
}