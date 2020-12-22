using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using HtmlAgilityPack;
using Newtonsoft.Json;
using VectorSolutions;

namespace RVC2JAM
{
    internal static class Program
    {
        private static void Main()
        {
            DateTime startTime = AppInit();
            if (File.Exists("BadUrls.csv"))
                File.Delete("BadUrls.csv");

            RLTLIB2.LogRepeatedChar('=', 120);
            List<Course> courses = ProcessContentSet();
            RLTLIB2.LogRepeatedChar('=', 120);

            RLTLIB2.Log("PROCESSED:");
            foreach (var course in courses)
                RLTLIB2.Log($"\t{course.RvSku} {course.CourseType} {course.Title}");
            RLTLIB2.LogRepeatedChar('=', 120);

            RLTLIB2.Log("PROCESSING SUMMARY:");
            if (RLTLIB2.LogWarnings.Count > 0)
                foreach (var warning in RLTLIB2.LogWarnings)
                    RLTLIB2.Log($"\t{warning}");
            else
                RLTLIB2.Log("\tNo Warnings");
            if (RLTLIB2.LogErrors.Count > 0)
                foreach (var error in RLTLIB2.LogErrors)
                    RLTLIB2.Log("\t" + error);
            else
                RLTLIB2.Log("\tNo Errors");

            RLTLIB2.LogRepeatedChar('=', 120);
            var elapsed = RLTLIB2.FormatElapsedTime(DateTime.Now - startTime);
            RLTLIB2.Log($"Processing completed in {elapsed}");

            RLTLIB2.Log(
                $"ContentSourceDirectory: file:///{ConfigurationManager.AppSettings["ContentSourceDirectory"]} - {RLTLIB2.GetDiskFreeSpaceInfo(ConfigurationManager.AppSettings["ContentSourceDirectory"])}");
            RLTLIB2.Log(
                $"WorkingDirectory: file:///{ConfigurationManager.AppSettings["WorkingDirectory"]} - {RLTLIB2.GetDiskFreeSpaceInfo(ConfigurationManager.AppSettings["WorkingDirectory"])}");
            RLTLIB2.Log(
                $"FinalDirectory:   file:///{ConfigurationManager.AppSettings["FinalDirectory"]} - {RLTLIB2.GetDiskFreeSpaceInfo(ConfigurationManager.AppSettings["FinalDirectory"])}");

            RLTLIB2.LogRepeatedChar('=', 120);
            EmailHelper.EmailProcessingResults(courses, elapsed);
        }

        private static List<Course> ProcessContentSet()
        {
            // Initialize client directory
            RLTLIB2.Log("Clearing " + ConfigurationManager.AppSettings["WorkingDirectory"]);
            RLTLIB2.CreateFolderWithDelete(ConfigurationManager.AppSettings["WorkingDirectory"]);
            var workingFolder = ConfigurationManager.AppSettings["WorkingDirectory"];
            RLTLIB2.CreateFolderWithDelete(workingFolder);

            if (bool.Parse(ConfigurationManager.AppSettings["ClearFinalDestinationBeforeProcessing"]))
            {
                var finalPath = workingFolder.Replace(ConfigurationManager.AppSettings["WorkingDirectory"],
                    ConfigurationManager.AppSettings["FinalDirectory"]);
                RLTLIB2.Log($"Clearing {finalPath} (This may take a while...)");
                RLTLIB2.CreateFolderWithDelete(finalPath);
            }

            DataTable CoursesToProcessTable = ContentSet.CoursesToProcessTable;
            List<Course> courses = new List<Course>();

            // Create exam images update file
            if (File.Exists("ExamImagesUpdates.csv")) File.Delete("ExamImagesUpdates.csv");
            const string update = "rv_sku,catalog_item_id,assessment_question_id,question_distractor_id,old_src,new_src";
            StreamWriter sw = new StreamWriter("ExamImagesUpdates.csv", true);
            sw.WriteLine(update);
            sw.Close();

            for (var index = 0; index < CoursesToProcessTable.Rows.Count; index++)
            {
                string rvSku = CoursesToProcessTable.Rows[index]["Legacy SKU"].ToString().Trim().ToUpper();
                //string jamSku = CoursesToProcessTable.Rows[index]["JAM SKU"].ToString().Trim().ToUpper();

                RLTLIB2.LogRepeatedChar('-', 120);
                //RLTLIB2.Log($"Course {index + 1:n0}/{CoursesToProcessTable.Rows.Count:n0} ({(index + 1) * 100.0 / CoursesToProcessTable.Rows.Count:n1}%) {rvSku} ({jamSku})");
                RLTLIB2.Log($"Course {index + 1:n0}/{CoursesToProcessTable.Rows.Count:n0} ({(index + 1) * 100.0 / CoursesToProcessTable.Rows.Count:n1}%) {rvSku}");

                if (ContentSet.ExcludedCourses().Contains(rvSku))
                {
                    RLTLIB2.Log("\tSkipped per CoursesExcluded.txt");
                }
                else if (ContentSet.BadCourses().Contains(rvSku))
                {
                    RLTLIB2.Log("\tSkipped per CoursesBad.txt (missing content)");
                }
                else if (ConfigurationManager.AppSettings["OnlyProcessThisRvSku"] != ""
                         && ConfigurationManager.AppSettings["OnlyProcessThisRvSku"] != rvSku)
                {
                    RLTLIB2.Log($"\tSkipped per OnlyProcessThisRvSku={ConfigurationManager.AppSettings["OnlyProcessThisRvSku"]}");
                }
                else
                {
                    Course course = ProcessCourse(rvSku);

                    if (!string.IsNullOrEmpty(course.DatabaseError))
                    {
                        Debug.WriteLine($"*** {course.DatabaseError}");
                    }
                    else
                    {
                        if (!bool.Parse(ConfigurationManager.AppSettings["DisableCourseProcessing"]))
                        {
                            if (!course.PreviouslyProcessed)
                                if (!bool.Parse(ConfigurationManager.AppSettings["DisableCopyingToFinalLocation"]))
                                {
                                    CopyCourseToFinalLocation(course);

                                    if (course.HasLessonPdf)
                                        if (File.Exists(course.LessonPdfSourcePath))
                                            CopyLessonPdf(course);
                                        else
                                            RLTLIB2.LogError(
                                                $"Course {course.RvSku} - Indicates PDF available but '{course.LessonPdfSourcePath}' does not exist");
                                    else
                                        RLTLIB2.Log("\tLesson PDF is not available for this course.");
                                }

                            Debug.WriteLine("*** {0}={1}", "course.finalDirectoryPath", course.FinalDirectoryPath);
                            course.FinalDirectorySizeInByes = RLTLIB2.DirSize(new DirectoryInfo(course.FinalDirectoryPath));
                            Debug.WriteLine("*** {0}={1}", "course.finalDirectorySize", RLTLIB2.FormatBytes(course.FinalDirectorySizeInByes));
                        }

                        courses.Add(course);
                    }
                }
            }

            return courses;
        }

        private static Course ProcessCourse(string rvSku)
        {
            Course course = Course.GetCourseInfo(rvSku);

            if (course.Title == "")
            {
                RLTLIB2.Log("course:");
                RLTLIB2.Log(JsonConvert.SerializeObject(course, Formatting.Indented));
                ContentSet.ContentControlUnsuccessful++;
                return course;
            }

            if (!string.IsNullOrEmpty(course.DatabaseError))
            {
                ContentSet.ContentControlUnsuccessful++;
                return course;
            }

            Course.LoadCourseContent(ref course);
            if (!string.IsNullOrEmpty(course.DatabaseError))
            {
                ContentSet.ContentControlUnsuccessful++;
                return course;
            }

            RLTLIB2.Log("course:");
            RLTLIB2.Log(JsonConvert.SerializeObject(course, Formatting.Indented));

            if (bool.Parse(ConfigurationManager.AppSettings["DisableCourseProcessing"]))
            {
                RLTLIB2.Log("Course processing disabled");
            }
            else
            {
                if (!course.PreviouslyProcessed)
                {
                    if (course.RvSku.StartsWith("RVTN-"))
                        RLTLIB2.Log("Lesson contains Exam.  Skipping Exam creation.");
                    else
                        ExamHelper.GenerateExam(course);

                    ProcessDirectory(ref course, new DirectoryInfo(course.WorkingDirectoryPath));

                    RvChanges.DoSpecialFileHacks(course);

                    ScormHelper.CreateManifestFile(course);
                    ScormHelper.CreateInfoFile(course);
                    ScormHelper.CreateScormPackage(course);

                    if (bool.Parse(ConfigurationManager.AppSettings["DeleteWorkingIntermediateFiles"]))
                        DeleteWorkingIntermediateFiles(course.WorkingDirectoryPath);

                    course.ProcessedDate = new DirectoryInfo(course.FinalDirectoryPath).CreationTime;
                }
            }

            ContentSet.ContentControlSuccessful++;
            return course;
        }

        private static void CopyCourseToFinalLocation(Course course)
        {
            RLTLIB2.Log($"Copying modified content to {course.FinalDirectoryPath}");
            RLTLIB2.CreateFolderWithDelete(course.FinalDirectoryPath);
            RLTLIB2.CopyFolderTree(course.WorkingDirectoryPath, course.FinalDirectoryPath, "*.*", "*.zip;*-tmp.*;*-old.*",
                RLTLIB2.CopyFolderTreeLoggingLevel.Hundred);
        }

        private static void CopyLessonPdf(Course course)
        {
            RLTLIB2.Log($"\tCopying lesson PDF '{course.LessonPdfSourcePath}' --> '{course.LessonPdfTargetPath}'");
            // ReSharper disable once AssignNullToNotNullAttribute
            Directory.CreateDirectory(Path.GetDirectoryName(course.LessonPdfTargetPath));
            File.Copy(course.LessonPdfSourcePath, course.LessonPdfTargetPath, true);
        }

        private static void ProcessDirectory(ref Course course, DirectoryInfo di)
        {
            RLTLIB2.Log("Directory " + di.Name);

            foreach (FileInfo fi in di.GetFiles("*.js"))
            {
                if (fi.Name.ToLower() == "player_compiled.js")
                {
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK: {fi.Name}> SKIPPED (causes bad problems)");
                    continue;
                }

                if (fi.Name.ToLower().Contains("jquery"))
                {
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK: {fi.Name}> SKIPPED (do not modify)");
                    continue;
                }

                //PerformHack16Fixes(ref course, fi);

                string text = RLTLIB2.ReadTextFile(fi.FullName, Encoding.UTF8);
                string s;
                string r;

                s = "window.top.close()";
                r = "window.parent.close()";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j01: {fi.Name}> {s} --> {r}");
                }

                s = "window.top.opener";
                r = "window.parent.opener";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j02: {fi.Name}> {s} --> {r}");
                }

                s = "top.window.close()";
                r = "window.parent.close()";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j03: {fi.Name}> {s} --> {r}");
                }

                s = "this.top";
                r = "this.window.parent";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j04 {fi.Name}> {s} --> {r}");
                }

                s = "window.top";
                r = "window.parent";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j05 {fi.Name}> {s} --> {r}");
                }

                s = "window.top._";
                r = "window.parent._";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j06 {fi.Name}> {s} --> {r}");
                }

                s = "top.addEventListener";
                r = "window.parent.addEventListener";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j07 {fi.Name}> {s} --> {r}");
                }

                s = "top.inner";
                r = "window.parent.inner";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j08 {fi.Name}> {s} --> {r}");
                }

                s = "parent.top";
                r = "window.parent";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j09 {fi.Name}> {s} --> {r}");
                }

                s = "top.opener";
                r = "window.parent.opener";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j09 {fi.Name}> {s} --> {r}");
                }

                s = "top.window.";
                r = "window.parent.";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j09 {fi.Name}> {s} --> {r}");
                }

                s = "top.close()";
                r = "window.parent.close()";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j10 {fi.Name}> {s} --> {r}");
                }

                s = "top.moveTo";
                r = "window.parent.moveTo";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j11 {fi.Name}> {s} --> {r}");
                }

                s = "window.parent.window";
                r = "window.parent";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j98: {fi.Name}> {s} --> {r}");
                }

                s = "window.window";
                r = "window";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK TOP.j99: {fi.Name}> {s} --> {r}");
                }

                s = "Español";
                r = "Espanol";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK Foreign01: {fi.Name}> {s} --> {r}");
                }

                s = "Português";
                r = "Portugues";
                while (text.Contains(s))
                {
                    text = text.Replace(s, r);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK Foreign02: {fi.Name}> {s} --> {r}");
                }

                RLTLIB2.WriteTextFile(fi.FullName, text);
            }

            foreach (var fi in di.GetFiles("*.htm*"))
            {
                if (LocalizationHelper.RvExcludedFiles.Contains(fi.Name.ToLower()))
                {
                    RLTLIB2.Log("Skipping " + fi.Name);
                    continue;
                }

                // HACK Replace Submit Issue Page Type 1
                if (fi.Name.ToLower() == "a001_submit_issue_page_1.html")
                {
                    RvChanges.ReplaceSubmitIssuePageType1(course, fi);
                    continue;
                }

                // Example: ST-0003A --> a001_02_clarification_page_3_cp1.html --> checkPoint_1.html
                if (fi.Name.ToLower().StartsWith("checkpoint_"))
                {
                    // HACK 2.4, Content Repair, Force '*checkpoint*.html' files to lowercase.
                    if (fi.Name.ToLower() != fi.Name)
                    {
                        RLTLIB2.Log($"\tHACK 2.4: {fi.Name} --> {fi.Name.ToLower()}");
                        File.Move(fi.FullName, fi.FullName + ".xxx");
                        File.Move(fi.FullName + ".xxx", fi.FullName.ToLower());
                    }

                    RLTLIB2.Log($"Skipping {fi.Name} (assumed to be nested iFrame content)");
                    continue;
                }

                var html = new HtmlDocument();
                html.Load(fi.FullName);

                // Example: RV-7730 --> a001_chapter2_page_4.html --> ch2pg04.html
                var htmlNodeCollection = html.DocumentNode.SelectNodes("//title");
                var firstOrDefault = htmlNodeCollection?.FirstOrDefault();
                if (firstOrDefault != null && firstOrDefault.InnerText == "Untitled Document")
                {
                    RLTLIB2.Log($"Skipping {fi.Name} (assumed to be nested iFrame content)");
                    continue;
                }

                var pathUnmodified = fi.FullName.Replace(fi.Extension, "-old" + fi.Extension);
                html.Save(pathUnmodified);

                var text = RLTLIB2.ReadTextFile(pathUnmodified, Encoding.UTF8);

                // Hack for stupid typo in course BW-05 index.html
                text = text
                    .Replace("<link rel=\"icon\" type=\"image/png\" href=\" https://www.redvector.com/favicon.ico\">",
                        "<link rel=\"icon\" type=\"image/png\" href=\"https://www.redvector.com/favicon.ico\">");

                text = LocalizationHelper.LocalizeUrls(course, fi, text);
                text = RvChanges.ContentHacks(course, fi, text);
                RLTLIB2.WriteTextFile(fi.FullName, text);
            }

            foreach (var dis in di.GetDirectories())
                ProcessDirectory(ref course, dis);
        }

        private static void DeleteWorkingIntermediateFiles(string clientCoursePath)
        {
            Directory.GetFiles(clientCoursePath, "*-old.html", SearchOption.AllDirectories).ToList()
                .ForEach(File.Delete);
            Directory.GetFiles(clientCoursePath, "*-tmp.html", SearchOption.AllDirectories).ToList()
                .ForEach(File.Delete);
        }

        private static DateTime AppInit()
        {
            // Create RLTLIB self documentation
            RLTLIB2.SelfDocument2Html();

            // Display program info
            var startTime = DateTime.Now;
            RLTLIB2.Log(null); // Reset log
            RLTLIB2.Log($"{RLTLIB2.AppNameAbbrVersion} started {DateTime.Now:g}");
            RLTLIB2.Log($"\t{RLTLIB2.Copyright}");
            RLTLIB2.Log($"\tRunning as '{Environment.UserName}' on '{Environment.MachineName}' in directory '{Environment.CurrentDirectory}'");
            RLTLIB2.Log($"\tThis program uses the Rick Tremmel Common Methods \'RLTLIB2.cs\' last modified {RLTLIB2.LastModifiedDate}");

            // ReSharper disable PossibleNullReferenceException
            var app = new DirectoryInfo("../../");
            var packages = Path.Combine(app.FullName, "packages.config");
            // ReSharper restore PossibleNullReferenceException
            RLTLIB2.Log("\tThis application requires the following NuGet packages:");
            var xElement = XDocument.Load(packages).Root;
            if (xElement != null)
                foreach (var package in from c in xElement.Descendants("package") select c)
                    RLTLIB2.Log($"\t\t{package.Attribute("id")}, {package.Attribute("version")}, {package.Attribute("targetFramework")}");

            foreach (string key in ConfigurationManager.AppSettings)
                RLTLIB2.Log($"\tAppSetting: {key} = {(!key.ToLower().Contains("password") ? ConfigurationManager.AppSettings[key] : new string('*', 10))}");

            LocalizationHelper.RvExcludedFiles = RLTLIB2.ReadTextFile("RvExcludedFiles.txt")
                .Split(new[] {"\r\n", "\n"}, StringSplitOptions.None);
            foreach (var filename in LocalizationHelper.RvExcludedFiles)
                RLTLIB2.Log($"\tIgnored File: {filename}");

            LocalizationHelper.RvIgnoredUrlFragments = RLTLIB2.ReadTextFile("RvIgnoredUrlFragments.txt")
                .Split(new[] {"\r\n", "\n"}, StringSplitOptions.None);
            //foreach (var url in LocalizationHelper.RvIgnoredUrlFragments)
            //    RLTLIB2.Log($"\tIgnored URL Fragment: {url}");

            return startTime;
        }
    }
}