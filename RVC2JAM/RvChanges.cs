using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Text;
using VectorSolutions;

namespace RVC2JAM
{
    internal class RvChanges
    {
        public static string ContentHacks(Course course, FileInfo fi, string text)
        {
            string s;
            string r;

            //byte[] bytes = Encoding.Default.GetBytes(text);
            //string text8 = Encoding.UTF8.GetString(bytes);

            //text = ContentHack("2.1", text, "�", "&copy;");
            text = ContentHack("5.1", fi, text, "'xceedIntroPlayer.swf'", "'images/xceedIntroPlayer.swf'");
            text = ContentHack("5.1", fi, text, "'xceedIntroVideo.flv'", "'images/xceedIntroVideo.flv'");

            // HACK RV-6685, Lectora Publisher v.12.1.1(9935), images/xceedintroplayer.swf
            text = ContentHack("5.4", fi, text, "xceedintroplayer.swf", "xceedIntroPlayer.swf");
            text = ContentHack("5.4", fi, text, "xceedintrovideo.flv", "xceedIntroVideo.flv");

            if (text.Contains("xceedIntroPlayer.swf"))
            {
                string source = ConfigurationManager.AppSettings["ContentSourceDirectory"] + @"\introVideos\xceedIntroPlayer.swf";
                string target = string.Format(@"{0}\images\xceedIntroPlayer.swf", course.WorkingDirectoryPath);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK 5.2: {fi.Name}> Copy {source} --> {target.Replace(course.WorkingDirectoryPath, "")}");

                string targetDirectory = Path.GetDirectoryName(target);
                if (targetDirectory != null && !Directory.Exists(targetDirectory))
                    Directory.CreateDirectory(targetDirectory);

                File.Copy(source, target, true);

                source = ConfigurationManager.AppSettings["ContentSourceDirectory"] + @"\introVideos\xceedIntroVideo.flv";
                target = string.Format(@"{0}\images\xceedIntroVideo.flv", course.WorkingDirectoryPath);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK 5.3: {fi.Name}> Copy {source} --> {target.Replace(course.WorkingDirectoryPath, "")}");
                File.Copy(source, target, true);
            }

            // HACK Replace Submit Issue Page Type 3
            s = "var httpReq = getHTTP( \'../../../../lms20/tools/contentIssueForm.aspx\', \'POST\', params );";
            while (text.Contains(s))
            {
                string replacementFileName = ReplaceSubmitIssuePageType3(course);
                r = string.Format("window.open('{0}','_blank','height=400,width=400');", replacementFileName);
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK 3.3: {fi.Name}> {s} --> {r}");
            }

            //// HACK Replace Submit Issue Page Type 4; Example: RVFC-5415, page816.html
            //List<string> matches = RLTLIB2.SearchString(text,
            //    "ObjLayerActionGoToNewWindow", ");", "Trivantis_Submit_Issue_Form",
            //    "Your browser does not support dynamic html.", true);

            //foreach (string match in matches)
            //{
            //    string replacementFileName = ReplaceSubmitIssuePageType4(course);
            //    r = string.Format("window.open('{0}','_blank','height=400,width=400');", replacementFileName);
            //    text = text.Replace(match, r);
            //    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
            //        RLTLIB2.Log($"\tHACK 3.4: {fi.Name}> {match} --> {r}");
            //}

            string ciOld = "../../CommonCourseLib.js";
            string ciNew = "CommonCourseLib.js";
            if (text.Contains(ciOld))
            {
                text = ContentHack("6.2", fi, text, ciOld, ciNew);

                Debug.Assert(fi.DirectoryName != null, "fi.DirectoryName != null");
                string newFilePath = Path.Combine(ConfigurationManager.AppSettings["WorkingDirectory"], fi.DirectoryName, ciNew);
                string replacementFilePath = Path.Combine(@"LhccSupport\", ciNew);
                if (!File.Exists(newFilePath))
                {
                    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        RLTLIB2.Log($"\tHACK 6.2: {fi.Name}> Writing {newFilePath}");
                    string code = RLTLIB2.ReadTextFile(replacementFilePath);
                    RLTLIB2.WriteTextFile(newFilePath, code, Encoding.UTF8);
                }
            }

            s = "var LMSStudentName = top.userFullName;";
            r = "var LMSStudentName = 'Unknown Student';";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h01: {fi.Name}> {s} --> {r}");
            }

            s = "var LMSStudentEmail = top.userEmail;";
            r = "var LMSStudentEmail = 'none@none.com';";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h02: {fi.Name}> {s} --> {r}");
            }

            s = "top.window";
            r = "window.parent";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h03: {fi.Name}> {s} --> {r}");
            }

            s = "top.location";
            r = "window.parent";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h04: {fi.Name}> {s} --> {r}");
            }

            s = "top.Unload()";
            r = "window.parent.Unload()";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h05: {fi.Name}> {s} --> {r}");
            }

            s = "top.close()";
            r = "window.parent.close()";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h06: {fi.Name}> {s} --> {r}");
            }

            s = "top.document";
            r = "window.parent.document";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h07: {fi.Name}> {s} --> {r}");
            }

            s = "window.parent.window";
            r = "window.parent";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h98: {fi.Name}> {s} --> {r}");
            }

            s = "window.window";
            r = "window";
            while (text.Contains(s))
            {
                text = text.Replace(s, r);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log($"\tHACK TOP.h99: {fi.Name}> {s} --> {r}");
            }

            // Support Hacks
            text = ContentHack("4.1", fi, text, "customer.care@criticalinfonet.com", "clientsupport@redvector.com");
            text = ContentHack("4.2", fi, text, "http://www.criticalinfonet.com", "http://www.redvector.com");
            return text;
        }

        public static string ContentHack(string hackId, FileInfo fi, string text, string oldText, string newText)
        {
            if (text.Contains(oldText))
            {
                text = text.Replace(oldText, newText);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log(string.Format("\tHACK {0} {1}: {2} --> {3}", hackId, fi.Name, oldText, newText));
            }

            return text;
        }

        public static void ReplaceSubmitIssuePageType1(Course course, FileInfo fi)
        {
            string oldPath = fi.FullName.Replace(fi.Extension, "-old" + fi.Extension);
            string newPath = @"LhccSupport\" + fi.Name;
            File.Copy(fi.FullName, oldPath, true);

            string code = RLTLIB2.ReadTextFile(newPath);
            code = Journey.ReplaceKeyFields(course, code); // New Journey Submit Issue
            RLTLIB2.WriteTextFile(fi.FullName, code, Encoding.UTF8);
            if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                RLTLIB2.Log(string.Format("\tHACK 3.1: {0} --> {1}", newPath, fi.Name));
        }

        public static string ReplaceSubmitIssuePageType2(Course course)
        {
            string newFileName = "submitIssue.html";
            string newFilePath = Path.Combine(ConfigurationManager.AppSettings["WorkingDirectory"],
                string.Format("{0}/{1}", course.RvSku, newFileName));
            string replacementFilePath = Path.Combine(@"LhccSupport\", newFileName);
            if (!File.Exists(newFilePath))
            {
                string code = RLTLIB2.ReadTextFile(replacementFilePath);
                code = Journey.ReplaceKeyFields(course, code); // New Journey Submit Issue
                RLTLIB2.WriteTextFile(newFilePath, code, Encoding.UTF8);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log(string.Format("\tHACK 3.2: {0} --> {1}", replacementFilePath, newFileName));
            }

            return newFileName;
        }

        public static string ReplaceSubmitIssuePageType3(Course course)
        {
            string newFileName = "submitIssue.html";
            string newFilePath = Path.Combine(course.WorkingDirectoryPath, newFileName);
            string replacementFilePath = Path.Combine(@"LhccSupport\", newFileName);

            if (!File.Exists(newFilePath))
            {
                string code = RLTLIB2.ReadTextFile(replacementFilePath);
                code = Journey.ReplaceKeyFields(course, code); // New Journey Submit Issue
                RLTLIB2.WriteTextFile(newFilePath, code, Encoding.UTF8);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log(string.Format("\tHACK 3.3: {0} --> {1}", replacementFilePath, newFileName));
            }

            return newFileName;
        }

        public static string ReplaceSubmitIssuePageType4(Course course)
        {
            string newFileName = "submitIssue.html";
            string newFilePath = Path.Combine(course.WorkingDirectoryPath, newFileName);
            string replacementFilePath = Path.Combine(@"LhccSupport\", newFileName);

            if (!File.Exists(newFilePath))
            {
                string code = RLTLIB2.ReadTextFile(replacementFilePath);
                code = Journey.ReplaceKeyFields(course, code); // New Journey Submit Issue
                RLTLIB2.WriteTextFile(newFilePath, code, Encoding.UTF8);
                if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                    RLTLIB2.Log(string.Format("\tHACK 3.4: {0} --> {1}", replacementFilePath, newFileName));
            }

            return newFileName;
        }

        public static void DoSpecialFileHacks(Course course)
        {
            FileInfo[] files = new DirectoryInfo(course.WorkingDirectoryPath).GetFiles("menutext.txt", SearchOption.AllDirectories);
            foreach (FileInfo file in files)
            {
                string code = RLTLIB2.ReadTextFile(file.FullName);
                code = ContentHack("4.4", file, code, "customer.care@criticalinfonet.com", "clientsupport@redvector.com");
                code = ContentHack("4.4", file, code, "http://www.criticalinfonet.com", "http://www.redvector.com");
                code = ContentHack("4.4", file, code, "972-309-4000 / 800-624-2272", "");
                RLTLIB2.WriteTextFile(file.FullName, code);
            }
        }
    }
}