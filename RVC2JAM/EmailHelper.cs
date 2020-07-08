using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;
using VectorSolutions;

namespace RVC2JAM
{
    internal class EmailHelper
    {
        public static void EmailProcessingResults(List<Course> courses, string elapsed)
        {
            string subject = $"{RLTLIB2.AppNameAbbrVersion}";
            const string beginDetail = "<!--BEGIN-DETAIL-->\n";
            const string endDetail = "<!--END-DETAIL-->\n";
            string emailReportFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{RLTLIB2.AppAbbr}.html");
            string emailSummaryFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{RLTLIB2.AppAbbr}-SUMMARY.html");

            // Start html email message
            string body = "<!DOCTYPE html>\n";
            body += "<html lang='en'><head><title>Email</title><style>\n";
            body += "body { font-family: sans-serif; font-size: 9pt; }\n";
            body += "h1, h2 { font-size: 200%; text-align: center; color: navy; width: 750px; margin: 0; padding: 0; }\n";
            body += "h2 { font-size: 150%; margin-bottom: 1em; }\n";
            body += "table, div { width: 750px; }\n";
            body += "table, th, td { border: solid 1px gray; border-collapse: collapse; }\n";
            body += "th { background-color: yellow; }\n";
            body += "td { background-color: #eeeeff; padding: 2px 5px; }\n";
            body += ".bold { font-weight: bold; }\n";
            body += ".center { text-align: center; }\n";
            body += ".centerbold { text-align: center; font-weight: bold; }\n";
            body += "</style></head><body>\n";
            body += $"<h1>{RLTLIB2.AppNameAbbrVersion}</h1>\n";
            body += $"<h2>{DateTime.Now}</h2>\n";

            body += "<table><tr><th colspan=\'2\'>CONFIGURATION SETTINGS</th></tr>\n";
            foreach (string key in ConfigurationManager.AppSettings)
                body +=
                    $"<tr><td>{key}</td><td>{(!key.ToLower().Contains("password") ? ConfigurationManager.AppSettings[key] : new string('*', 10))}</td></tr>\n";
            body += "</table><br><br>\n";
            body += beginDetail;

            body += "<table><tr><th colspan='5'>COURSES PROCESSED</th></tr>\n";
            body += "<tr><th>rv_sku</th><th>title</th><th>type</th><th>hours</th><th>Size</th></tr>\n";

            decimal totalHours = 0;
            long totalSize = 0;
            foreach (Course course in courses)
            {
                body += "<tr>";
                body += string.Format("<td>{0}</td>", course.RvSku);
                body += string.Format("<td>{0}</td>", course.Title);
                body += string.Format("<td>{0}</td>", course.CourseType.Substring(1).TrimEnd(']'));
                body += string.Format("<td class='center'>{0}</td>", course.Hours);
                body += string.Format("<td class='center'>{0}</td>", RLTLIB2.FormatBytes(course.FinalDirectorySizeInByes).Replace(" ", "&nbsp;"));
                body += "</tr>\n";
                totalHours += course.Hours;
                totalSize += course.FinalDirectorySizeInByes;
            }

            body +=
                string.Format(
                    "<tr><td class='centerbold'>SUMMARY</td><td class='bold' colspan='2'>Processed {0} in {1}</td><td class='centerbold'>{2}</td><td class='centerbold'>{3}</td></tr>\n",
                    RLTLIB2.Pluralize(courses.Count, "course"), elapsed, totalHours, RLTLIB2.FormatBytes(totalSize).Replace(" ", "&nbsp;"));

            body += "</table><br><br>\n";

            if (ContentSet.ExcludedCourses() != "''")
            {
                body += "NOTE: These courses have been manually EXCLUDED because they cannot be processed automatically: ";
                body += ContentSet.ExcludedCourses();
                body += "<br><br>\n";
            }

            if (ContentSet.BadCourses() != "''")
            {
                body += "NOTE: These courses have been manually EXCLUDED because they are missing content: ";
                body += ContentSet.BadCourses();
                body += "<br><br>\n";
            }

            body += "<table><tr><th>WARNINGS</th></tr>\n";
            if (RLTLIB2.LogWarnings.Count == 0)
                body += "<tr><td class='center'>None</td></tr>\n";
            else
                foreach (string warning in RLTLIB2.LogWarnings)
                    body += string.Format("<tr><td>{0}</td></tr>\n", warning);
            body += "</table><br><br>\n";

            body += "<table><tr><th>ERRORS</th></tr>\n";
            if (RLTLIB2.LogErrors.Count == 0)
                body += "<tr><td class='center'>None</td></tr>\n";
            else
                foreach (string error in RLTLIB2.LogErrors)
                    body += string.Format("<tr><td>{0}</td></tr>\n", error);
            body += "</table><br><br>\n";
            body += endDetail;

            // End html email message
            body += "</body></html>\n";

            // Save full report as attachment
            RLTLIB2.WriteTextFile(emailReportFilePath, body, Encoding.UTF8);

            // Modify full report for email body summary
            List<string> detail = RLTLIB2.SearchString(body, beginDetail, endDetail, "", "", false);

            string inline = "<table><tr><th colspan='2'>SUMMARY</th></tr>\n";
            inline += $"<tr><td>Course control spreadsheet total courses</td><td class='center'>{ContentSet.ContentControlTotalCount:n0}</td></tr>\n";
            inline += $"<tr><td>Course control spreadsheet selected courses</td><td class='center'>{ContentSet.ContentControlSelectedCount:n0}</td></tr>\n";
            inline += $"<tr><td>Courses successfully processed</td><td class='center'>{ContentSet.ContentControlSuccessful:n0}</td></tr>\n";
            inline += $"<tr><td>Courses unsuccessfully processed</td><td class='center'>{ContentSet.ContentControlUnsuccessful:n0}</td></tr>\n";
            inline += $"<tr><td>Processing warnings encountered</td><td class='center'>{RLTLIB2.LogWarnings.Count:n0}</td></tr>\n";
            inline += $"<tr><td>Processing errors encountered</td><td class='center'>{RLTLIB2.LogErrors.Count:n0}</td></tr>\n";
            inline += "</table><br><br>\n";
            inline += "<h2>Refer to attachment 'RVC2JAM-YYYYMMDDhhmmss.html' for full report</h2>\n";
            inline += $"<h2>Output files are in {ConfigurationManager.AppSettings["FinalDirectory"]}</h2>\n";
            body = body.Replace(detail[0], inline);

            // Save email summary for debugging
            RLTLIB2.WriteTextFile(emailSummaryFilePath, body, Encoding.UTF8);

            // Send email summary with full report attachment
            RLTLIB2.InsertRvEmailerService(
                ConfigurationManager.AppSettings["EmailFrom"],
                ConfigurationManager.AppSettings["EmailTo"],
                ConfigurationManager.AppSettings["EmailCc"],
                "",
                subject, true, body, emailReportFilePath);
            RLTLIB2.Log(string.Format("Email sent to {0}", ConfigurationManager.AppSettings["EmailTo"]));
        }
    }
}