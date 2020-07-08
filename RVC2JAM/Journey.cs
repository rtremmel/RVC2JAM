using System;
using System.Configuration;
using System.IO;
using Newtonsoft.Json;
using VectorSolutions;

namespace RVC2JAM
{
    public class Journey
    {
        private static string JourneyFolder => Path.Combine(ConfigurationManager.AppSettings["ContentSourceDirectory"], "Journey");
        private static string TemplateFolder => Path.Combine(JourneyFolder, "_TEMPLATE_");

        public static void CreateRedVectorLaunchFolder(Course course)
        {
            string courseFolder = Path.Combine(JourneyFolder, course.RvSku);
            RLTLIB2.CreateFolderWithDelete(courseFolder);
            RLTLIB2.CopyFolderTree(TemplateFolder, courseFolder, "*.*", "", RLTLIB2.CopyFolderTreeLoggingLevel.Verbose);

            string manifestFile = Path.Combine(courseFolder, "imsmanifest.xml");
            string manifestText = RLTLIB2.ReadTextFile(manifestFile);
            manifestText = ReplaceKeyFields(course, manifestText);
            RLTLIB2.WriteTextFile(manifestFile, manifestText);
            RLTLIB2.Log($"Updated '{manifestFile}'");

            string launchFile = Path.Combine(courseFolder, "launch.json");
            string launchText = RLTLIB2.ReadTextFile(launchFile);
            launchText = ReplaceKeyFields(course, launchText);
            RLTLIB2.WriteTextFile(launchFile, launchText);
            RLTLIB2.Log($"Updated '{launchFile}'");

            string infoFile = Path.Combine(courseFolder, "info.txt");
            string text = $"This folder was created by {RLTLIB2.AppNameAbbrVersion} on {DateTime.Now}\n";
            text += JsonConvert.SerializeObject(course, Formatting.Indented);
            RLTLIB2.WriteTextFile(infoFile, text);
            RLTLIB2.Log($"Added '{infoFile}'");
        }

        public static string ReplaceKeyFields(Course course, string text)
        {
            text = text.Replace("#RVSKU#", course.RvSku);
            text = text.Replace("#TITLE#", course.Title);
            text = text.Replace("#CATALOGITEMID#", course.CatalogItemId.ToString());
            text = text.Replace("#COURSEID#", "");
            text = text.Replace("#LESSONID#", "");
            text = text.Replace("#VERSION#", $"Converted by {RLTLIB2.AppNameAbbrVersion} on {DateTime.Now}");
            return text;
        }
    }
}