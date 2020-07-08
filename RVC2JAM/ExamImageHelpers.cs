using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using HtmlAgilityPack;
using VectorSolutions;

namespace RVC2JAM
{
    public static class ExamImageHelpers

    {
        public static string LocalizeExamImages(Course course, string reference, string qid, string aid, string material)
        {
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(material);
            if (htmlDoc.ParseErrors != null && htmlDoc.ParseErrors.Any())
                DocumentParsingErrors(reference, htmlDoc);

            if (htmlDoc.DocumentNode != null)
            {
                HtmlNodeCollection images = htmlDoc.DocumentNode.SelectNodes("//img");
                if (images != null)
                    foreach (HtmlNode image in images)
                    {
                        string src = image.Attributes["src"].Value;
                        string absoluteUrl = CleanRelativeUrl(src);
                        string relativeUrl = DownloadExamImage(course, reference, absoluteUrl);
                        image.Attributes["src"].Value = relativeUrl;
                        //RLTLIB2.Log($"\t{reference} image was: {src}");
                        //RLTLIB2.Log($"\t{reference} image now: {relativeUrl}");

                        string update = $"{course.RvSku},{course.CatalogItemId},{qid},{aid},{src},{relativeUrl}";
                        StreamWriter sw = new StreamWriter("ExamImagesUpdates.csv", true);
                        sw.WriteLine(update);
                        sw.Close();
                    }

                return htmlDoc.DocumentNode.InnerHtml;
            }

            return material;
        }

        private static string CleanRelativeUrl(string src)
        {
            string path = src.Replace(@"\", "/");
            path = path.TrimStart('~');
            path = path.TrimStart('.');
            path = path.TrimStart('/');
            var url = path.ToLower().Contains("redvector.com") ? path : "https://www.redvector.com/" + path;
            //RLTLIB2.Log(string.Format("CleanRelativeUrl: {0} --> {1}", src, url));
            return url;
        }

        private static void DocumentParsingErrors(string reference, HtmlDocument htmlDoc)
        {
            foreach (
                HtmlParseError error in htmlDoc.ParseErrors)
                RLTLIB2.Log($"WARNING: HTML parsing error for {reference} in Line {error.Line} Column {error.LinePosition} - {error.Reason}");
        }


        private static string DownloadExamImage(Course course, string reference, string absoluteUrl)
        {
            // ReSharper disable once AssignNullToNotNullAttribute
            string localPath = Path.Combine(course.ExamImages, Path.GetFileName(absoluteUrl));

            //<img alt="" src="/CoursesDev/bba57e5664c9e711a97d02ec32550f44/images/AOIMP_GP25.gif"
            string relativeUrl = $"/CoursesDev/{course.RvSku}/exam_images/{Path.GetFileName(absoluteUrl)}";

            if (File.Exists(localPath))
                if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalization"]))
                    RLTLIB2.Log($"\tLocal file '{localPath}' already exists");

            using (WebClient wc = new WebClient())
            {
                try
                {
                    wc.DownloadFile(absoluteUrl, localPath);
                    long size = new FileInfo(localPath).Length;
                    if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalizationDownloads"]))
                        RLTLIB2.Log($"\tDownloaded '{localPath}' ({RLTLIB2.FormatBytes(size)})");
                }
                catch (Exception ex)
                {
                    RLTLIB2.LogError($"Downloading {absoluteUrl} for {reference} in {localPath} - {ex.Message}");
                }
            }

            if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalization"]))
                RLTLIB2.Log($"\tLocalized  '{absoluteUrl}' --> {relativeUrl}");

            return relativeUrl;
        }
    }
}