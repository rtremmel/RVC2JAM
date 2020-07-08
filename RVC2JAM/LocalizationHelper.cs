using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using VectorSolutions;

// ReSharper disable StringIndexOfIsCultureSpecific.1
// ReSharper disable StringIndexOfIsCultureSpecific.2

namespace RVC2JAM
{
    internal class LocalizationHelper
    {
        public static string[] RvExcludedFiles;
        public static string[] RvIgnoredUrlFragments;

        public static string LocalizeUrls(Course course, FileInfo fi, string text)
        {
            //List<string> links = new List<string>();
            //links.AddRange(RLTLIB2.SearchString(text, 
            //    "'", "'", "http", "Your browser does not support dynamic html.", true, false));
            //links.AddRange(RLTLIB2.SearchString(text, 
            //    "\"", "\"", "http", "Your browser does not support dynamic html.", true, false));

            //foreach (string link in links.OrderBy(s => s).ToList())
            //{
            //    RLTLIB2.Log("LINK: " + link);
            //}            

            string[] sites =
            {
                "http://www.redvector.com",
                "https://www.redvector.com",
                "http://www.care2learn.com",
                "https://www.care2learn.com",
                "http://www.smartteam.com",
                "https://www.smartteam.com"
            };
            foreach (string site in sites)
            {
                int j = 0;
                while (j < text.Length && text.IndexOf(site, j) > -1)
                {
                    int i = text.IndexOf(site, j);
                    string delimiter = text.Substring(i - 1, 1);
                    if (delimiter == "(") delimiter = ")"; // Handle CSS URLs
                    j = text.IndexOf(delimiter, i + 1);
                    string url = text.Substring(i, j - i);
                    j++;

                    // Jump over URLs that are already relative!
                    if (text.Substring(i - 1, 1) == ".")
                        continue;

                    // Ignore non-file support URLs
                    if (url.ToLower() == "http://www.redvector.com/support") continue;
                    if (url.ToLower() == "https://www.redvector.com/support") continue;
                    if (url.ToLower() == "http://www.redvector.com/support.aspx") continue;
                    if (url.ToLower() == "https://www.redvector.com/support.aspx") continue;

                    if (url.EndsWith(".pdf\\"))
                        url = url.TrimEnd('\\');

                    url = CleanUrl(url);
                    if (!IsValidUrl(url))
                        DocumentBadUrl(course, fi, url);

                    // HACK Replace Submit Issue Page Type 2
                    if (url.ToLower() == "https://www.redvector.com/extra/submitissue/submitissue.aspx")
                    {
                        string replacementUrl = RvChanges.ReplaceSubmitIssuePageType2(course);
                        text = text.Replace(url, replacementUrl);
                        text = text.Replace(@"method=""post"" action=""submitIssue.html""", @"method=""get"" action=""submitIssue.html""");
                        if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                            RLTLIB2.Log(string.Format("\tHACK 3.2: {0} --> {1}", url, replacementUrl));
                        continue;
                    }

                    // Document URL
                    if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalization"]))
                        RLTLIB2.Log(string.Format("\tLocalizing '{0}'", url));

                    // Begin Custom URL Hack
                    if (url.ToLower() == "https://www.redvector.com/extra/rv/intro_con/flashfox.swf")
                    {
                        string hack = "https://www.redvector.com/extra/intro_con/flashfox.swf";
                        if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                            RLTLIB2.Log(string.Format("\tHACK 2.2: {0} --> {1}", url.ToLower(), hack));
                        text = text.Replace(url, hack);
                        url = hack;
                    }
                    // End Custom URL Hack

                    if (UrlContainsIgnoredFragment(url))
                    {
                        j = i + url.Length + 1;
                        continue;
                    }

                    string absoluteUrl = url;
                    string relativeUrl = url.Replace(site, "").toRelative();
                    string newUrl = relativeUrl.toForwardSlash();
                    // Plus signs in file paths cause SCORM Cloud imports to break
                    newUrl = newUrl.Replace("+", "_");

                    string localPath = string.Format(@"{0}\{1}", course.WorkingDirectoryPath, relativeUrl);
                    localPath = CleanUrl(localPath);

                    // Begin hack for 'client-support.aspx', 'contact-us', 'contact-us/', 'contact-us.aspx', and 'support/'
                    if (newUrl.ToLower().Contains("client-support.aspx") ||
                        newUrl.ToLower().Contains("contact-us") ||
                        newUrl.ToLower().Contains("contact-us/") ||
                        newUrl.ToLower().Contains("contact-us.aspx") ||
                        newUrl.ToLower().Contains("support/"))
                    {
                        localPath = string.Format(@"{0}\support.html", course.WorkingDirectoryPath);
                        string code = RLTLIB2.ReadTextFile(@"LhccSupport\support.html");
                        code = Journey.ReplaceKeyFields(course, code); // New Journey Submit Issue
                        if (!File.Exists(localPath))
                        {
                            RLTLIB2.WriteTextFile(localPath, code, Encoding.UTF8);
                            if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                                RLTLIB2.Log($"\tHACK 4.1: {@"LhccSupport\support.html"} --> {localPath}");
                        }

                        newUrl = "support.html";
                        text = text.Replace(url, newUrl);
                        j = i + newUrl.Length + 1;

                        if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalization"]))
                            RLTLIB2.Log($"\tLocalized  '{url}' --> {newUrl}");
                        continue;
                    }
                    // End hack for 'client-support.aspx', 'contact-us', 'contact-us/', 'contact-us.aspx', and 'support/'

                    // Begin SmartTeam HTML4activityFiles Hack
                    if (url.ToLower().Contains("html4activityfiles"))
                    {
                        string activitySource = string.Format(@"{0}\SmartTeam\HTML4activityFiles", ConfigurationManager.AppSettings["ContentSourceDirectory"]);
                        string activityTargetRoot = string.Format("{0}\\extra", course.WorkingDirectoryPath);
                        string activityTarget = string.Format(@"{0}\SmartTeam\HTML4activityFiles", activityTargetRoot);
                        if (!Directory.Exists(activityTargetRoot))
                            RLTLIB2.CreateFolderWithDelete(activityTargetRoot);
                        if (!Directory.Exists(activityTarget))
                        {
                            if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                                RLTLIB2.Log(string.Format("\tHACK 6.1: {0} --> {1}", activitySource, activityTarget));
                            RLTLIB2.CopyFolderTree(activitySource, activityTarget, "*.*", "", RLTLIB2.CopyFolderTreeLoggingLevel.Hundred);
                        }

                        text = text.Replace(url, newUrl);
                        j = i + newUrl.Length + 1;

                        if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalization"]))
                            RLTLIB2.Log(string.Format("\tLocalized  '{0}' --> {1}", url, newUrl));
                        continue;
                    }
                    // End SmartTeam HTML4activityFiles Hack

                    //if (relativeUrl.ToLower().Contains("extra") && !relativeUrl.ToLower().Contains("intro_con"))
                    //    throw new SystemException("Oh No Mr. Bill!  Possible missed common files.");

                    // Lightlight CDN Hack
                    if (newUrl.StartsWith("o3/"))
                    {
                        Debug.Assert(course.WorkingDirectoryPath != null, "course.workingDirectoryPath != null");
                        Debug.Assert(fi.Directory != null, "fi.Directory != null");
                        Debug.Assert(fi.Directory.Parent != null, "fi.Directory.Parent != null");
                        if (course.WorkingDirectoryPath.ToLower() == fi.Directory.FullName.ToLower()
                            && !text.Contains("data='media/player.swf'")) // RVI-10884
                            if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                                RLTLIB2.Log("\tHACK 7.3: Ignored, same directory");
                        //else
                        //{
                        //    string oldnewUrl = newUrl;
                        //    newUrl = newUrl.Replace("o3/", "../o3/");
                        //    if (bool.Parse(ConfigurationManager.AppSettings["LogHacks"]))
                        //        RLTLIB2.Log(string.Format("\tHACK 7.3: {0} --> {1}", oldnewUrl, newUrl));
                        //}
                    }

                    if (File.Exists(localPath))
                    {
                        if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalization"]))
                            RLTLIB2.Log(string.Format("\tLocal file '{0}' already exists", localPath));
                    }
                    else
                    {
                        string localFolder = new FileInfo(localPath).DirectoryName;
                        Debug.Assert(localFolder != null, "localFolder != null");
                        if (!Directory.Exists(localFolder))
                            Directory.CreateDirectory(localFolder);

                        if (bool.Parse(ConfigurationManager.AppSettings["DisableDownloadingMp4WebmOgv"]) &&
                            (url.ToLower().EndsWith(".mp4") ||
                             url.ToLower().EndsWith(".webm") ||
                             url.ToLower().EndsWith(".ogv")))
                        {
                            RLTLIB2.Log(string.Format("\tFile '{0}' NOT downloaded for fast testing", localPath));
                        }
                        else
                        {
                            // Plus signs in file paths cause SCORM Cloud imports to break
                            localPath = localPath.Replace("+", "_");

                            using (WebClient wc = new WebClient())
                            {
                                try
                                {
                                    wc.DownloadFile(absoluteUrl, localPath);
                                    long size = new FileInfo(localPath).Length;
                                    if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalizationDownloads"]))
                                        RLTLIB2.Log(string.Format("\tDownloaded '{0}' ({1})", localPath, RLTLIB2.FormatBytes(size)));
                                }
                                catch (Exception ex)
                                {
                                    RLTLIB2.LogError(string.Format("Downloading {0} for {1} in {2} - {3}",
                                        absoluteUrl, course.RvSku, fi, ex.Message));
                                }
                            }
                        }
                    }

                    text = text.Replace(url, newUrl);
                    if (bool.Parse(ConfigurationManager.AppSettings["LogUrlLocalization"]))
                        RLTLIB2.Log(string.Format("\tLocalized  '{0}' --> {1}", url, newUrl));

                    // Reposition end pointer
                    j = i + newUrl.Length + 1;
                }
            }

            foreach (string unknown in new List<string> {"http:", "https:"})
            {
                int k = 0;
                while (k < text.Length && text.IndexOf(unknown, k) > -1)
                {
                    int i = text.IndexOf(unknown, k);
                    string delimiter = text.Substring(i - 1, 1);
                    if (delimiter == "(") delimiter = ")"; // Handle CSS URLs
                    k = text.IndexOf(delimiter, i + 1);
                    string url = text.Substring(i, k - i);
                    url = CleanUrl(url);
                    k++;

                    if (!UrlContainsIgnoredFragment(url) && url != "http://" && url != "https://")
                        if (url.ToLower() != "http://www.redvector.com/support" &&
                            url.ToLower() != "https://www.redvector.com/support" &&
                            url.ToLower() != "http://www.redvector.com/support.aspx" &&
                            url.ToLower() != "https://www.redvector.com/support.aspx" &&
                            !url.ToLower().Contains("https://delivery.learnsmartsystems.com/aicc/distribution/includes/"))
                            RLTLIB2.LogWarning("Possible unlocalized URL - " + url);
                }
            }

            return text;
        }

        private static void DocumentBadUrl(Course course, FileInfo fi, string url)
        {
            RLTLIB2.LogWarning($"Course {course.RvSku}  - File '{fi.Name}' Bad URL '{url}'");
            bool writeHeader = !File.Exists("BadUrls.csv");
            StreamWriter sw = new StreamWriter("BadUrls.csv", true);
            if (writeHeader)
                sw.WriteLine("rv_sku,file_path,url_referenced");
            sw.WriteLine($"{course.RvSku},\"{fi.FullName}\",\"{url}\"");
            sw.Close();
        }

        private static bool UrlContainsIgnoredFragment(string url)
        {
            foreach (string frag in RvIgnoredUrlFragments)
                if (url.ToLower().Contains(frag))
                    return true;
            return false;
        }

        private static string CleanUrl(string url)
        {
            // Delete query string
            if (url.Contains("?"))
                url = url.Substring(0, url.IndexOf("?"));

            // Delete trailing '>' 1206\index.cfm
            if (url.Contains(">"))
                url = url.Substring(0, url.IndexOf(">"));

            // Delete trailing '&src' ST-0014AD\a001_course_title_screen.html
            if (url.Contains("&"))
                url = url.Substring(0, url.IndexOf("&"));

            // Trim trailing commas; e.g., ST-0004, a001_03_what_is_sexual_harassment_c03p05.html
            if (url.Contains(","))
                url = url.Substring(0, url.IndexOf(","));

            //Trim trailing HTML; e.g., ST-0004, a001_10_references_page_1.html
            if (url.Contains("<"))
                url = url.Substring(0, url.IndexOf("<"));

            // Catch usual errortype
            if (url.Length <= 3) throw new SystemException("url.Length <= 3");
            return url;
        }

        private static bool IsValidUrl(string url)
        {
            Regex mask = new Regex(@"^(ht|f)tp(s?)\:\/\/[0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*(:(0-9)*)*(\/?)([a-zA-Z0-9\-\.\?\,\'\/\\\+&%\$#_]*)?$",
                RegexOptions.IgnoreCase);
            return mask.IsMatch(url);
        }
    }
}