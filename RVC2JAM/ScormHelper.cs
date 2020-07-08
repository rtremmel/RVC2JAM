using System;
using System.Configuration;
using System.IO;
using System.Text;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
using VectorSolutions;

namespace RVC2JAM
{
    internal class ScormHelper
    {
        public static void CreateManifestFile(Course course)
        {
            RLTLIB2.Log("Generating SCORM manifest file");

            string xml = ManifestBeginXml(course);
            xml += $"    <resource identifier=\'RES\' type=\'webcontent\' href=\'{course.LaunchFileName}\' adlcp:scormtype=\'sco\'>\n";

            foreach (FileInfo f in new DirectoryInfo(course.WorkingDirectoryPath).GetFiles("*.*", SearchOption.AllDirectories))
            {
                if (f.FullName.Contains("\\.")) continue;
                if (f.FullName.Contains("-old.html")) continue;
                if (f.FullName.Contains("-tmp.html")) continue;
                if (".bak|.log|.shs|.zip".Contains(f.Extension)) continue;

                string asset = f.FullName.Replace(course.WorkingDirectoryPath, "");
                asset = asset.Replace("\\", "/");
                asset = asset.TrimStart('/');
                xml += "      <file href='" + RLTLIB2.ReplaceXmlSpecialCharacters(asset) + "'/>\n";
            }

            xml += ManifestEndXml();

            // Write SCORM manifest
            string manifestPath = Path.Combine(course.WorkingDirectoryPath, "imsmanifest.xml");
            if (File.Exists(manifestPath)) File.Delete(manifestPath);
            RLTLIB2.WriteTextFile(manifestPath, xml, Encoding.UTF8);
        }

        private static string ManifestBeginXml(Course course)
        {
            string xml = "";
            xml += "<?xml version='1.0' encoding='utf-8'?>\n";
            xml += $"<!-- Course {course.RvSku} - {EscapeXml(course.Title)} -->\n";
            xml += $"<!-- RedVector catalog_item_id {course.CatalogItemId} -->\n";
            xml += $"<!-- Converted by {RLTLIB2.AppNameAbbrVersion} on {DateTime.Now:g} -->\n";
            xml += "<manifest identifier='" + course.RvSku + "' version='1.0' \n";
            xml += "  xmlns='http://www.imsproject.org/xsd/imscp_rootv1p1p2'  \n";
            xml += "  xmlns:adlcp='http://www.adlnet.org/xsd/adlcp_rootv1p2'  \n";
            xml += "  xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'  \n";
            xml += "  xsi:schemaLocation='http://www.imsproject.org/xsd/imscp_rootv1p1p2  \n";
            xml += "    imscp_rootv1p1p2.xsd http://www.imsglobal.org/xsd/imsmd_rootv1p2p1  \n";
            xml += "    imsmd_rootv1p2p1.xsd http://www.adlnet.org/xsd/adlcp_rootv1p2  \n";
            xml += "    adlcp_rootv1p2.xsd'>\n";
            xml += "  <metadata>\n";
            xml += "    <schema>ADL SCORM</schema>\n";
            xml += "    <schemaversion>1.2</schemaversion>\n";
            xml += "    <lom xmlns='http://www.imsglobal.org/xsd/imsmd_rootv1p2p1'  \n";
            xml += "      xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'  \n";
            xml += "      xsi:schemaLocation='http://www.imsglobal.org/xsd/imsmd_rootv1p2p1 imsmd_rootv1p2p1.xsd'>\n";
            xml += "      <general>\n";
            xml += "        <title>\n";
            xml += "          <langstring xml:lang='x-none'>" + EscapeXml(course.Title) + "</langstring>\n";
            xml += "        </title>\n";
            xml += "      </general>\n";
            xml += "      <lifecycle>\n";
            xml += "        <version>\n";
            xml += "          <langstring xml:lang='x-none'>1.0</langstring>\n";
            xml += "        </version>\n";
            xml += "        <status>\n";
            xml += "          <source>\n";
            xml += "            <langstring xml:lang='x-none'>LOMv1.0</langstring>\n";
            xml += "          </source>\n";
            xml += "          <value>\n";
            xml += "            <langstring xml:lang='x-none'>Final</langstring>\n";
            xml += "          </value>\n";
            xml += "        </status>\n";
            xml += "      </lifecycle>\n";
            xml += "      <metametadata>\n";
            xml += "        <metadatascheme>ADL SCORM 1.2</metadatascheme>\n";
            xml += "      </metametadata>\n";
            xml += "    </lom>\n";
            xml += "  </metadata>\n";
            xml += "  <organizations default='ORG'>\n";
            xml += "    <organization identifier='ORG'>\n";
            xml += "      <title>" + EscapeXml(course.Title) + "</title>\n";
            xml += "      <item identifier='SCO' isvisible='true' identifierref='RES'>\n";
            xml += "        <title>" + EscapeXml(course.Title) + "</title>\n";
            xml += "        <adlcp:masteryscore>" + course.DefMasteryScore + "</adlcp:masteryscore>\n";
            xml += "      </item>\n";
            xml += "    </organization>\n";
            xml += "  </organizations>\n";
            xml += "  <resources>\n";
            return xml;
        }

        private static string ManifestEndXml()
        {
            string xml = "";
            xml += "    </resource>\n";
            xml += "  </resources>\n";
            xml += "</manifest>\n";
            return xml;
        }

        public static void CreateInfoFile(Course course)
        {
            string json = "";
            json += "{\n";
            json += $"\t\"rv_sku\":\"{course.RvSku}\",\n";
            json += $"\t\"title\":\"{course.Title}\",\n";
            json += $"\t\"catalog_item_id\":\"{course.CatalogItemId}\",\n";
            json += $"\t\"lesson_unit_id\":\"{course.LessonUnitId}\",\n";
            json += $"\t\"source\":\"{course.ProductionContentPath.Replace(@"\", @"\\")}\",\n";
            json += $"\t\"launch_url\":\"{course.LaunchUrl.Replace(@"\", @"\\")}\",\n";
            json += $"\t\"processed\":\"Converted by {RLTLIB2.AppNameAbbrVersion} on {DateTime.Now}\"\n";
            json += "}\n";
            string infoFile = Path.Combine(course.WorkingDirectoryPath, $"_{course.RvSku}_.json");
            if (File.Exists(infoFile)) File.Delete(infoFile);
            RLTLIB2.WriteTextFile(infoFile, json, Encoding.UTF8);
        }

        public static string EscapeXml(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            string returnString = s;
            returnString = returnString.Replace("'", "&apos;");
            returnString = returnString.Replace("\"", "&quot;");
            returnString = returnString.Replace(">", "&gt;");
            returnString = returnString.Replace("<", "&lt;");
            returnString = returnString.Replace("&", "&amp;");
            return returnString;
        }

        public static void CreateScormPackage(Course course)
        {
            RLTLIB2.Log("Generating SCORM package");

            // Create SCORM zip file
            string workingZipPath = CreateScormZipFile(course);
            string finalZipPath = workingZipPath.Replace(course.WorkingScormDirectoryPath, course.FinalScormDirectoryPath);

            // Copy SCORM file to final location
            if (!Directory.Exists(course.FinalScormDirectoryPath))
                Directory.CreateDirectory(course.FinalScormDirectoryPath);
            RLTLIB2.Log(string.Format("Copying SCORM archive to {0}", finalZipPath));
            File.Copy(workingZipPath, finalZipPath, true);
        }

        private static string CreateScormZipFile(Course course)
        {
            if (!Directory.Exists(course.WorkingScormDirectoryPath))
                Directory.CreateDirectory(course.WorkingScormDirectoryPath);

            string zipPath = Path.Combine(course.WorkingScormDirectoryPath, course.RvSku + ".zip");
            FileStream fsOut = File.Create(zipPath);
            ZipOutputStream zipStream = new ZipOutputStream(fsOut);

            int fileCount = 0;
            foreach (FileInfo f in new DirectoryInfo(course.WorkingDirectoryPath).GetFiles("*.*", SearchOption.AllDirectories))
            {
                if (f.FullName.Contains("\\.") ||
                    f.Name.Contains("_bak.") ||
                    f.Name.Contains("-old.") ||
                    f.Name.Contains("-tmp."))
                {
                    RLTLIB2.Log(string.Format("\tIgnoring file '{0}'", f.Name));
                    continue;
                }

                string entryName = f.FullName.Replace(course.WorkingDirectoryPath, "").TrimStart('\\');
                ZipEntry newEntry = new ZipEntry(entryName)
                {
                    DateTime = f.LastWriteTime,
                    Size = f.Length
                };

                zipStream.PutNextEntry(newEntry);
                byte[] buffer = new byte[4096];
                using (FileStream streamReader = File.OpenRead(f.FullName))
                {
                    StreamUtils.Copy(streamReader, zipStream, buffer);
                }

                zipStream.CloseEntry();

                fileCount++;
                if (bool.Parse(ConfigurationManager.AppSettings["LogScormFileDetails"]))
                    if (fileCount % 100 == 0)
                        RLTLIB2.Log($"\tAdded {RLTLIB2.Pluralize(fileCount, "files...")}");
            }

            zipStream.SetComment(string.Format("Created by {0} on {1:g}", RLTLIB2.AppNameAbbrVersion, DateTime.Now));
            zipStream.IsStreamOwner = true; // Makes the Close also Close the underlying stream
            zipStream.Close();

            RLTLIB2.Log(string.Format("Created SCORM archive {0} ({1})", zipPath, RLTLIB2.FormatBytes(new FileInfo(zipPath).Length)));
            return zipPath;
        }
    }
}