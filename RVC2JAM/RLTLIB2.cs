#define NLOG
//#define VERBOSE_SEARCHSTRING_DEBUGGING

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Style;

// ReSharper disable UnusedMember.Global
// ReSharper disable once CheckNamespace
namespace VectorSolutions
{
    internal class RLTLIB2
    {
        #region Help (about this class)

        [Category("Help")]
        [Description("RLTLIB2 Help")]
        public static string Help()
        {
            string help = "\nRick Tremmel's Common Method Library Version 2 - ";
            help += $"RLTLIB2 last modified {LastModifiedDate}\n";
            help += "(A collection of common properties and methods used by stand alone applications)\n";
            help += "Categories:\n";
            help += "\tHelp(about this class), Application, Console, DataTable, Database, \n";
            help += "\tDocumentation, Email, Excel, File/Folder, Format, Log, String, URL String Extensions. \n";
            help += "Requires:\n";
            help += "\tNuGet Package NLOG (optional).  If NLOG not defined, logging uses Debug.WriteLine and local file.\n";
            help += "\tNuGet Package EPPlus used to create Excel Workbooks.\n";
            help += "\tMicrosoft.ACE.OLEDB.12.0 installed: http://www.microsoft.com/en-us/download/details.aspx?id=13255\n";
            help += "\tReference to System.Configuration\n";
            help += "\tReference to System.Drawing\n";
            help += "\tConfigurationManager.AppSettings[\"RvEmailerServiceAttachmentPathExternal\"],\n";
            help += "\tConfigurationManager.AppSettings[\"RvEmailerServiceAttachmentPathInternal\"]);\n";
            help += "\tConfigurationManager.ConnectionStrings[\"RVLMSDB\"] (optional or ConnectionString property can be set dynamically)";
            return help;
        }

        #endregion Help (about this class)

        #region Console Helpers

        [Category("Console")]
        [Description("Display prompt on console.  If key is 'Y' or 'y' return true.  If timeout reached, return false.")]
        public static bool ConsoleYesNoPromptWithTimeout(string prompt, int timeoutSeconds)
        {
            // Example: if (ConsoleYesNoPromptWithTimeout("Test expired license (N) ?", 5))
            Console.WriteLine(prompt);
            string ch = "";
            DateTime beginWait = DateTime.Now;
            while (DateTime.Now.Subtract(beginWait).TotalSeconds < 5)
            {
                if (Console.KeyAvailable)
                {
                    ch = Console.ReadKey().KeyChar.ToString();
                    break;
                }

                Thread.Sleep(1000 * timeoutSeconds);
            }

            return ch == "y" || ch == "Y";
        }

        #endregion Console Helpers

        #region Email Helpers

        [Category("Email")]
        [Description("Insert email into RvEmailerService with optional attachment files.")]
        public static void InsertRvEmailerService(string fromEmail, string toEmails, string ccEmails, string bccEmails,
            string subjectText, bool isHTML, string bodyText, string attachmentPaths)
        {
            try
            {
                string attachments = "";
                if (!String.IsNullOrWhiteSpace(attachmentPaths))
                    foreach (string localFilePath in attachmentPaths.Split(';'))
                    {
                        string remoteFilePathInternal;
                        if (localFilePath.Trim().ToLower().StartsWith("http") || localFilePath.Trim().ToLower().StartsWith("ftp"))
                        {
                            // RvEmailService can handle URLs
                            remoteFilePathInternal = localFilePath;
                        }
                        else
                        {
                            // Else copy attachments to location RvEmailService can access
                            FileInfo f = new FileInfo(localFilePath);
                            string remoteFilePathExternal =
                                $"{ConfigurationManager.AppSettings["RvEmailerServiceAttachmentPathExternal"]}{Path.GetFileNameWithoutExtension(f.Name)}-{DateTime.Now:yyyyMMddHHmmss}{f.Extension}";
                            File.Copy(localFilePath, remoteFilePathExternal);
                            remoteFilePathInternal = remoteFilePathExternal.Replace(
                                ConfigurationManager.AppSettings["RvEmailerServiceAttachmentPathExternal"],
                                ConfigurationManager.AppSettings["RvEmailerServiceAttachmentPathInternal"]);
                        }

                        attachments += remoteFilePathInternal + ";";
                    }

                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = "dbo.insert_RvEmailerService";
                        cmd.Parameters.AddWithValue("@source_ref", AppNameAbbrVersion);
                        cmd.Parameters.AddWithValue("@from_email", fromEmail ?? "");
                        cmd.Parameters.AddWithValue("@to_emails", toEmails ?? "");
                        cmd.Parameters.AddWithValue("@cc_emails", ccEmails ?? "");
                        cmd.Parameters.AddWithValue("@bcc_emails", bccEmails ?? "");
                        cmd.Parameters.AddWithValue("@subject_text", subjectText ?? "");
                        cmd.Parameters.AddWithValue("@subject_content_id", null);
                        cmd.Parameters.AddWithValue("@subject_data", "");
                        cmd.Parameters.AddWithValue("@is_html", isHTML);
                        cmd.Parameters.AddWithValue("@is_priority", false);
                        cmd.Parameters.AddWithValue("@body_text", bodyText ?? "");
                        cmd.Parameters.AddWithValue("@body_content_id", null);
                        cmd.Parameters.AddWithValue("@body_path", "");
                        cmd.Parameters.AddWithValue("@body_data", "");
                        cmd.Parameters.AddWithValue("@attachment_paths", attachments.TrimEnd(';'));
                        cmd.ExecuteNonQuery();
                    }

                    conn.Close();
                }
            }
            catch (Exception e)
            {
                Log($"Unable to insert into 'rv_emailer_service' table. {e.Message} {e.StackTrace}");
            }
        }

        #endregion Email Helpers

        #region Application Helpers

        private static readonly Assembly Ea = Assembly.GetExecutingAssembly();
        private static readonly AssemblyName Ean = Ea.GetName();
        private static readonly Version Eanv = Ean.Version;
        private static readonly Attribute Ada = Attribute.GetCustomAttribute(Ea, typeof(AssemblyDescriptionAttribute));
        private static readonly Attribute Aca = Attribute.GetCustomAttribute(Ea, typeof(AssemblyCopyrightAttribute));

        [Category("Application")] [Description("The date RLTLIB2.cs was last updated.")]
        public static string LastModifiedDate = "07/01/2020";

        [Category("Application")] [Description("The value of the major component of the version number for the currently executing application.")]
        private static readonly int Major = Eanv.Major;

        [Category("Application")] [Description("The value of the minor component of the version number for the currently executing application.")]
        private static readonly int Minor = Eanv.Minor;

        [Category("Application")] [Description("The value of the build component of the version number for the currently executing application.")]
        private static readonly int Build = Eanv.Build;

        [Category("Application")] [Description("The value of the revision component of the version number for the currently executing application.")]
        private static readonly int Revision = Eanv.Revision;

        [Category("Application")] [Description("The description of the currently executing application.")]
        public static string AppName = ((AssemblyDescriptionAttribute) Ada).Description;

        [Category("Application")] [Description("The copyright of the currently executing application.")]
        public static string Copyright = ((AssemblyCopyrightAttribute) Aca).Copyright;

        [Category("Application")] [Description("The author of the currently executing application.")]
        public static string Author = "Rick Tremmel, 813-864-2656, Rick.Tremmel@VectorSolutions.com";

        [Category("Application")] [Description("The abbreviation of the currently executing application.")]
        public static string AppAbbr = Ean.Name;

        [Category("Application")] [Description("The name plus abbreviation of the currently executing application.")]
        public static string AppNameAbbr = $"{AppName} ({AppAbbr})";

        [Category("Application")] [Description("The name abbreviation plus version of the currently executing application.")]
        public static string AppAbbrVersion = $"{AppAbbr} v{Major}.{Minor}";

        [Category("Application")] [Description("The name plus abbreviation plus version of the currently executing application.")]
        public static string AppNameAbbrVersion = $"{AppName} ({AppAbbr} v{Major}.{Minor}.{Build})";

        [Category("Application")] [Description("The name plus abbreviation plus version plus build of the currently executing application.")]
        public static string AppNameAbbrVersionBuild = $"{AppName} ({AppAbbr} v{Major}.{Minor}.{Build} Build {Revision})";

        [Category("Application")] [Description("The verbose name plus abbreviation plus version plus build of the currently executing application.")]
        public static string AppNameAbbrVersionBuildLong = $"{AppName} ({AppAbbr}) Version {Major}.{Minor}.{Build} Build {Revision}";

        #endregion Application Helpers

        #region Database Helpers

        private static string _connectionString;

        [Category("Database")]
        [Description("Database connection string set or default loaded from app.config or web.config.")]
        public static string ConnectionString
        {
            get => _connectionString ?? ConfigurationManager.ConnectionStrings["RVLMSDB"].ToString();
            set => _connectionString = value;
        }


        [Category("Database")]
        [Description("Return basic information about connected database.")]
        public static DataTable DatabaseTest()
        {
            const string sql = "SELECT @@Servername AS [server], DB_NAME() AS [database], @@SPID AS [spid], USER AS [user], @@VERSION AS [version]";
            return ExecuteQuery(sql, out _);
        }

        [Category("Database")]
        [Description("Execute passed SQL query string returning DataTable.  Output parameter 'elapsedTime' returns execution TimeSpan.")]
        public static DataTable ExecuteQuery(string query, out TimeSpan elapsedTime)
        {
            try
            {
                DateTime startQuery = DateTime.Now;
                DataTable dt = new DataTable();
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandTimeout = 600;
                        SqlDataAdapter da = new SqlDataAdapter {SelectCommand = cmd};
                        da.Fill(dt);
                        elapsedTime = DateTime.Now - startQuery;
                        return dt;
                    }
                }
            }
            catch (Exception e)
            {
                string error = $"{e.Message} {e.StackTrace}";
                LogError(error);
                throw new SystemException(error);
            }
        }

        [Category("Database")]
        [Description(
            "Execute passed SQL stored procedure returning DataTable. " +
            "Parameters are passing as Dictionary object. " +
            "Output parameter 'elapsedTime' returns execution TimeSpan.")]
        public static DataTable ExecuteSproc(string sproc, Dictionary<string, string> parameters, out TimeSpan elapsedTime)
        {
            try
            {
                DateTime startQuery = DateTime.Now;
                DataTable dt = new DataTable();
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandTimeout = Convert.ToInt16(ConfigurationManager.AppSettings["SQLCommandTimeoutSeconds"]);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = sproc;
                        if (parameters != null)
                            foreach (KeyValuePair<string, string> p in parameters)
                                cmd.Parameters.AddWithValue(p.Key, p.Value);
                        SqlDataAdapter da = new SqlDataAdapter {SelectCommand = cmd};
                        da.Fill(dt);
                        elapsedTime = DateTime.Now - startQuery;
                        return dt;
                    }
                }
            }
            catch (Exception e)
            {
                string error = $"{e.Message} {e.StackTrace}";
                LogError(error);
                throw new SystemException(error);
            }
        }

        [Category("Database")]
        [Description(
            "Execute passed SQL stored procedure returning DataTable. " +
            "Parameters are passing as Dictionary object. " +
            "Output parameter 'elapsedTime' returns execution TimeSpan." +
            "OVERLOAD that allows <object> parameter value.")]
        public static DataTable ExecuteSprocO(string sproc, Dictionary<string, object> parameters, out TimeSpan elapsedTime)
        {
            try
            {
                DateTime startQuery = DateTime.Now;
                DataTable dt = new DataTable();
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandTimeout = Convert.ToInt16(ConfigurationManager.AppSettings["SQLCommandTimeoutSeconds"]);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandText = sproc;
                        if (parameters != null)
                            foreach (KeyValuePair<string, object> p in parameters)
                                cmd.Parameters.AddWithValue(p.Key, p.Value);
                        SqlDataAdapter da = new SqlDataAdapter { SelectCommand = cmd };
                        da.Fill(dt);
                        elapsedTime = DateTime.Now - startQuery;
                        return dt;
                    }
                }
            }
            catch (Exception e)
            {
                string error = $"{e.Message} {e.StackTrace}";
                LogError(error);
                throw new SystemException(error);
            }
        }

        #endregion Database Helpers

        #region DataTable Helpers

        [Category("DataTable")]
        [Description("Load passed Comma Separated Values (CSV) file into DataTable.")]
        public static DataTable LoadCsvFile(string fileDirectory, string filePath, out string errorMessage)
        {
            try
            {
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
                if (String.IsNullOrWhiteSpace(fileNameWithoutExtension))
                {
                    errorMessage = $"Invalid filename '{filePath}'";
                    return new DataTable();
                }

                if (fileNameWithoutExtension.IndexOf(".", StringComparison.Ordinal) > 0)
                {
                    errorMessage = $"Invalid filename '{filePath}' - Microsoft Jet OLEDB driver cannot read filenames containing multiple periods";
                    return new DataTable();
                }

                const string provider = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"{0}\";Extended Properties=\"text;HDR=Yes;FMT=Delimited(,)\"";

                using (OleDbConnection conn = new OleDbConnection(String.Format(provider, fileDirectory)))
                {
                    conn.Open();
                    using (OleDbDataAdapter da = new OleDbDataAdapter($"SELECT * FROM [{filePath}]", conn))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        errorMessage = "";
                        return dt;
                    }
                }
            }
            catch (Exception e)
            {
                string error = $"{e.Message}";
                LogError(error);
                errorMessage = error;
                return new DataTable();
            }
        }

        [Category("DataTable")]
        [Description("Load passed Tab Delimited text (TXT) file into DataTable.")]
        public static DataTable LoadTxtFile(string fileDirectory, string filePath, out string errorMessage)
        {
            try
            {
                DataTable dt = new DataTable();
                using (StreamReader sr = new StreamReader(Path.Combine(fileDirectory, filePath)))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        string[] items = line.Split('\t');
                        if (dt.Columns.Count == 0)
                            // Create the data columns for the data table based on the number of items on the first line of the file
                            foreach (string fld in items)
                                dt.Columns.Add(fld.ToLower().Contains("date")
                                    ? new DataColumn(fld, typeof(DateTime))
                                    : new DataColumn(fld, typeof(string)));
                        else
                            // Add row of data
                            // ReSharper disable once CoVariantArrayConversion
                            dt.Rows.Add(items);
                    }
                }

                errorMessage = "";
                return dt;
            }
            catch (Exception e)
            {
                string error = $"{e.Message}";
                LogError(error);
                errorMessage = error;
                return new DataTable();
            }
        }

        #endregion DataTable Helpers

        #region Excel Helpers

        [Category("Excel")]
        [Description(
            "Load passed Excel 'xls' or 'xlsx' Workbook into DataTable.  " +
            "The 'sheetName' is optional; if not passed, the first sheet is loaded.  " +
            "File extension 'xlsx' requires Microsoft.ACE.OLEDB.12.0 installed: http://www.microsoft.com/en-us/download/details.aspx?id=13255")]
        public static DataTable LoadExcelSheet(string filePath, string sheetName, out string errorMessage)
        {
            try
            {
                string provider = Path.GetExtension(filePath.ToLower()) == ".xlsx"
                    ? "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{0}\";Jet OLEDB:Engine Type=5;Extended Properties=\"Excel 12.0;HDR=1;IMEX=1\""
                    : "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\"{0}\";Extended Properties=\"Excel 12.0;HDR=1;IMEX=1\"";

                using (OleDbConnection conn = new OleDbConnection(String.Format(provider, filePath)))
                {
                    conn.Open();
                    // If sheetName is not specified, assume there is only one sheet
                    if (String.IsNullOrWhiteSpace(sheetName))
                    {
                        DataTable sheetTable = conn.GetSchema("Tables");

                        sheetName = "Sheet1$";
                        foreach (DataRow dr in sheetTable.Rows.Cast<DataRow>().Where(dr => dr["TABLE_NAME"].ToString().EndsWith("$")))
                            sheetName = dr["TABLE_NAME"].ToString();
                    }

                    // Remove any single quotes from sheetName
                    sheetName = sheetName.Replace("'", "");

                    // Append $ to end of sheet name for OlDb query
                    if (!sheetName.EndsWith("$"))
                        sheetName += "$";

                    using (OleDbDataAdapter sheetAdapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", conn))
                    {
                        DataTable sheetData = new DataTable();
                        sheetAdapter.Fill(sheetData);
                        errorMessage = "";
                        return sheetData;
                    }
                }
            }
            catch (Exception e)
            {
                string error = $"{e.Message}";
                LogError(error);
                errorMessage = error;
                return new DataTable();
            }
        }

        [Category("Excel")]
        [Description("Save passed DataTable in Excel file.")]
        public static void WriteExcel(DataTable dt, string source, string savePath)
        {
            try
            {
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(savePath)))
                {
                    // Set package properties
                    string description =
                        ((AssemblyDescriptionAttribute) Attribute.GetCustomAttribute(Ea, typeof(AssemblyDescriptionAttribute)))
                        .Description;
                    string version = $"{Ean.Name} v{Eanv.Major}.{Eanv.Minor}";
                    string title = $"{description} ({version})";

                    package.Workbook.Properties.Subject = description;
                    package.Workbook.Properties.Title = title;
                    package.Workbook.Properties.Author = "Rick Tremmel";
                    package.Workbook.Properties.Comments = $"Converted from '{source}'";

                    // Create the worksheet
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add("Sheet1");

                    // Load the datatable into the sheet, starting from cell A1. Print the column names on row 1
                    ws.Cells["A1"].LoadFromDataTable(dt, true);

                    // Format headers
                    using (ExcelRange range = ws.Cells[1, 1, 1, dt.Columns.Count])
                    {
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    }

                    ws.Cells[1, 1, dt.Rows.Count + 1, dt.Columns.Count].AutoFitColumns(0);

                    package.Save();
                }
            }
            catch (Exception e)
            {
                string error = $"{e.Message} {e.StackTrace}";
                LogError(error);
                throw new SystemException(error);
            }
        }

        [Category("Excel")]
        [Description("Update Excel worksheet using simple SQL condition (untested).")]
        public static void UpdateExcelSheet(string fileName, string sheetName, string column, string value, string where)
        {
            try
            {
                string provider = Path.GetExtension(fileName.ToLower()) == ".xlsx"
                    ? "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Jet OLEDB:Engine Type=5;Extended Properties=\"Excel 12.0;HDR=NO;IMEX=1\""
                    : "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Jet OLEDB:Engine Type=5;Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
                using (OleDbConnection conn = new OleDbConnection(String.Format(provider, fileName)))
                {
                    conn.Open();
                    // If sheetName is not specified, assume there is only one sheet
                    if (String.IsNullOrWhiteSpace(sheetName))
                    {
                        DataTable sheetTable = conn.GetSchema("Tables");
                        DataRow rowSheetName = sheetTable.Rows[0];
                        sheetName = rowSheetName["TABLE_NAME"].ToString();
                    }
                    else
                    {
                        sheetName += "$";
                    }

                    string sql = $"UPDATE [{sheetName}] SET {column}='{value}' WHERE {@where}";
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        Debug.WriteLine(sql);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception e)
            {
                string error = $"{e.Message} {e.StackTrace}";
                LogError(error);
                throw new SystemException(error);
            }
        }

        #endregion Excel Helpers

        #region Format Helpers

        [Category("Format")]
        [Description("Smart format long file or folder size as string.")]
        public static string FormatBytes(long size)
        {
            if (size < 0)
                return "Unknown";
            if (size < 1024)
                return size + " bytes";
            if (size < 1024 * 1024)
                return (int) (size / 1024) + " KB";
            if (size < 1024 * 1024 * 1024)
                return (int) (size / (1024 * 1024)) + " MB";
            if (size >= 1024 * 1024 * 1024)
                return (int) (size / (1024 * 1024 * 1024)) + " GB";
            return "FormatBytes Error";
        }

        [Category("Format")]
        [Description("Smart format elapsed time.")]
        public static string FormatElapsedTime(TimeSpan ts)
        {
            string days = ts.Days > 0 ? $"{ts.Days:n0} day{(ts.Days == 1 ? "" : "s")}, " : String.Empty;
            string hours = ts.Hours > 0 ? $"{ts.Hours:n0} hour{(ts.Hours == 1 ? "" : "s")}, " : String.Empty;
            string minutes = ts.Minutes > 0 ? $"{ts.Minutes:n0} minute{(ts.Minutes == 1 ? "" : "s")}, " : String.Empty;
            string seconds = ts.Seconds > 0 ? $"{ts.Seconds:n0} second{(ts.Seconds == 1 ? "" : "s")}, " : String.Empty;
            string milliseconds = ts.Milliseconds > 0
                ? $"{ts.Milliseconds:n0} millisecond{(ts.Milliseconds == 1 ? "" : "s")}, "
                : String.Empty;
            string s = $"{days}{hours}{minutes}{seconds}{milliseconds}";
            if (s.EndsWith(", "))
                s = s.Substring(0, s.Length - 2);
            if (s == "")
                s = "0 milliseconds";
            return s;
        }

        #endregion Format Helpers

        #region File and Folder Helpers

        [Category("File/Folder")]
        [Description("Read text file.")]
        public static string ReadTextFile(string filePath, Encoding encoding = null)
        {
            string text = null;
            try
            {
                using (StreamReader sr = new StreamReader(filePath, encoding ?? Encoding.Default))
                {
                    text = sr.ReadToEnd();
                    sr.Close();
                }
            }
            catch (Exception e)
            {
                LogError($"Reading {filePath} - {e.Message}");
            }

            return text;
        }

        [Category("File/Folder")]
        [Description("Write text file.")]
        public static string WriteTextFile(string filePath, string text, Encoding encoding = null)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(filePath, false, encoding ?? Encoding.Default))
                {
                    sw.Write(text);
                    sw.Close();
                }
            }
            catch (Exception e)
            {
                LogError($"Writing {filePath} - {e.Message}");
            }

            return text;
        }

        [Category("File/Folder")]
        [Description("Reset all file attributes (to not read-only), delete folder and all contents recursively, then recreate folder.")]
        public static void CreateFolderWithDelete(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                foreach (FileInfo file in new DirectoryInfo(folderPath).GetFiles("*.*", SearchOption.AllDirectories))
                    file.Attributes = FileAttributes.Normal;

                Directory.Delete(folderPath, true);
            }

            Directory.CreateDirectory(folderPath);
        }

        private static bool FitsOneOfMultipleMasks(string fileName, string fileMasks)
        {
            return fileMasks.Split(new[] {"\r\n", "\n", ";", ",", "|", " "}, StringSplitOptions.RemoveEmptyEntries)
                .Any(fileMask => FitsMask(fileName, fileMask));
        }

        private static bool FitsMask(string fileName, string fileMask)
        {
            Regex mask = new Regex(
                '^' +
                fileMask
                    .Replace(".", "[.]")
                    .Replace("*", ".*")
                    .Replace("?", ".")
                + '$',
                RegexOptions.IgnoreCase);
            return mask.IsMatch(fileName);
        }

        [Category("File/Folder")]
        [Description(
            "CopyFolderTree logging level Enum. None = No logging, Summary=Only log summary, BigFiles=Log files > 10MB, Thousand=Log every 1000th file copied, Hundred=Log every 100th file copied, IgnoresExcludes=Log ignored and excluded files, All=Log all files copied, Verbose=Log folders created."
        )]
        public enum CopyFolderTreeLoggingLevel
        {
            None, // No logging
            Summary, // Only log summary
            BigFiles, // Log files > 10MB
            Thousand, // Log every 1000th file copied
            Hundred, // Log every 100th file copied
            IgnoresExcludes, // Log ignored and excluded files
            All, // Log all files copied
            Verbose // Log all files copied and folders created
        }

        [Category("File/Folder")]
        [Description("Copy source folder to target folder with optional logging.")]
        public static void CopyFolderTree(string sourceFolder, string targetFolder, string includeMasks, string excludeMasks, CopyFolderTreeLoggingLevel ll)
        {
            DateTime start = DateTime.Now;
            int countDirectories = 0;
            int countFiles = 0;

            // Create target directory if needed
            if (!Directory.Exists(targetFolder))
            {
                if (ll >= CopyFolderTreeLoggingLevel.Verbose)
                    Log($"\tCreating '{targetFolder}'");
                Directory.CreateDirectory(targetFolder);
                countDirectories++;
            }

            // Create target subdirectories if needed
            foreach (string dirPath in Directory.GetDirectories(sourceFolder, "*", SearchOption.AllDirectories))
            {
                if (ll >= CopyFolderTreeLoggingLevel.Verbose)
                    Log($"\tCreating '{dirPath.Replace(sourceFolder, targetFolder)}'");
                Directory.CreateDirectory(dirPath.Replace(sourceFolder, targetFolder));
                countDirectories++;
            }

            // Copy all the files and replace any files with the same name
            foreach (string sourceFilePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
            {
                if (!FitsOneOfMultipleMasks(Path.GetFileName(sourceFilePath), includeMasks))
                {
                    if (ll >= CopyFolderTreeLoggingLevel.IgnoresExcludes)
                        Log($"\tIgnoring '{sourceFilePath}'");
                    continue;
                }

                if (FitsOneOfMultipleMasks(Path.GetFileName(sourceFilePath), excludeMasks))
                {
                    if (ll >= CopyFolderTreeLoggingLevel.IgnoresExcludes)
                        Log($"\tExcluding '{sourceFilePath}'");
                    continue;
                }

                if (ll > CopyFolderTreeLoggingLevel.All)
                {
                    Log($"\tCopying '{sourceFilePath}' --> '{sourceFilePath.Replace(sourceFolder, targetFolder)}'");
                }
                else if (ll >= CopyFolderTreeLoggingLevel.BigFiles)
                {
                    long size = new FileInfo(sourceFilePath).Length;
                    if (size > 10485760)
                        Log($"\tCopying '{sourceFilePath}' ({FormatBytes(size)})");
                }

                File.Copy(sourceFilePath, sourceFilePath.Replace(sourceFolder, targetFolder), true);
                countFiles++;

                if (ll >= CopyFolderTreeLoggingLevel.Thousand && countFiles % 1000 == 0)
                    Log($"\tCopied {countFiles:n0} files...");
                else if (ll >= CopyFolderTreeLoggingLevel.Hundred && countFiles % 100 == 0)
                    Log($"\tCopied {countFiles:n0} files...");
            }

            if (ll >= CopyFolderTreeLoggingLevel.Summary)
                Log(
                    $"\tCopied {countFiles:n0} files and created {countDirectories:n0} directories in {FormatElapsedTime(DateTime.Now - start)} ({FormatBytes(DirSize(new DirectoryInfo(targetFolder)))})");
        }

        [Category("File/Folder")]
        [Description("Reset attributes on directory tree.")]
        public static void ResetFolderAttributes(string folderPath)
        {
            foreach (string filePath in Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories))
                ResetFileAttributes(filePath);
        }

        [Category("File/Folder")]
        [Description("Reset attributes on file.")]
        public static void ResetFileAttributes(string filePath)
        {
            if (File.Exists(filePath))
                File.SetAttributes(filePath, FileAttributes.Normal);
        }

        [Category("File/Folder")]
        [Description("Determine total size of directory tree.")]
        public static long DirSize(DirectoryInfo d)
        {
            try
            {
                // Add file sizes
                FileInfo[] fis = d.GetFiles();
                long size = fis.Sum(fi => fi.Length);
                // Add subdirectory sizes
                DirectoryInfo[] dis = d.GetDirectories();
                // ReSharper disable once ConvertClosureToMethodGroup
                size += dis.Sum(di => DirSize(di));
                return size;
            }
            catch (Exception)
            {
                return -1;
            }
        }

        [Category("File/Folder")]
        [Description("Determine newest file LastWriteTime in directory tree.")]
        public static DateTime DirMaxDate(DirectoryInfo d)
        {
            DateTime maxDate = DateTime.MinValue;
            try
            {
                FileInfo[]
                    fis = d.GetFiles();
                maxDate = new[]
                    {maxDate, fis.Max(fi => fi.LastWriteTime)}.Max();

                foreach (DirectoryInfo di in d.GetDirectories())
                    maxDate = new[]
                        {maxDate, DirMaxDate(di)}.Max();

                return maxDate;
            }
            catch (Exception)
            {
                return DateTime.MinValue;
            }
        }


        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetDiskFreeSpaceEx(string lpDirectoryName,
            out ulong lpFreeBytesAvailable,
            out ulong lpTotalNumberOfBytes,
            out ulong lpTotalNumberOfFreeBytes);


        [Category("File/Folder")]
        [Description("Get UNC path free disk space using Win32 API.  Returns false if error.")]
        public static string GetDiskFreeSpaceInfo(string directoryName)
        {
            // ReSharper disable once NotAccessedVariable
            ulong freeBytesAvailable;
            // ReSharper disable once InlineOutVariableDeclaration
            ulong lpTotalNumberOfBytes;
            // ReSharper disable once InlineOutVariableDeclaration
            ulong lpTotalNumberOfFreeBytes;
            bool success = GetDiskFreeSpaceEx(directoryName, out freeBytesAvailable, out lpTotalNumberOfBytes, out lpTotalNumberOfFreeBytes);
            return !success
                ? null
                : $"{(long) (lpTotalNumberOfFreeBytes * 100 / lpTotalNumberOfBytes):n1}% Free ({FormatBytes((long) lpTotalNumberOfFreeBytes)} of {FormatBytes((long) lpTotalNumberOfBytes)})";
        }

        #endregion File and Folder Helpers

        #region Log Helpers

        private static string _logFilePath;

        [Category("Log")]
        [Description("Get/Set log file path.  If LogFilePath is null, return appname.log")]
        public static string LogFilePath
        {
            get => _logFilePath ?? (_logFilePath = $"{Ean.Name}.log");
            set => _logFilePath = value;
        }

        [Category("Log")]
        [Description("Return current contents of log file.")]
        public static string GetLogFileContents()
        {
            string text;
            using (StreamReader sr = new StreamReader(LogFilePath))
            {
                text = sr.ReadToEnd();
                sr.Close();
            }

            return text;
        }

        [Category("Log")]
        [Description("Append repeated char to log file.")]
        public static void LogRepeatedChar(char chr, int len)
        {
            Log(new string(chr, len));
        }

        [Category("Log")]
        [Description("Append line to log file.  If line is null, delete any existing log file.")]
        public static void Log(string msg, LogLevel level = null)
        {
#if (NLOG)
            Logger log = LogManager.GetCurrentClassLogger();
            log.Log(level ?? LogLevel.Info, msg);
#else
	        if (msg == null)
	        {
	            if (File.Exists(LogFilePath))
	                File.Delete(LogFilePath);
	            return;
	        }

	        //msg = $"{DateTime.Now:g} - {msg}";
	        Console.WriteLine(msg); //Command Prompt display
	        Debug.WriteLine(msg); //Visual Studio Output Window display
	        StreamWriter sr = new StreamWriter(LogFilePath, true);
	        sr.WriteLine(msg); //Log file output
	        sr.Close();
#endif
        }

        [Category("Log")] [Description("Public list used to document warnings.")]
        public static List<string> LogWarnings = new List<string>();

        [Category("Log")]
        [Description("Add line to LogWarnings then append line to log file.")]
        public static void LogWarning(string msg, bool log = true)
        {
            msg = "WARNING: " + msg;
            LogWarnings.Add(msg);
            if (log) Log(msg);
        }

        [Category("Log")] [Description("Public list used to document errors.")]
        public static List<string> LogErrors = new List<string>();

        [Category("Log")]
        [Description("Add line to LogErrors then append line to log file.")]
        public static void LogError(string msg, bool log = true)
        {
            msg = "ERROR: " + msg;
            LogErrors.Add(msg);
            if (log) Log(msg);
        }

        [Category("Log")]
        [Description("Smart conversion of log file string to HTML string for emailing.")]
        public static string Log2Html(string log, bool bodyOnly = false)
        {
            log = log.Replace("<", "&lt;");
            log = log.Replace(">", "&gt;");
            log = log.Replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;");
            log = log.Replace("  ", "&nbsp;&nbsp;");
            log = log.Replace("ERROR:", "<span style='color:red;font-weight:bold;'>ERROR:</span>");
            log = log.Replace("WARNING:", "<span style='color:magenta;font-weight:bold;'>WARNING:</span>");
            log = log.Replace("SUCCESS:", "<span style='color:darkgreen;font-weight:bold;'>SUCCESS:</span>");
            log = log.Replace(Environment.NewLine, "\n<br>");

            //while (log.Contains("'"))
            //{
            //    Int32 i = log.IndexOf("'");
            //    Int32 j = log.IndexOf("'", i + 1);
            //    if (j > 0)
            //    {
            //        string quoteString = log.Substring(i, j - i + 1);
            //        string quoted = log.Substring(i + 1, j - i - 1);
            //        log = log.Replace(quoteString, $"<strong>{quoted}</strong>");
            //    }
            //    else
            //    {
            //        break;
            //    }
            //}

            return bodyOnly
                ? log
                : $"<html><body style='font-family:sans-serif;'>{log}</body></html>";
        }

        #endregion Log Helpers

        #region String Helpers

        [Category("String")]
        [Description("Simple Pluralize.  Returns 'val' formatted as ':n0' plus pluralized 'sng' string")]
        public static string Pluralize(int val, string sng)
        {
            return Pluralize(val, sng, sng + "s");
        }

        [Description("Custom Pluralize.  Returns 'val' formatted as ':n0' plus 'sng' if singular otherwise 'plu' string")]
        public static string Pluralize(int val, string sng, string plu)
        {
            return $"{val:n0} {(val == 1 ? sng : plu)}";
        }
        
        [Category("String")]
        [Description(
            "Search string for text that 'beginsWith' and 'endsWith' and contains 'contains' but excludes 'excludes'.  Returns list of matched strings.  Include 'beginsWith' and 'endsWith' in returned strings if 'includeBeginsEnds' is true"
        )]
        public static List<string> SearchString(string text, string beginsWith, string endsWith, string contains, string excludes,
            bool ignoreCase, bool includeBeginsEnds = true)
        {
            List<string> matches = new List<string>();

            string debugInfo = "";

            if (String.IsNullOrWhiteSpace(text))
                return matches;

            string txt = ignoreCase ? text.ToLower() : text;
            string beg = ignoreCase ? beginsWith.ToLower() : beginsWith;
            string eds = ignoreCase ? endsWith.ToLower() : endsWith;
            string con = ignoreCase ? contains.ToLower() : contains;
            string exc = ignoreCase ? excludes.ToLower() : excludes;
            if (String.IsNullOrWhiteSpace(exc))
                exc = "DO-NOT-FIND-ME";

            debugInfo += "\n----- Begin SearchString() Debugging -----\n";
            //debugInfo += $"txt={txt}\n";
            debugInfo += $"beg={beg}\n";
            debugInfo += $"eds={eds}\n";
            debugInfo += $"con={con}\n";
            debugInfo += $"exc={exc}\n";
            debugInfo += $"ignoreCase={ignoreCase}\n";
            debugInfo += $"includeBeginsEnds={includeBeginsEnds}\n";

            int i = 0;
            int j = txt.IndexOf(eds, i, StringComparison.Ordinal);
            while (i < txt.Length && txt.IndexOf(beg, i, StringComparison.Ordinal) >= 0)
            {
                i = txt.IndexOf(beg, i, StringComparison.Ordinal);
                if (!includeBeginsEnds)
                    i += beg.Length;

                if (j == -1)
                    break;
                if (includeBeginsEnds)
                    j += eds.Length;
                string s = txt.Substring(i, j - i);
                debugInfo += $"i={i}\n";
                debugInfo += $"j={j}\n";
                debugInfo += $"s={s}\n";

                if (!s.Contains(con) || s.Contains(exc))
                {
                    i = j + 1;
                    continue;
                }

                string m = text.Substring(i, j - i);
                debugInfo += $"m={m}\n";

                matches.Add(m);
                i = j + 1;
            }

#if (VERBOSE_SEARCHSTRING_DEBUGGING)
            if (matches.Count > 0)
            {
                debugInfo += "----- End SearchString() Debugging -----\n\n";
                Log(debugInfo, LogLevel.Trace);
            }
#endif
            return matches;
        }

        [Category("String")]
        [Description("Trim string if longer than maxChars and append ellipsis.")]
        public static string TruncateString(string value, int maxChars)
        {
            // Strip non-ASCII characters
            Regex reg_exp = new Regex("[^ -~]+");
            value = reg_exp.Replace(value, "");
            return value.Length <= maxChars ? value : value.Substring(0, maxChars) + "...";
        }

        [Category("String")]
        [Description("Strip HTML from string and convert HTML entities to non-special characters.  Used to convert course descriptions")]
        public static string StripHtml(string s)
        {
            // ReSharper disable CommentTypo
            // ReSharper disable StringLiteralTypo
            const string htmlTagPattern = "<.*?>";
            s = Regex.Replace(s, htmlTagPattern, String.Empty); //Strip HTML
            s = s.Replace("\n", " "); //Strip CrLf
            s = s.Replace("\r", " "); //Strip Cr
            //
            // The following code was generated by Entities.xlsx
            // Based on http://htmlhelp.com/reference/html40/entities/e
            //
            s = s.Replace("" + (char) 34, "\"");
            s = s.Replace("&#34;", "\""); //quotation mark = APL quote (&quot;)
            s = s.Replace("" + (char) 38, "&");
            s = s.Replace("&#38;", "&"); //ampersand (&amp;)
            s = s.Replace("" + (char) 60, "<");
            s = s.Replace("&#60;", "<"); //less-than sign (&lt;)
            s = s.Replace("" + (char) 62, ">");
            s = s.Replace("&#62;", ">"); //greater-than sign (&gt;)
            s = s.Replace("" + (char) 160, " ");
            s = s.Replace("&#160;", " "); //no-break space = non-breaking space (&nbsp;)
            s = s.Replace("" + (char) 161, "!");
            s = s.Replace("&#161;", "!"); //inverted exclamation mark (&iexcl;)
            s = s.Replace("" + (char) 162, "cents");
            s = s.Replace("&#162;", "cents"); //cent sign (&cent;)
            s = s.Replace("" + (char) 163, "lbs");
            s = s.Replace("&#163;", "lbs"); //pound sign (&pound;)
            s = s.Replace("" + (char) 164, "(currency)");
            s = s.Replace("&#164;", "(currency)"); //currency sign (&curren;)
            s = s.Replace("" + (char) 165, "yen");
            s = s.Replace("&#165;", "yen"); //yen sign = yuan sign (&yen;)
            s = s.Replace("" + (char) 166, "¦");
            s = s.Replace("&#166;", "¦"); //broken bar = broken vertical bar (&brvbar;)
            s = s.Replace("" + (char) 167, "(section)");
            s = s.Replace("&#167;", "(section)"); //section sign (&sect;)
            s = s.Replace("" + (char) 168, "..");
            s = s.Replace("&#168;", ".."); //diaeresis = spacing diaeresis (&uml;)
            s = s.Replace("" + (char) 169, "(c)");
            s = s.Replace("&#169;", "(c)"); //copyright sign (&copy;)
            s = s.Replace("" + (char) 170, "(a)");
            s = s.Replace("&#170;", "(a)"); //feminine ordinal indicator (&ordf;)
            s = s.Replace("" + (char) 171, "<<");
            s = s.Replace("&#171;", "<<");
            //left-pointing double angle quotation mark = left pointing guillemet (&laquo;)
            s = s.Replace("" + (char) 172, "!=");
            s = s.Replace("&#172;", "!="); //not sign (&not;)
            s = s.Replace("" + (char) 173, "-");
            s = s.Replace("&#173;", "-"); //soft hyphen = discretionary hyphen (&shy;)
            s = s.Replace("" + (char) 174, "(r)");
            s = s.Replace("&#174;", "(r)"); //registered sign = registered trade mark sign (&reg;)
            s = s.Replace("" + (char) 175, "-");
            s = s.Replace("&#175;", "-"); //macron = spacing macron = overline = APL overbar (&macr;)
            s = s.Replace("" + (char) 176, "degrees");
            s = s.Replace("&#176;", "degrees"); //degree sign (&deg;)
            s = s.Replace("" + (char) 177, "+/-");
            s = s.Replace("&#177;", "+/-"); //plus-minus sign = plus-or-minus sign (&plusmn;)
            s = s.Replace("" + (char) 178, "-2");
            s = s.Replace("&#178;", "-2"); //superscript two = superscript digit two = squared (&sup2;)
            s = s.Replace("" + (char) 179, "-3");
            s = s.Replace("&#179;", "-3"); //superscript three = superscript digit three = cubed (&sup3;)
            s = s.Replace("" + (char) 180, "");
            s = s.Replace("&#180;", ""); //acute accent = spacing acute (&acute;)
            s = s.Replace("" + (char) 181, "u");
            s = s.Replace("&#181;", "u"); //micro sign (&micro;)
            s = s.Replace("" + (char) 182, "P");
            s = s.Replace("&#182;", "P"); //pilcrow sign = paragraph sign (&para;)
            s = s.Replace("" + (char) 183, "*");
            s = s.Replace("&#183;", "*"); //middle dot = Georgian comma = Greek middle dot (&middot;)
            s = s.Replace("" + (char) 184, ",");
            s = s.Replace("&#184;", ","); //cedilla = spacing cedilla (&cedil;)
            s = s.Replace("" + (char) 185, "-1");
            s = s.Replace("&#185;", "-1"); //superscript one = superscript digit one (&sup1;)
            s = s.Replace("" + (char) 186, "(o)");
            s = s.Replace("&#186;", "(o)"); //masculine ordinal indicator (&ordm;)
            s = s.Replace("" + (char) 187, ">>");
            s = s.Replace("&#187;", ">>");
            //right-pointing double angle quotation mark = right pointing guillemet (&raquo;)
            s = s.Replace("" + (char) 188, "1/4");
            s = s.Replace("&#188;", "1/4"); //vulgar fraction one quarter = fraction one quarter (&frac14;)
            s = s.Replace("" + (char) 189, "1/2");
            s = s.Replace("&#189;", "1/2"); //vulgar fraction one half = fraction one half (&frac12;)
            s = s.Replace("" + (char) 190, "3/4");
            s = s.Replace("&#190;", "3/4"); //vulgar fraction three quarters = fraction three quarters (&frac34;)
            s = s.Replace("" + (char) 191, "?");
            s = s.Replace("&#191;", "?"); //inverted question mark = turned question mark (&iquest;)
            s = s.Replace("" + (char) 192, "A");
            s = s.Replace("&#192;", "A"); //Latin capital letter A with grave = Latin capital letter A grave (&Agrave;)
            s = s.Replace("" + (char) 193, "A");
            s = s.Replace("&#193;", "A"); //Latin capital letter A with acute (&Aacute;)
            s = s.Replace("" + (char) 194, "A");
            s = s.Replace("&#194;", "A"); //Latin capital letter A with circumflex (&Acirc;)
            s = s.Replace("" + (char) 195, "A");
            s = s.Replace("&#195;", "A"); //Latin capital letter A with tilde (&Atilde;)
            s = s.Replace("" + (char) 196, "A");
            s = s.Replace("&#196;", "A"); //Latin capital letter A with diaeresis (&Auml;)
            s = s.Replace("" + (char) 197, "A");
            s = s.Replace("&#197;", "A");
            //Latin capital letter A with ring above = Latin capital letter A ring (&Aring;)
            s = s.Replace("" + (char) 198, "AE");
            s = s.Replace("&#198;", "AE"); //Latin capital letter AE = Latin capital ligature AE (&AElig;)
            s = s.Replace("" + (char) 199, "C");
            s = s.Replace("&#199;", "C"); //Latin capital letter C with cedilla (&Ccedil;)
            s = s.Replace("" + (char) 200, "E");
            s = s.Replace("&#200;", "E"); //Latin capital letter E with grave (&Egrave;)
            s = s.Replace("" + (char) 201, "E");
            s = s.Replace("&#201;", "E"); //Latin capital letter E with acute (&Eacute;)
            s = s.Replace("" + (char) 202, "E");
            s = s.Replace("&#202;", "E"); //Latin capital letter E with circumflex (&Ecirc;)
            s = s.Replace("" + (char) 203, "E");
            s = s.Replace("&#203;", "E"); //Latin capital letter E with diaeresis (&Euml;)
            s = s.Replace("" + (char) 204, "I");
            s = s.Replace("&#204;", "I"); //Latin capital letter I with grave (&Igrave;)
            s = s.Replace("" + (char) 205, "I");
            s = s.Replace("&#205;", "I"); //Latin capital letter I with acute (&Iacute;)
            s = s.Replace("" + (char) 206, "I");
            s = s.Replace("&#206;", "I"); //Latin capital letter I with circumflex (&Icirc;)
            s = s.Replace("" + (char) 207, "I");
            s = s.Replace("&#207;", "I"); //Latin capital letter I with diaeresis (&Iuml;)
            s = s.Replace("" + (char) 208, "D");
            s = s.Replace("&#208;", "D"); //Latin capital letter ETH (&ETH;)
            s = s.Replace("" + (char) 209, "N");
            s = s.Replace("&#209;", "N"); //Latin capital letter N with tilde (&Ntilde;)
            s = s.Replace("" + (char) 210, "O");
            s = s.Replace("&#210;", "O"); //Latin capital letter O with grave (&Ograve;)
            s = s.Replace("" + (char) 211, "O");
            s = s.Replace("&#211;", "O"); //Latin capital letter O with acute (&Oacute;)
            s = s.Replace("" + (char) 212, "O");
            s = s.Replace("&#212;", "O"); //Latin capital letter O with circumflex (&Ocirc;)
            s = s.Replace("" + (char) 213, "O");
            s = s.Replace("&#213;", "O"); //Latin capital letter O with tilde (&Otilde;)
            s = s.Replace("" + (char) 214, "O");
            s = s.Replace("&#214;", "O"); //Latin capital letter O with diaeresis (&Ouml;)
            s = s.Replace("" + (char) 215, "x");
            s = s.Replace("&#215;", "x"); //multiplication sign (&times;)
            s = s.Replace("" + (char) 216, "O");
            s = s.Replace("&#216;", "O"); //Latin capital letter O with stroke = Latin capital letter O slash (&Oslash;)
            s = s.Replace("" + (char) 217, "U");
            s = s.Replace("&#217;", "U"); //Latin capital letter U with grave (&Ugrave;)
            s = s.Replace("" + (char) 218, "U");
            s = s.Replace("&#218;", "U"); //Latin capital letter U with acute (&Uacute;)
            s = s.Replace("" + (char) 219, "U");
            s = s.Replace("&#219;", "U"); //Latin capital letter U with circumflex (&Ucirc;)
            s = s.Replace("" + (char) 220, "U");
            s = s.Replace("&#220;", "U"); //Latin capital letter U with diaeresis (&Uuml;)
            s = s.Replace("" + (char) 221, "Y");
            s = s.Replace("&#221;", "Y"); //Latin capital letter Y with acute (&Yacute;)
            s = s.Replace("" + (char) 222, "b");
            s = s.Replace("&#222;", "b"); //Latin capital letter THORN (&THORN;)
            s = s.Replace("" + (char) 223, "B");
            s = s.Replace("&#223;", "B"); //Latin small letter sharp s = ess-zed (&szlig;)
            s = s.Replace("" + (char) 224, "a");
            s = s.Replace("&#224;", "a"); //Latin small letter a with grave = Latin small letter a grave (&agrave;)
            s = s.Replace("" + (char) 225, "a");
            s = s.Replace("&#225;", "a"); //Latin small letter a with acute (&aacute;)
            s = s.Replace("" + (char) 226, "a");
            s = s.Replace("&#226;", "a"); //Latin small letter a with circumflex (&acirc;)
            s = s.Replace("" + (char) 227, "a");
            s = s.Replace("&#227;", "a"); //Latin small letter a with tilde (&atilde;)
            s = s.Replace("" + (char) 228, "a");
            s = s.Replace("&#228;", "a"); //Latin small letter a with diaeresis (&auml;)
            s = s.Replace("" + (char) 229, "a");
            s = s.Replace("&#229;", "a"); //Latin small letter a with ring above = Latin small letter a ring (&aring;)
            s = s.Replace("" + (char) 230, "ae");
            s = s.Replace("&#230;", "ae"); //Latin small letter ae = Latin small ligature ae (&aelig;)
            s = s.Replace("" + (char) 231, "c");
            s = s.Replace("&#231;", "c"); //Latin small letter c with cedilla (&ccedil;)
            s = s.Replace("" + (char) 232, "e");
            s = s.Replace("&#232;", "e"); //Latin small letter e with grave (&egrave;)
            s = s.Replace("" + (char) 233, "e");
            s = s.Replace("&#233;", "e"); //Latin small letter e with acute (&eacute;)
            s = s.Replace("" + (char) 234, "e");
            s = s.Replace("&#234;", "e"); //Latin small letter e with circumflex (&ecirc;)
            s = s.Replace("" + (char) 235, "e");
            s = s.Replace("&#235;", "e"); //Latin small letter e with diaeresis (&euml;)
            s = s.Replace("" + (char) 236, "I");
            s = s.Replace("&#236;", "I"); //Latin small letter i with grave (&igrave;)
            s = s.Replace("" + (char) 237, "I");
            s = s.Replace("&#237;", "I"); //Latin small letter i with acute (&iacute;)
            s = s.Replace("" + (char) 238, "I");
            s = s.Replace("&#238;", "I"); //Latin small letter i with circumflex (&icirc;)
            s = s.Replace("" + (char) 239, "I");
            s = s.Replace("&#239;", "I"); //Latin small letter i with diaeresis (&iuml;)
            s = s.Replace("" + (char) 240, "o");
            s = s.Replace("&#240;", "o"); //Latin small letter eth (&eth;)
            s = s.Replace("" + (char) 241, "n");
            s = s.Replace("&#241;", "n"); //Latin small letter n with tilde (&ntilde;)
            s = s.Replace("" + (char) 242, "o");
            s = s.Replace("&#242;", "o"); //Latin small letter o with grave (&ograve;)
            s = s.Replace("" + (char) 243, "o");
            s = s.Replace("&#243;", "o"); //Latin small letter o with acute (&oacute;)
            s = s.Replace("" + (char) 244, "o");
            s = s.Replace("&#244;", "o"); //Latin small letter o with circumflex (&ocirc;)
            s = s.Replace("" + (char) 245, "o");
            s = s.Replace("&#245;", "o"); //Latin small letter o with tilde (&otilde;)
            s = s.Replace("" + (char) 246, "o");
            s = s.Replace("&#246;", "o"); //Latin small letter o with diaeresis (&ouml;)
            s = s.Replace("" + (char) 247, "/");
            s = s.Replace("&#247;", "/"); //division sign (&divide;)
            s = s.Replace("" + (char) 248, "o");
            s = s.Replace("&#248;", "o"); //Latin small letter o with stroke = Latin small letter o slash (&oslash;)
            s = s.Replace("" + (char) 249, "u");
            s = s.Replace("&#249;", "u"); //Latin small letter u with grave (&ugrave;)
            s = s.Replace("" + (char) 250, "u");
            s = s.Replace("&#250;", "u"); //Latin small letter u with acute (&uacute;)
            s = s.Replace("" + (char) 251, "u");
            s = s.Replace("&#251;", "u"); //Latin small letter u with circumflex (&ucirc;)
            s = s.Replace("" + (char) 252, "u");
            s = s.Replace("&#252;", "u"); //Latin small letter u with diaeresis (&uuml;)
            s = s.Replace("" + (char) 253, "y");
            s = s.Replace("&#253;", "y"); //Latin small letter y with acute (&yacute;)
            s = s.Replace("" + (char) 254, "b");
            s = s.Replace("&#254;", "b"); //Latin small letter thorn (&thorn;)
            s = s.Replace("" + (char) 255, "y");
            s = s.Replace("&#255;", "y"); //Latin small letter y with diaeresis (&yuml;)
            s = s.Replace("" + (char) 338, "CE");
            s = s.Replace("&#338;", "CE"); //Latin capital ligature OE (&OElig;)
            s = s.Replace("" + (char) 339, "ce");
            s = s.Replace("&#339;", "ce"); //Latin small ligature oe (&oelig;)
            s = s.Replace("" + (char) 352, "S");
            s = s.Replace("&#352;", "S"); //Latin capital letter S with caron (&Scaron;)
            s = s.Replace("" + (char) 353, "s");
            s = s.Replace("&#353;", "s"); //Latin small letter s with caron (&scaron;)
            s = s.Replace("" + (char) 376, "Y");
            s = s.Replace("&#376;", "Y"); //Latin capital letter Y with diaeresis (&Yuml;)
            s = s.Replace("" + (char) 402, "f");
            s = s.Replace("&#402;", "f"); //Latin small f with hook = function = florin (&fnof;)
            s = s.Replace("" + (char) 710, "^");
            s = s.Replace("&#710;", "^"); //modifier letter circumflex accent (&circ;)
            s = s.Replace("" + (char) 732, "~");
            s = s.Replace("&#732;", "~"); //small tilde (&tilde;)
            s = s.Replace("" + (char) 913, "A");
            s = s.Replace("&#913;", "A"); //Greek capital letter alpha (&Alpha;)
            s = s.Replace("" + (char) 914, "B");
            s = s.Replace("&#914;", "B"); //Greek capital letter beta (&Beta;)
            s = s.Replace("" + (char) 915, "G");
            s = s.Replace("&#915;", "G"); //Greek capital letter gamma (&Gamma;)
            s = s.Replace("" + (char) 916, "D");
            s = s.Replace("&#916;", "D"); //Greek capital letter delta (&Delta;)
            s = s.Replace("" + (char) 917, "E");
            s = s.Replace("&#917;", "E"); //Greek capital letter epsilon (&Epsilon;)
            s = s.Replace("" + (char) 918, "Z");
            s = s.Replace("&#918;", "Z"); //Greek capital letter zeta (&Zeta;)
            s = s.Replace("" + (char) 919, "H");
            s = s.Replace("&#919;", "H"); //Greek capital letter eta (&Eta;)
            s = s.Replace("" + (char) 920, "TH");
            s = s.Replace("&#920;", "TH"); //Greek capital letter theta (&Theta;)
            s = s.Replace("" + (char) 921, "I");
            s = s.Replace("&#921;", "I"); //Greek capital letter iota (&Iota;)
            s = s.Replace("" + (char) 922, "K");
            s = s.Replace("&#922;", "K"); //Greek capital letter kappa (&Kappa;)
            s = s.Replace("" + (char) 923, "LAM");
            s = s.Replace("&#923;", "LAM"); //Greek capital letter lambda (&Lambda;)
            s = s.Replace("" + (char) 924, "M");
            s = s.Replace("&#924;", "M"); //Greek capital letter mu (&Mu;)
            s = s.Replace("" + (char) 925, "N");
            s = s.Replace("&#925;", "N"); //Greek capital letter nu (&Nu;)
            s = s.Replace("" + (char) 926, "XI");
            s = s.Replace("&#926;", "XI"); //Greek capital letter xi (&Xi;)
            s = s.Replace("" + (char) 927, "O");
            s = s.Replace("&#927;", "O"); //Greek capital letter omicron (&Omicron;)
            s = s.Replace("" + (char) 928, "PI");
            s = s.Replace("&#928;", "PI"); //Greek capital letter pi (&Pi;)
            s = s.Replace("" + (char) 929, "P");
            s = s.Replace("&#929;", "P"); //Greek capital letter rho (&Rho;)
            s = s.Replace("" + (char) 931, "SIGMA");
            s = s.Replace("&#931;", "SIGMA"); //Greek capital letter sigma (&Sigma;)
            s = s.Replace("" + (char) 932, "T");
            s = s.Replace("&#932;", "T"); //Greek capital letter tau (&Tau;)
            s = s.Replace("" + (char) 933, "Y");
            s = s.Replace("&#933;", "Y"); //Greek capital letter upsilon (&Upsilon;)
            s = s.Replace("" + (char) 934, "PHI");
            s = s.Replace("&#934;", "PHI"); //Greek capital letter phi (&Phi;)
            s = s.Replace("" + (char) 935, "X");
            s = s.Replace("&#935;", "X"); //Greek capital letter chi (&Chi;)
            s = s.Replace("" + (char) 936, "PSI");
            s = s.Replace("&#936;", "PSI"); //Greek capital letter psi (&Psi;)
            s = s.Replace("" + (char) 937, "OMEGA");
            s = s.Replace("&#937;", "OMEGA"); //Greek capital letter omega (&Omega;)
            s = s.Replace("" + (char) 945, "a");
            s = s.Replace("&#945;", "a"); //Greek small letter alpha (&alpha;)
            s = s.Replace("" + (char) 946, "b");
            s = s.Replace("&#946;", "b"); //Greek small letter beta (&beta;)
            s = s.Replace("" + (char) 947, "g");
            s = s.Replace("&#947;", "g"); //Greek small letter gamma (&gamma;)
            s = s.Replace("" + (char) 948, "d");
            s = s.Replace("&#948;", "d"); //Greek small letter delta (&delta;)
            s = s.Replace("" + (char) 949, "e");
            s = s.Replace("&#949;", "e"); //Greek small letter epsilon (&epsilon;)
            s = s.Replace("" + (char) 950, "z");
            s = s.Replace("&#950;", "z"); //Greek small letter zeta (&zeta;)
            s = s.Replace("" + (char) 951, "eta");
            s = s.Replace("&#951;", "eta"); //Greek small letter eta (&eta;)
            s = s.Replace("" + (char) 952, "theta");
            s = s.Replace("&#952;", "theta"); //Greek small letter theta (&theta;)
            s = s.Replace("" + (char) 953, "iota");
            s = s.Replace("&#953;", "iota"); //Greek small letter iota (&iota;)
            s = s.Replace("" + (char) 954, "k");
            s = s.Replace("&#954;", "k"); //Greek small letter kappa (&kappa;)
            s = s.Replace("" + (char) 955, "lamda");
            s = s.Replace("&#955;", "lamda"); //Greek small letter lambda (&lambda;)
            s = s.Replace("" + (char) 956, "mu");
            s = s.Replace("&#956;", "mu"); //Greek small letter mu (&mu;)
            s = s.Replace("" + (char) 957, "nu");
            s = s.Replace("&#957;", "nu"); //Greek small letter nu (&nu;)
            s = s.Replace("" + (char) 958, "xi");
            s = s.Replace("&#958;", "xi"); //Greek small letter xi (&xi;)
            s = s.Replace("" + (char) 959, "omicron");
            s = s.Replace("&#959;", "omicron"); //Greek small letter omicron (&omicron;)
            s = s.Replace("" + (char) 960, "pi");
            s = s.Replace("&#960;", "pi"); //Greek small letter pi (&pi;)
            s = s.Replace("" + (char) 961, "rho");
            s = s.Replace("&#961;", "rho"); //Greek small letter rho (&rho;)
            s = s.Replace("" + (char) 962, "sigmaf");
            s = s.Replace("&#962;", "sigmaf"); //Greek small letter final sigma (&sigmaf;)
            s = s.Replace("" + (char) 963, "sigmaf");
            s = s.Replace("&#963;", "sigmaf"); //Greek small letter sigma (&sigma;)
            s = s.Replace("" + (char) 964, "tau");
            s = s.Replace("&#964;", "tau"); //Greek small letter tau (&tau;)
            s = s.Replace("" + (char) 965, "upsilon");
            s = s.Replace("&#965;", "upsilon"); //Greek small letter upsilon (&upsilon;)
            s = s.Replace("" + (char) 966, "phi");
            s = s.Replace("&#966;", "phi"); //Greek small letter phi (&phi;)
            s = s.Replace("" + (char) 967, "chi");
            s = s.Replace("&#967;", "chi"); //Greek small letter chi (&chi;)
            s = s.Replace("" + (char) 968, "psi");
            s = s.Replace("&#968;", "psi"); //Greek small letter psi (&psi;)
            s = s.Replace("" + (char) 969, "omega");
            s = s.Replace("&#969;", "omega"); //Greek small letter omega (&omega;)
            s = s.Replace("" + (char) 977, "theta");
            s = s.Replace("&#977;", "theta"); //Greek small letter theta symbol (&thetasym;)
            s = s.Replace("" + (char) 978, "upsilonh");
            s = s.Replace("&#978;", "upsilonh"); //Greek upsilon with hook symbol (&upsih;)
            s = s.Replace("" + (char) 982, "pi");
            s = s.Replace("&#982;", "pi"); //Greek pi symbol (&piv;)
            s = s.Replace("" + (char) 8194, " ");
            s = s.Replace("&#8194;", " "); //en space (&ensp;)
            s = s.Replace("" + (char) 8195, " ");
            s = s.Replace("&#8195;", " "); //em space (&emsp;)
            s = s.Replace("" + (char) 8201, " ");
            s = s.Replace("&#8201;", " "); //thin space (&thinsp;)
            s = s.Replace("" + (char) 8204, "");
            s = s.Replace("&#8204;", ""); //zero width non-joiner (&zwnj;)
            s = s.Replace("" + (char) 8205, "");
            s = s.Replace("&#8205;", ""); //zero width joiner (&zwj;)
            s = s.Replace("" + (char) 8206, "");
            s = s.Replace("&#8206;", ""); //left-to-right mark (&lrm;)
            s = s.Replace("" + (char) 8207, "");
            s = s.Replace("&#8207;", ""); //right-to-left mark (&rlm;)
            s = s.Replace("" + (char) 8211, "-");
            s = s.Replace("&#8211;", "-"); //en dash (&ndash;)
            s = s.Replace("" + (char) 8212, "--");
            s = s.Replace("&#8212;", "--"); //em dash (&mdash;)
            s = s.Replace("" + (char) 8216, "'");
            s = s.Replace("&#8216;", "'"); //left single quotation mark (&lsquo;)
            s = s.Replace("" + (char) 8217, "'");
            s = s.Replace("&#8217;", "'"); //right single quotation mark (&rsquo;)
            s = s.Replace("" + (char) 8218, ",");
            s = s.Replace("&#8218;", ","); //single low-9 quotation mark (&sbquo;)
            s = s.Replace("" + (char) 8220, "\"");
            s = s.Replace("&#8220;", "\""); //left double quotation mark (&ldquo;)
            s = s.Replace("" + (char) 8221, "\"");
            s = s.Replace("&#8221;", "\""); //right double quotation mark (&rdquo;)
            s = s.Replace("" + (char) 8222, ",,");
            s = s.Replace("&#8222;", ",,"); //double low-9 quotation mark (&bdquo;)
            s = s.Replace("" + (char) 8224, "+");
            s = s.Replace("&#8224;", "+"); //dagger (&dagger;)
            s = s.Replace("" + (char) 8225, "++");
            s = s.Replace("&#8225;", "++"); //double dagger (&Dagger;)
            s = s.Replace("" + (char) 8226, "*");
            s = s.Replace("&#8226;", "*"); //bullet = black small circle (&bull;)
            s = s.Replace("" + (char) 8230, "…");
            s = s.Replace("&#8230;", "…"); //horizontal ellipsis = three dot leader (&hellip;)
            s = s.Replace("" + (char) 8240, "/100");
            s = s.Replace("&#8240;", "/100"); //per mille sign (&permil;)
            s = s.Replace("" + (char) 8242, "'");
            s = s.Replace("&#8242;", "'"); //prime = minutes = feet (&prime;)
            s = s.Replace("" + (char) 8243, "\"");
            s = s.Replace("&#8243;", "\""); //double prime = seconds = inches (&Prime;)
            s = s.Replace("" + (char) 8249, "<");
            s = s.Replace("&#8249;", "<"); //single left-pointing angle quotation mark (&lsaquo;)
            s = s.Replace("" + (char) 8250, ">");
            s = s.Replace("&#8250;", ">"); //single right-pointing angle quotation mark (&rsaquo;)
            s = s.Replace("" + (char) 8254, "-");
            s = s.Replace("&#8254;", "-"); //overline = spacing overscore (&oline;)
            s = s.Replace("" + (char) 8260, "/");
            s = s.Replace("&#8260;", "/"); //fraction slash (&frasl;)
            s = s.Replace("" + (char) 8364, "euro");
            s = s.Replace("&#8364;", "euro"); //euro sign (&euro;)
            s = s.Replace("" + (char) 8465, "I");
            s = s.Replace("&#8465;", "I"); //blackletter capital I = imaginary part (&image;)
            s = s.Replace("" + (char) 8472, "P");
            s = s.Replace("&#8472;", "P"); //script capital P = power set = Weierstrass p (&weierp;)
            s = s.Replace("" + (char) 8476, "R");
            s = s.Replace("&#8476;", "R"); //blackletter capital R = real part symbol (&real;)
            s = s.Replace("" + (char) 8482, "™");
            s = s.Replace("&#8482;", "™"); //trade mark sign (&trade;)
            s = s.Replace("" + (char) 8501, "X");
            s = s.Replace("&#8501;", "X"); //alef symbol = first transfinite cardinal (&alefsym;)
            s = s.Replace("" + (char) 8592, "<-");
            s = s.Replace("&#8592;", "<-"); //leftwards arrow (&larr;)
            s = s.Replace("" + (char) 8593, "^");
            s = s.Replace("&#8593;", "^"); //upwards arrow (&uarr;)
            s = s.Replace("" + (char) 8594, "->");
            s = s.Replace("&#8594;", "->"); //rightwards arrow (&rarr;)
            s = s.Replace("" + (char) 8595, "v");
            s = s.Replace("&#8595;", "v"); //downwards arrow (&darr;)
            s = s.Replace("" + (char) 8596, "<->");
            s = s.Replace("&#8596;", "<->"); //left right arrow (&harr;)
            s = s.Replace("" + (char) 8629, "<-");
            s = s.Replace("&#8629;", "<-"); //downwards arrow with corner leftwards = carriage return (&crarr;)
            s = s.Replace("" + (char) 8656, "<=");
            s = s.Replace("&#8656;", "<="); //leftwards double arrow (&lArr;)
            s = s.Replace("" + (char) 8657, "^");
            s = s.Replace("&#8657;", "^"); //upwards double arrow (&uArr;)
            s = s.Replace("" + (char) 8658, "=>");
            s = s.Replace("&#8658;", "=>"); //rightwards double arrow (&rArr;)
            s = s.Replace("" + (char) 8659, "V");
            s = s.Replace("&#8659;", "V"); //downwards double arrow (&dArr;)
            s = s.Replace("" + (char) 8660, "<=>");
            s = s.Replace("&#8660;", "<=>"); //left right double arrow (&hArr;)
            s = s.Replace("" + (char) 8704, "for all");
            s = s.Replace("&#8704;", "for all"); //for all (&forall;)
            s = s.Replace("" + (char) 8706, "partial differential");
            s = s.Replace("&#8706;", "partial differential"); //partial differential (&part;)
            s = s.Replace("" + (char) 8707, "there exists");
            s = s.Replace("&#8707;", "there exists"); //there exists (&exist;)
            s = s.Replace("" + (char) 8709, "empty set = null set = diameter");
            s = s.Replace("&#8709;", "empty set = null set = diameter"); //empty set = null set = diameter (&empty;)
            s = s.Replace("" + (char) 8711, "nabla = backward difference");
            s = s.Replace("&#8711;", "nabla = backward difference"); //nabla = backward difference (&nabla;)
            s = s.Replace("" + (char) 8712, "element of");
            s = s.Replace("&#8712;", "element of"); //element of (&isin;)
            s = s.Replace("" + (char) 8713, "not an element of");
            s = s.Replace("&#8713;", "not an element of"); //not an element of (&notin;)
            s = s.Replace("" + (char) 8715, "contains as member");
            s = s.Replace("&#8715;", "contains as member"); //contains as member (&ni;)
            s = s.Replace("" + (char) 8719, "n-ary product = product sign");
            s = s.Replace("&#8719;", "n-ary product = product sign"); //n-ary product = product sign (&prod;)
            s = s.Replace("" + (char) 8721, "n-ary sumation");
            s = s.Replace("&#8721;", "n-ary sumation"); //n-ary sumation (&sum;)
            s = s.Replace("" + (char) 8722, "minus sign");
            s = s.Replace("&#8722;", "minus sign"); //minus sign (&minus;)
            s = s.Replace("" + (char) 8727, "asterisk operator");
            s = s.Replace("&#8727;", "asterisk operator"); //asterisk operator (&lowast;)
            s = s.Replace("" + (char) 8730, "square root = radical sign");
            s = s.Replace("&#8730;", "square root = radical sign"); //square root = radical sign (&radic;)
            s = s.Replace("" + (char) 8733, "proportional to");
            s = s.Replace("&#8733;", "proportional to"); //proportional to (&prop;)
            s = s.Replace("" + (char) 8734, "infinity");
            s = s.Replace("&#8734;", "infinity"); //infinity (&infin;)
            s = s.Replace("" + (char) 8736, "angle");
            s = s.Replace("&#8736;", "angle"); //angle (&ang;)
            s = s.Replace("" + (char) 8743, "logical and = wedge");
            s = s.Replace("&#8743;", "logical and = wedge"); //logical and = wedge (&and;)
            s = s.Replace("" + (char) 8744, "logical or = vee");
            s = s.Replace("&#8744;", "logical or = vee"); //logical or = vee (&or;)
            s = s.Replace("" + (char) 8745, "intersection = cap");
            s = s.Replace("&#8745;", "intersection = cap"); //intersection = cap (&cap;)
            s = s.Replace("" + (char) 8746, "union = cup");
            s = s.Replace("&#8746;", "union = cup"); //union = cup (&cup;)
            s = s.Replace("" + (char) 8747, "integral");
            s = s.Replace("&#8747;", "integral"); //integral (&int;)
            s = s.Replace("" + (char) 8756, "therefore");
            s = s.Replace("&#8756;", "therefore"); //therefore (&there4;)
            s = s.Replace("" + (char) 8764, "tilde operator = varies with = similar to");
            s = s.Replace("&#8764;", "tilde operator = varies with = similar to");
            //tilde operator = varies with = similar to (&sim;)
            s = s.Replace("" + (char) 8773, "approximately equal to");
            s = s.Replace("&#8773;", "approximately equal to"); //approximately equal to (&cong;)
            s = s.Replace("" + (char) 8776, "almost equal to = asymptotic to");
            s = s.Replace("&#8776;", "almost equal to = asymptotic to"); //almost equal to = asymptotic to (&asymp;)
            s = s.Replace("" + (char) 8800, "not equal to");
            s = s.Replace("&#8800;", "not equal to"); //not equal to (&ne;)
            s = s.Replace("" + (char) 8801, "identical to");
            s = s.Replace("&#8801;", "identical to"); //identical to (&equiv;)
            s = s.Replace("" + (char) 8804, "less-than or equal to");
            s = s.Replace("&#8804;", "less-than or equal to"); //less-than or equal to (&le;)
            s = s.Replace("" + (char) 8805, "greater-than or equal to");
            s = s.Replace("&#8805;", "greater-than or equal to"); //greater-than or equal to (&ge;)
            s = s.Replace("" + (char) 8834, "subset of");
            s = s.Replace("&#8834;", "subset of"); //subset of (&sub;)
            s = s.Replace("" + (char) 8835, "superset of");
            s = s.Replace("&#8835;", "superset of"); //superset of (&sup;)
            s = s.Replace("" + (char) 8836, "not a subset of");
            s = s.Replace("&#8836;", "not a subset of"); //not a subset of (&nsub;)
            s = s.Replace("" + (char) 8838, "subset of or equal to");
            s = s.Replace("&#8838;", "subset of or equal to"); //subset of or equal to (&sube;)
            s = s.Replace("" + (char) 8839, "superset of or equal to");
            s = s.Replace("&#8839;", "superset of or equal to"); //superset of or equal to (&supe;)
            s = s.Replace("" + (char) 8853, "circled plus = direct sum");
            s = s.Replace("&#8853;", "circled plus = direct sum"); //circled plus = direct sum (&oplus;)
            s = s.Replace("" + (char) 8855, "circled times = vector product");
            s = s.Replace("&#8855;", "circled times = vector product"); //circled times = vector product (&otimes;)
            s = s.Replace("" + (char) 8869, "up tack = orthogonal to = perpendicular");
            s = s.Replace("&#8869;", "up tack = orthogonal to = perpendicular");
            //up tack = orthogonal to = perpendicular (&perp;)
            s = s.Replace("" + (char) 8901, "dot operator");
            s = s.Replace("&#8901;", "dot operator"); //dot operator (&sdot;)
            s = s.Replace("" + (char) 8968, "left ceiling = APL upstile");
            s = s.Replace("&#8968;", "left ceiling = APL upstile"); //left ceiling = APL upstile (&lceil;)
            s = s.Replace("" + (char) 8969, "right ceiling");
            s = s.Replace("&#8969;", "right ceiling"); //right ceiling (&rceil;)
            s = s.Replace("" + (char) 8970, "left floor = APL downstile");
            s = s.Replace("&#8970;", "left floor = APL downstile"); //left floor = APL downstile (&lfloor;)
            s = s.Replace("" + (char) 8971, "right floor");
            s = s.Replace("&#8971;", "right floor"); //right floor (&rfloor;)
            s = s.Replace("" + (char) 9001, "left-pointing angle bracket = bra");
            s = s.Replace("&#9001;", "left-pointing angle bracket = bra"); //left-pointing angle bracket = bra (&lang;)
            s = s.Replace("" + (char) 9002, "right-pointing angle bracket = ket");
            s = s.Replace("&#9002;", "right-pointing angle bracket = ket");
            //right-pointing angle bracket = ket (&rang;)
            s = s.Replace("" + (char) 9674, "lozenge");
            s = s.Replace("&#9674;", "lozenge"); //lozenge (&loz;)
            s = s.Replace("" + (char) 9824, "black spade suit");
            s = s.Replace("&#9824;", "black spade suit"); //black spade suit (&spades;)
            s = s.Replace("" + (char) 9827, "black club suit = shamrock");
            s = s.Replace("&#9827;", "black club suit = shamrock"); //black club suit = shamrock (&clubs;)
            s = s.Replace("" + (char) 9829, "black heart suit = valentine");
            s = s.Replace("&#9829;", "black heart suit = valentine"); //black heart suit = valentine (&hearts;)
            s = s.Replace("" + (char) 9830, "black diamond suit");
            s = s.Replace("&#9830;", "black diamond suit"); //black diamond suit (&diams;)
            //
            // End of code generated by Entities.xlsx
            //
            s = s.Replace("aeuro™", "'"); //Convert weird Word aprostrophe
            s = s.Replace("&nbsp;", " "); //Convert no-break-spaces
            s = s.Replace("&amp;", " "); //Convert ampersand
            s = s.Replace("  ", " "); //Convert multiple spaces
            s = s.Replace("\"", "'"); //convert multiple double-quotes to single-quotes for CSV file
            s = s.Replace("''", "'"); //Convert multiple single quotes
            return s.Trim();
            // ReSharper restore CommentTypo
            // ReSharper restore StringLiteralTypo
        }

        [Category("String")]
        [Description("Replace XML predefined entities per https://en.wikipedia.org/wiki/List_of_XML_and_HTML_character_entity_references")]
        public static string ReplaceXmlSpecialCharacters(string xml)
        {
            xml = xml.Replace("&", "&amp;");
            xml = xml.Replace("\"", "&quot;");
            xml = xml.Replace("'", "&apos;");
            xml = xml.Replace("<", "&lt;");
            xml = xml.Replace(">", "&gt;");
            return xml;
        }

        [Category("String")]
        [Description("Remove non-ASCII characters")]
        public static string RemoveNonAsciiCharacters(string s)
        {
            // Remove all non-ASCII characters (used to remove Microsoft and Unicode characters)
            s = Encoding.ASCII.GetString(Encoding.ASCII.GetBytes(s));
            s = s.Replace("?", "");
            return s;
        }

        #endregion String Helpers

        #region Documentation Helpers

        [AttributeUsage(AttributeTargets.All)]
        public class CategoryAttribute : Attribute
        {
            public readonly string Category;

            public CategoryAttribute(string category)
            {
                Category = category;
            }
        }

        [AttributeUsage(AttributeTargets.All)]
        public class DescriptionAttribute : Attribute
        {
            public readonly string Description;

            public DescriptionAttribute(string description)
            {
                Description = description;
            }
        }

        public class DocItem
        {
            public DocItem(string category, string name, string objType, string retType, Dictionary<string, string> parms, string description)
            {
                this.category = category;
                this.name = name;
                this.objType = objType;
                this.retType = retType;
                this.parms = parms;
                this.description = description;
            }

            public string category { get; set; }
            public string name { get; set; }
            public string objType { get; set; }
            public string retType { get; set; }
            public Dictionary<string, string> parms { get; set; }
            public string description { get; set; }
        }

        [Category("Documentation")]
        [Description("Auto-generate HTML document for this class.")]
        public static void SelfDocument2Html()
        {
            string ret = "<!DOCTYPE html>\n";
            ret += "<html lang='en'><head>\n";
            ret += "<meta charset='UTF-8'>\n";
            ret += "<title>RLTLIB2.cs</title>\n";
            ret += "<style>\n";
            ret += "body { font-family: sans-serif; font-size: small; }\n";
            ret += "table { border: solid 1px Black; border-collapse: collapse; font-size: small; }\n";
            ret += "td { border: solid 1px Black; padding: 0 0.5em 0 0.5em; }\n";
            ret += "th { background-color: LightBlue; border: solid 1px Black; padding: 0.25em 1em 0.25em 1em; }\n";
            ret += ".category { color: DarkGreen; }\n";
            ret += ".title { font-size: 150%; }\n";
            ret += ".name { font-weight: bold; color: Navy; }\n";
            ret += ".param { color: Maroon; }\n";
            ret += ".desc { color: DarkGreen; width: 600px; }\n";
            ret += "</style>\n";
            ret += "</head><body>\n";
            ret += $"<table><tr><th colspan='6' class='title'>{AppAbbr} last modified {LastModifiedDate}</th></tr>\n";
            ret += "<tr><th>Category</th><th>Name</th><th>Type</th><th>Returns</th><th>Parameters</th><th>Description</th></tr>\n";

            Type rltlib2 = typeof(RLTLIB2);
            List<DocItem> docItems = new List<DocItem>();

            foreach (MethodInfo mi in rltlib2.GetMethods())
                // ReSharper disable once PossibleNullReferenceException
                if (mi.Name != ".ctor" && mi.DeclaringType.Namespace != "System" && !mi.Name.StartsWith("get_") && !mi.Name.StartsWith("set_"))
                {
                    CustomAttributeData c = mi.CustomAttributes.FirstOrDefault(a => a.AttributeType == typeof(CategoryAttribute));
                    string category = c == null ? "" : c.ConstructorArguments[0].Value.ToString();
                    string name = mi.Name;
                    string objType = mi.MemberType.ToString();
                    string retType = mi.ReturnType.Name;
                    Dictionary<string, string> parms = mi.GetParameters().ToDictionary(p => p.Name, p => p.ParameterType.Name);
                    CustomAttributeData d = mi.CustomAttributes.FirstOrDefault(a => a.AttributeType == typeof(DescriptionAttribute));
                    string description = d == null ? "" : d.ConstructorArguments[0].Value.ToString();
                    docItems.Add(new DocItem(category, name, objType, retType, parms, description));
                }

            foreach (PropertyInfo pi in rltlib2.GetProperties())
                // ReSharper disable once PossibleNullReferenceException
                if (pi.Name != ".ctor" && pi.DeclaringType.Namespace != "System")
                {
                    CustomAttributeData c = pi.CustomAttributes.FirstOrDefault(a => a.AttributeType == typeof(CategoryAttribute));
                    string category = c == null ? "" : c.ConstructorArguments[0].Value.ToString();
                    string name = pi.Name;
                    string objType = pi.MemberType.ToString();
                    string retType = "";
                    Dictionary<string, string> parms = new Dictionary<string, string>();
                    CustomAttributeData d = pi.CustomAttributes.FirstOrDefault(a => a.AttributeType == typeof(DescriptionAttribute));
                    string description = d == null ? "" : d.ConstructorArguments[0].Value.ToString();
                    docItems.Add(new DocItem(category, name, objType, retType, parms, description));
                }

            foreach (MemberInfo mi in rltlib2.GetMembers())
                // ReSharper disable once PossibleNullReferenceException
                if (mi.Name != ".ctor" && mi.DeclaringType.Namespace != "System" && mi.MemberType.ToString() == "Field")
                {
                    CustomAttributeData c = mi.CustomAttributes.FirstOrDefault(a => a.AttributeType == typeof(CategoryAttribute));
                    string category = c == null ? "" : c.ConstructorArguments[0].Value.ToString();
                    string name = mi.Name;
                    string objType = mi.MemberType.ToString();
                    string retType = "";
                    Dictionary<string, string> parms = new Dictionary<string, string>();
                    CustomAttributeData d = mi.CustomAttributes.FirstOrDefault(a => a.AttributeType == typeof(DescriptionAttribute));
                    string description = d == null ? "" : d.ConstructorArguments[0].Value.ToString();
                    docItems.Add(new DocItem(category, name, objType, retType, parms, description));
                }

            docItems.Add(new DocItem("File/Folder", "CopyFolderTreeLoggingLevel", "Enum", "value",
                new Dictionary<string, string>
                {
                    {"None", "No logging"},
                    {"Summary", "Only log summary"},
                    {"BigFiles", "Log files > 10MB"},
                    {"Thousand", "Log every 1000th file copied"},
                    {"Hundred", "Log every 100th file copied"},
                    {"IgnoresExcludes", "Log ignored and excluded files"},
                    {"All", "Log all files copied"},
                    {"Verbose", "Log all files copied and folders created"}
                },
                "CopyFolderTree Logging Level"));

            docItems.Add(new DocItem("StringExtensions", "toRelative", "Extension", "String",
                new Dictionary<string, string> {{"url", "String"}}, "Return url with backslashes and trim initial slash."));

            docItems.Add(new DocItem("StringExtensions", "toRelativeWithoutHref", "Extension", "String",
                new Dictionary<string, string> {{"url", "String"}}, "Return relative url minus href."));

            docItems.Add(new DocItem("StringExtensions", "toBackSlash", "Extension", "String",
                new Dictionary<string, string> {{"url", "String"}}, "Return url with backslashes."));

            docItems.Add(new DocItem("StringExtensions", "toForwardSlash", "Extension", "String",
                new Dictionary<string, string> {{"url", "String"}}, "Return url with forward slashes."));

            docItems.Add(new DocItem("StringExtensions", "toCompressWhiteSpace", "Extension", "String",
                new Dictionary<string, string> {{"text", "String"}}, "Compress whitespace."));

            foreach (DocItem item in docItems.OrderBy(x => x.category).ThenBy(x => x.name).ToList())
            {
                ret += "<tr>";
                ret += $"<td class='category'>{item.category}</td>";
                ret += $"<td class='name'>{item.name}</td>";
                ret += $"<td>{item.objType}</td>";
                ret += $"<td>{item.retType}</td>";
                ret +=
                    $"<td>{item.parms.Aggregate("", (current, pair) => current + $"<span class='param'>{pair.Key}</span>&nbsp;({pair.Value})<br>")}</td>";
                ret += $"<td class='desc'>{item.description}</td>";
                ret += "</tr>\n";
            }

            ret += "</table></body></html>\n";

            WriteTextFile("RLTLIB2.html", ret);
        }

        #endregion Documentation Helper
    }

    #region String Extensions for URL Manipulation

    public static class StringExtensions
    {
        public static string toRelative(this string url)
        {
            return url.toBackSlash().TrimStart('\\');
        }

        public static string toRelativeWithoutHref(this string url, string href)
        {
            return url.toRelative().Replace(href.toRelative(), string.Empty);
        }

        public static string toBackSlash(this string url)
        {
            return url.Replace("/", "\\");
        }

        public static string toForwardSlash(this string url)
        {
            return url.Replace("\\", "/");
        }

        public static string toCompressWhiteSpace(this string text)
        {
            text = text.Trim();
            text = text.Replace("\t", " ");
            text = text.Replace("\n", " ");
            text = text.Replace("\r", " ");
            text = text.Replace("  ", " ");
            while (text.Contains("  "))
                text = text.Replace("  ", " ");
            return text;
        }
    }

    #endregion String Extensions for URL Manipulation
}