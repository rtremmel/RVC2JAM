using System;
using System.Configuration;
using System.Data;
using VectorSolutions;

namespace RVC2JAM
{
    public static class ContentSet
    {
        public static readonly string ContentControlJamXlsx = ConfigurationManager.AppSettings["ContentControlJamXlsx"];
        public static readonly string OnlyProcessThisPriority = ConfigurationManager.AppSettings["OnlyProcessThisPriority"];
        public static readonly string OnlyProcessThisRvSku = ConfigurationManager.AppSettings["OnlyProcessThisRvSku"];
        public static int ContentControlTotalCount;
        public static int ContentControlSelectedCount;
        public static int ContentControlSuccessful;
        public static int ContentControlUnsuccessful;

        public static DataTable CoursesToProcessTable
        {
            get
            {
                RLTLIB2.Log($"Loading Control Spreadsheet {ContentControlJamXlsx}'...");
                DataTable dt = RLTLIB2.LoadExcelSheet(ContentControlJamXlsx, "", out string error);
                if (!string.IsNullOrEmpty(error))
                {
                    error = $"Error loading '{ContentControlJamXlsx}'.  Ensure it exists.  If it does, logoff and login again.";
                    throw new SystemException(error);
                }

                ContentControlTotalCount = dt.Rows.Count;
                RLTLIB2.Log($"Content Control Spreadsheet contains {RLTLIB2.Pluralize(ContentControlTotalCount,"course")}");

                if (!string.IsNullOrWhiteSpace(OnlyProcessThisPriority))
                {
                    RLTLIB2.Log($"*** ONLY PROCESSING PRIORITY '{OnlyProcessThisPriority}' ***");
                    dt = dt.AsEnumerable().Where(r => r["Priority"].ToString() == OnlyProcessThisPriority).CopyToDataTable();
                }

                if (!string.IsNullOrWhiteSpace(OnlyProcessThisRvSku))
                {
                    RLTLIB2.Log($"*** ONLY PROCESSING COURSE '{OnlyProcessThisRvSku}' ***");
                    dt = dt.AsEnumerable().Where(r => r["Course ID"].ToString() == OnlyProcessThisRvSku).CopyToDataTable();
                }

                ContentControlSelectedCount = dt.Rows.Count;
                RLTLIB2.Log($"Ready to process {RLTLIB2.Pluralize(ContentControlSelectedCount, "selected course")}");

                return dt;
            }
        }

        public static string ExcludedCourses()
        {
            string excluded = "";
            string text = RLTLIB2.ReadTextFile("CoursesExcluded.txt");
            string[] rvskus = text.Split(new[] {Environment.NewLine}, StringSplitOptions.None);
            foreach (var rvsku in rvskus)
                excluded += $",\'{rvsku.Split('\t')[0].Trim()}\'";
            return excluded.Substring(1);
        }

        public static string BadCourses()
        {
            string bad = "";
            string text = RLTLIB2.ReadTextFile("CoursesBad.txt");
            string[] rvskus = text.Split(new[] {Environment.NewLine}, StringSplitOptions.None);
            foreach (var rvsku in rvskus)
                bad += $",\'{rvsku.Split('\t')[0].Trim()}\'";
            return bad.Substring(1);
        }
    }
}