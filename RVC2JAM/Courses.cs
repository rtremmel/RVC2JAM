using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using VectorSolutions;

namespace RVC2JAM
{
    // This code was copied and modified from RVCC2J on 08/27/2019
    public class Courses
    {
        private static DataTable LoadRvCourseInfo(List<string> rvSkus)
        {
            string sql = "SELECT ci.rv_sku, ";
            sql += "       ci.catalog_item_id, ";
            sql += "       '' AS course_uuid, ";
            sql += "       ci.title AS course_title, ";
            sql += "       ci.status, ";
            sql += "       ci.course_player_type, ";
            sql += "       cu.title AS lesson_title, ";
            sql += "       cu.course_unit_id, ";
            sql += "       cu.def_mastery_score, ";
            sql += "       '' AS lesson_uuid, ";
            sql += "       cis.code_ref AS import_schema, ";
            sql += "       LOWER(cu.launch_url) AS launch_url, ";
            sql += "       CASE ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.htm%' THEN 'Presentation' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.mp4' THEN 'Presentation' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.pdf' THEN 'Presentation' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.avi' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.csv' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.doc%' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.mov' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.pp%' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.swf' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.txt' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.wmv' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.xls%' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'HTML' AND cu.launch_url LIKE '%.zip' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'SCRM12' AND cu.launch_url LIKE '%.htm%' THEN 'SCORM Classic' ";
            sql += "           WHEN cis.code_ref = 'AICC' THEN 'AICC' ";
            sql += "           WHEN cis.code_ref = 'SCRM12' AND cu.launch_url LIKE '%.doc%' THEN 'Document' ";
            sql += "           WHEN cis.code_ref = 'SCRM12' AND cu.launch_url LIKE '%.pdf' THEN 'Presentation' ";
            sql += "           ELSE 'Error' ";
            sql += "       END AS module_type, ";
            sql += "       '' AS journey_module_type, ";
            sql += "       2 AS journey_ordinal, ";
            sql += "       '' AS journey_launch_file, ";
            sql += "       '' AS processing_error ";
            sql += "FROM dbo.rv_cat_catalog_item AS ci WITH (NOLOCK) ";
            sql += "INNER JOIN dbo.uvw_rvm_cat_course_list AS ccl WITH (NOLOCK) ON ccl.catalog_item_id = ci.catalog_item_id ";
            sql += "INNER JOIN dbo.rv_crs_course_unit AS cu WITH (NOLOCK) ON ccl.course_unit_id = cu.course_unit_id ";
            sql += "INNER JOIN dbo.rv_crs_import_schema AS cis WITH (NOLOCK) ON cis.import_schema_id = cu.import_schema_id ";
            sql += "WHERE ci.rv_sku='{0}' ";
            sql += "AND (ci.status IN (" + ConfigurationManager.AppSettings["CatalogStatusToProcess"] + ")) ";
            sql += "AND (cis.code_ref NOT IN ('RVLMS20TOP', 'PRETEST', 'RVLMS20PDF', 'RVLMS20ASS', 'RVLMS20SUR') ";
            sql += "     OR (cis.code_ref IN ('RVLMS20ASS') AND ci.rv_sku LIKE 'RVTN-%'))";
            sql += "ORDER BY ci.rv_sku, course_title ";

            DataTable courses = null;

            foreach (string rvSku in rvSkus)
            {
                DataRow row = CheckForNonStandardCourse(rvSku); // Returns non-standard course row or null

                if (row == null)
                {
                    DataTable dt = RLTLIB2.ExecuteQuery(string.Format(sql, rvSku), out _);

                    if (dt.Rows.Count == 0)
                    {
                        row = dt.NewRow();
                        row["rv_sku"] = rvSku;
                        row["course_title"] = "Unknown";
                        row["processing_error"] = $"Course {rvSku} not found in RedVector OR there is no lesson";
                        dt.Rows.Add(row);
                    }
                    else if (dt.Rows.Count > 1)
                    {
                        row = dt.Rows[0];
                        row["processing_error"] = $"Course {rvSku} in RedVector has multiple records OR multiple lessons";
                    }
                    else
                    {
                        row = dt.Rows[0];
                        if (IsCustomCourse(rvSku))
                            row["processing_error"] = $"Course {rvSku} is a RedVector Custom Course (process using RVCC2JV3)";
                    }
                }

                if (courses == null)
                    courses = row.Table;
                else
                    courses.ImportRow(row);
            }

            return courses;
        }

        private static DataRow CheckForNonStandardCourse(string rvSku)
        {
            string sql = "SELECT ci.rv_sku, ";
            sql += "       ci.catalog_item_id, ";
            sql += "       '00000000-0000-0000-0000-000000000000' AS course_uuid, ";
            sql += "       ci.title AS course_title, ";
            sql += "       ci.status, ";
            sql += "       ci.course_player_type, ";
            sql += "       '' AS lesson_title, ";
            sql += "       '00000000-0000-0000-0000-000000000000' as course_unit_id, ";
            sql += "       0 AS def_mastery_score, ";
            sql += "       '00000000-0000-0000-0000-000000000000' AS lesson_uuid, ";
            sql += "       '' AS import_schema, ";
            sql += "       '' AS launch_url, ";
            sql += "       '' module_type, ";
            sql += "       '' AS journey_module_type, ";
            sql += "       2 AS journey_ordinal, ";
            sql += "       '' AS journey_launch_file, ";
            sql += "       '' AS processing_error ";
            sql += "FROM dbo.rv_cat_catalog_item AS ci WITH (NOLOCK) ";
            sql += "WHERE ci.rv_sku='{0}' ";

            DataTable dt = RLTLIB2.ExecuteQuery(string.Format(sql, rvSku), out _);
            if (dt.Rows.Count == 0)
                return null;

            DataRow row = dt.Rows[0];
            string player;
            switch (row["course_player_type"].ToString())
            {
                case "":
                case "0":
                    player = "Normal";
                    break;
                case "1":
                    player = "Maestro Custom Course";
                    break;
                case "2":
                    player = "Career Development Tool";
                    break;
                default:
                    player = "Unknown course_player_type";
                    break;
            }

            if (player != "Normal")
            {
                row["processing_error"] = $"Course {rvSku} unable to convert {player}";
                return row;
            }

            return null;
        }

        private static bool IsCustomCourse(string rvSku)
        {
            string sql = $"SELECT rv_sku FROM dbo.rv_acc_custom_course WITH(NOLOCK) WHERE rv_sku='{rvSku}'";
            DataTable dt = RLTLIB2.ExecuteQuery(sql, out _);
            return dt.Rows.Count > 0;
        }
    }
}