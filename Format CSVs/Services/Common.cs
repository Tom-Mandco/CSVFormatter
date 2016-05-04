using System.IO;

namespace MandCo.CSVFormatter.Services
{
    class Common
    {
        public static int CSVRowCount(string strFilePath)
        {
            int rowCount = 0;
            StreamReader sr = new StreamReader(strFilePath);
            System.Data.DataTable dt = new System.Data.DataTable();
            sr.ReadLine();
            while (!sr.EndOfStream)
            {
                string[] rows = sr.ReadLine().Split(',');
                System.Data.DataRow dr = dt.NewRow();
                rowCount++;
                dt.Rows.Add(dr);
            }
            return rowCount;
        }

        public static string CleanSpreadsheetName(string rawSpreadsheetName)
        {
            var charsToRemove = new string[] { ":", "\\", "/", "?", "*", "[", "]" };
            foreach (var c in charsToRemove)
            {
                rawSpreadsheetName = rawSpreadsheetName.Replace(c, " ");
            }

            return rawSpreadsheetName;
        }



    }
}
