using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using NLog;

namespace MandCo.CSVFormatter.Applications.Programs
{
    class Al510_Format
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public static void Run(string UniqueBatchNo)
        {
            List<string> PCC_IDs = new List<string>();
            List<string> csvFileNames = new List<string>();
            string csvFilePath = ConfigurationManager.AppSettings["RawReportPath"];
            string outputFilePath = ConfigurationManager.AppSettings["AL510OutputPath"];
            string fileName;
            string amalgamatedSpreadsheetName = (@outputFilePath + "(Al510) Run by " + Environment.UserName + " at " + ValidFilePathDate(DateTime.Now) + ".xlsx");
            int fileStartIndex;

            string[] files = Directory.GetFiles(@"\\" + csvFilePath, "[Raw]" + "Al510" + "(" + UniqueBatchNo + ")*.csv");
            Console.WriteLine("Running program: Al510 - PCC Report By Packs. URN: " + UniqueBatchNo + "\n\n");
            XLWorkbook xlwb = new XLWorkbook();

            int csvFileCounter = 0;

            foreach (string file in files)
            {
                int csvLineCounter = 0;
                fileStartIndex = file.IndexOf("\\[Raw]") + 1;
                fileName = file.Substring(fileStartIndex, (file.Length - fileStartIndex));
                Console.WriteLine("\n");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(csvFileCounter + 1 + " " + file);
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(" > " + fileName);

                int totalCSVLines = Services.Common.CSVRowCount(file);
                csvFileNames.Add(csvFilePath + fileName);
                var reader = new StreamReader(File.OpenRead(csvFilePath + fileName));
                System.Data.DataTable res = ConvertCSVtoDataTable(csvFileNames[csvFileCounter]);



                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    if (csvLineCounter == 0)
                    {
                        PCC_IDs.Add(values[0].Trim());

                        Console.Write(" > New PCC ID Found: ");
                        Console.ForegroundColor = ConsoleColor.Cyan;
                        Console.Write(values[1].Trim() + "\n");

                        xlwb.Worksheets.Add(res, CleanSpreadsheetName(values[1].Trim() + " (" + values[6].Trim() + ")"));
                        Console.ForegroundColor = ConsoleColor.DarkCyan;
                        Console.WriteLine(" > > Writing datatable to: " + values[1].Trim() + " (" + values[6].Trim() + ")");
                        Console.WriteLine(" > > > Department: (" + values[5].Trim() + ") " + values[6].Trim() + " to PCC Code: " + values[8].Trim() + " - " + values[9].Trim());
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                    Services.DrawProgressBar.Draw(csvLineCounter, totalCSVLines, 20, '█');
                    csvLineCounter++;
                }
                
                reader.Close();
                reader.Dispose();
                csvFileCounter++;
            }
            Console.WriteLine("\nOpening Excel ... ");
            if (xlwb.Worksheets.Count != 0)
            {
                xlwb.SaveAs(amalgamatedSpreadsheetName);
                System.Diagnostics.Process.Start(amalgamatedSpreadsheetName);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nDeleting raw csv files ... ");
                Console.ForegroundColor = ConsoleColor.White;
                if (ConfigurationManager.AppSettings["Debug"] == "No")
                {
                    foreach (string csvFileName in csvFileNames)
                    {
                        try
                        {
                            File.Delete(csvFileName);
                        }
                        catch(Exception ex)
                        {
                            logger.Error(ex.Message);
                            logger.Error(ex.StackTrace);
                            Console.ForegroundColor = ConsoleColor.DarkRed;
                            Console.WriteLine("Could not delete file: " + csvFileName);
                        }
                    }
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nExcel file does not exist.");
                Console.ForegroundColor = ConsoleColor.White;
            }
            Console.WriteLine("\nFin.");
        }

        public static string ValidFilePathDate(DateTime dt)
        {
            string result = "";
            result += dt.Hour;
            result += ".";
            result += dt.Minute;
            result += ".";
            result += dt.Second;
            result += " on ";
            result += dt.Day;
            result += ".";
            result += dt.Month;
            result += ".";
            result += dt.Year;
            return result;
        }

        public static System.Data.DataTable ConvertCSVtoDataTable(string strFilePath)
        {
            List<string> headersNullValues = new List<string>();
            headersNullValues.Add("");
            StreamReader sr = new StreamReader(strFilePath);
            System.Data.DataTable dt = new System.Data.DataTable();
            sr.ReadLine();
            if (sr.BaseStream.Length != 0)
            {
                string[] headers = sr.ReadLine().Split(',');
                foreach (string header in headers)
                {
                    dt.Columns.Add(header);
                    /*
                     * Commented out due to an issue with 'Assorted' size items - the price column is throwing a wobbly because its not in double format
                     * Fair enough really but for a quick fix I have wheecht out the offending pixels.
                     * 
                    if (header == "Full Price" || header == "Current Price" || header == "Proposed Price")
                    {
                        dt.Columns.Add(header, typeof(double));
                    }
                    else
                    {
                        dt.Columns.Add(header);
                    }
                    */
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(',');
                    System.Data.DataRow dr = dt.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dt.Rows.Add(dr);
                }
            }
            sr.Dispose();
            return dt;
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
