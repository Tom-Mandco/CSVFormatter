using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using System.Configuration;

namespace MandCo.CSVFormatter.Applications.Programs
{
    class Al365_Format
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public static void Run(string UniqueBatchNo)
        {
            

            List<string> csvFileNames = new List<string>();
            string csvFilePath = ConfigurationManager.AppSettings["RawReportPath"];
            string outputFilePath = ConfigurationManager.AppSettings["AL365OutputPath"];
            string fileName;
            string departmentBreakdown = "Al365";
            int fileStartIndex;

            string[] files = Directory.GetFiles(@"\\" + csvFilePath, "[Raw]" + "Al365" + "(" + UniqueBatchNo + ")*.csv");
            Console.WriteLine("Running program: Al365 - PCC Report By Packs. URN: " + UniqueBatchNo + "\n\n");
            logger.Info("Starting Al365 Format, Unique batch number: " + UniqueBatchNo);

            XLWorkbook xlwb = new XLWorkbook();

            int csvFileCounter = 0;

            foreach (string file in files)
            {
                int csvLineCounter = 0;
                fileStartIndex = file.IndexOf("\\[Raw]") + 1;
                fileName = file.Substring(fileStartIndex, (file.Length - fileStartIndex));
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(csvFileCounter + 1 + " " + file);
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine(" > Formatting csv file ... ");
                logger.Info("Formatting CSV File: " + fileName);


                try
                {
                    csvFileNames.Add(@"\\" + csvFilePath + fileName);
                }
                catch(Exception ex)
                {
                    logger.Error("Error adding csv file: " + csvFilePath + fileName);
                    logger.Error(ex.Message);
                    logger.Error(ex.StackTrace);
                }

                var reader = new StreamReader(File.OpenRead(@"\\" + csvFilePath + fileName));
                System.Data.DataTable res = ConvertCSVtoDataTable(csvFileNames[csvFileCounter]);

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    if (csvLineCounter == 0)
                    {
                        departmentBreakdown = "(" + values[0];
                        if (values[2] != "")
                            departmentBreakdown += "-" + values[2];

                        if (values[4] != "")
                            departmentBreakdown += "-" + values[4];

                        departmentBreakdown += ") " + values[1];
                        if (values[2] != "")
                            departmentBreakdown += " - " + values[3];

                        if (values[4] != "")
                            departmentBreakdown += " - " + values[5];

                        xlwb.Worksheets.Add(res, (values[0] + " - " + values[1]));
                    }
                    csvLineCounter++;
                }
                reader.Close();
                reader.Dispose();
                csvFileCounter++;
            }
            string amalgamatedSpreadsheetName = (outputFilePath + "(Al365) " + departmentBreakdown + " --- Run by " + Environment.UserName + " at " + ValidFilePathDate(DateTime.Now) + ".xlsx");

            if (xlwb.Worksheets.Count != 0)
            {
                logger.Info("Format completed. Opening Excel");
                Console.WriteLine("\nOpening Excel ... ");
                xlwb.SaveAs(amalgamatedSpreadsheetName);
                System.Diagnostics.Process.Start(amalgamatedSpreadsheetName);
                DeleteRawCSVs(csvFileNames);
            }
            else
            {
                logger.Warn("Excel output file does not exist.");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nExcel file does not exist.");
                Console.ForegroundColor = ConsoleColor.White;
            }
            logger.Info("Completed successfully.");
            Console.WriteLine("\nFin.");
        }

        public static void DeleteRawCSVs(List<string> RawCSVs)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Deleting raw csv files ... ");
            Console.ForegroundColor = ConsoleColor.White;
            foreach (string csvFileName in RawCSVs)
            {
                File.Delete(csvFileName);
            }
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
                    try
                    {
                        dt.Columns.Add(header);
                    }
                    catch(Exception ex)
                    {
                        logger.Error("Two headers cannot be the same. '" + header + "'");
                        logger.Error(ex.Message);
                        logger.Error(ex.StackTrace);
                    }
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
    }
}
