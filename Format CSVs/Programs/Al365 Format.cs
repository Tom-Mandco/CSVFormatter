using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using System.Configuration;
using MandCo.CSVFormatter;

namespace MandCo.CSVFormatter.Applications.Programs
{
    class Al365_Format
    {
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public static void Run(string UniqueBatchNo)
        {
            string csvFileName, fileName;
            
            logger.Debug("Aquiring csv & output path");
            string csvFilePath = ConfigurationManager.AppSettings["RawReportPath"];
            string outputFilePath = ConfigurationManager.AppSettings["AL365OutputPath"];
            logger.Debug("Successful");

            string departmentBreakdown = "Al365";
            int fileStartIndex;

            logger.Debug("Finding list of csv files");
                string[] files = Directory.GetFiles(@"\\" + csvFilePath, "[Raw]" + "Al365" + "(" + UniqueBatchNo + ")*.csv");
                logger.Debug("Found " + files.Count() + " files");
                string file = files.First();
            Console.WriteLine("Running program: Al365 - PCC Report By Packs. URN: " + UniqueBatchNo + "\n\n");
            logger.Info("Starting Al365 Format, Unique batch number: " + UniqueBatchNo);

            XLWorkbook xlwb = new XLWorkbook();

            int csvLineCounter = 0;

            fileStartIndex = file.IndexOf("\\[Raw]") + 1;
            fileName = file.Substring(fileStartIndex, (file.Length - fileStartIndex));

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("1. " + file);
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(" > Formatting csv file ... \n");
            logger.Info("Formatting CSV File: " + fileName);

            csvFileName = (@"\\" + csvFilePath + fileName);

            logger.Debug("Attempting to open streamreader for file: " + csvFileName);
            var reader = new StreamReader(File.OpenRead(@"\\" + csvFilePath + fileName));
            logger.Debug("Reading ... ");

            logger.Debug("Converting csv data to Datatable");
            System.Data.DataTable res = ConvertCSVtoDataTable(csvFileName);
            int totalCSVLines = Services.Common.CSVRowCount(csvFileName);
            logger.Debug("Successful");

            logger.Debug("Getting Spreadsheet info");
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

                    logger.Debug("Successful");
                    logger.Debug("Creating new worksheet");
                    xlwb.Worksheets.Add(res, Services.Common.CleanSpreadsheetName(values[0] + " - " + values[1]));
                    logger.Debug("Successful");
                }
                Services.DrawProgressBar.Draw(csvLineCounter, totalCSVLines, 20, '█');
                csvLineCounter++;
            }
            logger.Debug("Successful");

            logger.Debug("Closing reader");
            reader.Close();
            reader.Dispose();

            logger.Debug("Creating spreadsheet name");
            string amalgamatedSpreadsheetName = (outputFilePath + "(Al365) " + departmentBreakdown + " --- Run by " + Environment.UserName + " at " + ValidFilePathDate(DateTime.Now) + ".xlsx");
            logger.Debug("Successful");
            if (xlwb.Worksheets.Count != 0)
            {
                logger.Info("Format completed. Opening Excel");
                Console.WriteLine("\n\nOpening Excel ... ");
                xlwb.SaveAs(amalgamatedSpreadsheetName);
                System.Diagnostics.Process.Start(amalgamatedSpreadsheetName);
                logger.Debug("Successful");
                DeleteRawCSV(csvFileName);
            }
            else
            {
                logger.Warn("Excel output file does not exist.");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\nExcel file does not exist.");
                Console.ForegroundColor = ConsoleColor.White;
            }
            logger.Info("Completed Successfully.");
            Console.WriteLine("\nFin.");
        }

        public static void DeleteRawCSV(string RawCSV)
        {
            logger.Info("Deleting raw csv files");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Deleting raw csv file ... ");
            Console.ForegroundColor = ConsoleColor.White;
            File.Delete(RawCSV);
            logger.Debug("Successful");
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
            int rowCount = 0;
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
                    catch (Exception ex)
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
                        rowCount++;
                        dr[i] = rows[i];
                    }
                    rowCount++;
                    dt.Rows.Add(dr);
                }
            }
            sr.Dispose();
            return dt;
        }
    }
}