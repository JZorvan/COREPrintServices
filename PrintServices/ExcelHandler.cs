using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace PrintServices
{
    class ExcelHandler
    {
        public struct ValuePair
        {
            public int pageCount;
            public string printQueue;
        }

        public static List<KeyValuePair<string, ValuePair>> ReadSpreadsheet()
        {
            ConsoleHandler.Print("import");

            List<KeyValuePair<string, ValuePair>> JobCountPairs = new List<KeyValuePair<string, ValuePair>>();

            Application excel = new Application();
            excel.DisplayAlerts = false;
            try
            {
                Workbook workbook = excel.Workbooks.Open(@"C:\PrintServices\Application Files\MasterList.xlsx", ReadOnly: true, Editable: false);
                Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
                if (worksheet == null)
                    return JobCountPairs = null;

                Range range = worksheet.UsedRange;
                int totalRows = worksheet.UsedRange.Rows.Count;

                for (int i = 2; i <= totalRows; i++)
                {
                    string jobName = range.Cells[i, 1].Value.Trim();
                    int pageCount = Convert.ToInt32(range.Cells[i, 2].Value);
                    string printQueue = range.Cells[i, 3].Value.Trim();
                    KeyValuePair<string, ValuePair> JobCountPair = new KeyValuePair<string, ValuePair>(jobName, new ValuePair { pageCount = pageCount, printQueue = printQueue });
                    JobCountPairs.Add(JobCountPair);
                }
                workbook.Close();
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                ConsoleHandler.Print("importsuccess");
                return JobCountPairs;
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine();
                Console.WriteLine();
                ConsoleHandler.Print("error");
                Console.WriteLine(e.Message);
                Console.WriteLine();
                ConsoleHandler.Print("exit");
                return JobCountPairs = null;
            }
        }
        public static string getSpreadsheetName()  // Gets today's date in the desired format
        {
            string today = DateTime.Now.ToString("MM.dd.yyyy");
            
            string filename = today + " MasterList.xlsx";
            return filename;
        }
        public static void populateSpreadsheet(List<KeyValuePair<string, ValuePair>> jobs) // Writes the calcuted pagecount to a new copy of the MasterList for today's jobs
        {
            ConsoleHandler.Print("spreadsheet");
            Application excel = new Application();
            excel.DisplayAlerts = false;
            try
            {
                Workbook workbook = excel.Workbooks.Open(@"C:\PrintServices\Application Files\MasterList.xlsx", ReadOnly: false, Editable: true);
                Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
                if (worksheet == null)
                    return;

                Range nameColumn = worksheet.Columns["A"];
                Range countColumn = worksheet.Columns["B"];
                Range foundJob = null;

                foreach (KeyValuePair<string, ValuePair> job in jobs)
                {
                    foundJob = nameColumn.Find(job.Key);
                    int rowNum = Convert.ToInt32(foundJob.Address.Substring(3));
                    worksheet.Cells[rowNum, 2] = job.Value.pageCount;
                }

                string filename = @"C:\PrintServices\" + getSpreadsheetName();

                if (File.Exists(filename))
                {
                    File.Delete(filename);
                }

                workbook.SaveAs(Filename: filename);

                workbook.Close();
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                ConsoleHandler.Print("spreadsheeted");
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine();
                Console.WriteLine();
                ConsoleHandler.Print("error");
                Console.WriteLine(e.Message);
                Console.WriteLine();
                ConsoleHandler.Print("exit");
            }
        }
    }
}
