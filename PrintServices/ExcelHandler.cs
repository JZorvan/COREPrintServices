using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using PrintServices.Models;

namespace PrintServices
{
    class ExcelHandler
    {
        public static string getSpreadsheetName()  // Gets today's date in the desired format
        {
            string today = DateTime.Now.ToString("MM.dd.yyyy");
            
            string filename = today + " MasterList.xlsx";
            return filename;
        }
        public static void populateSpreadsheet(List<Job> jobs) // Writes the calcuted pagecount to a new copy of the MasterList for today's jobs
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

                foreach (Job job in jobs)
                {
                    foundJob = nameColumn.Find(job.FileName);
                    int rowNum = Convert.ToInt32(foundJob.Address.Substring(3));
                    worksheet.Cells[rowNum, 2] = job.PageCount;
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
