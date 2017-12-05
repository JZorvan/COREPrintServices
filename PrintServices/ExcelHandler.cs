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
        public static string getSpreadsheetName()
        {
            string today = DateTime.Now.ToString("MM.dd.yyyy");
            
            string filename = today + " MasterList.xlsm";
            Console.WriteLine(filename);
            return filename;
        }
        public static void populateSpreadsheet(List<Job> jobs)
        {
            Application excel = new Application();
            excel.DisplayAlerts = false;
            Workbook workbook = excel.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + @"\Application Files\MasterList.xlsm", ReadOnly: false, Editable: true);
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

            workbook.SaveAs(Filename: "c:\\PrintServices\\" + getSpreadsheetName());

            workbook.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
    }
}
