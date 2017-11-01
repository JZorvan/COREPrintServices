using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

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
        public static void populateSpreadsheet()
        {
            Console.WriteLine("Starting to populate spreadsheet.");
            Application excel = new Application();
            excel.DisplayAlerts = false;
            Workbook workbook = excel.Workbooks.Open("C:/PrintServices\\MasterList.xlsm", ReadOnly: false, Editable: true);
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
            if (worksheet == null)
                return;

            Dictionary<string, string> countDictionary = new Dictionary<string, string>();
            string fileName = "";
            string pageCount = "";
            foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.pdf"))
            {
                fileName = PdfHandler.filenameTrimmer(file);

                if (FileInfo.TwoSidedFiles.Contains(file))
                {
                    pageCount = (Counter.getNumberOfPages(file) / 2).ToString();
                }
                else
                {
                    pageCount = (Counter.getNumberOfPages(file)).ToString();
                }

                countDictionary.Add(fileName, pageCount);
                Counter.AddToTotalCount(file, Counter.getNumberOfPages(file));
            }

            Range jobColumn = worksheet.Columns["A"];
            Range foundJob = null;
            foreach (KeyValuePair<string, string> kvp in countDictionary)
            {
                foundJob = jobColumn.Find(kvp.Key);
                int rowNum = Convert.ToInt32(foundJob.Address.Substring(3));
                worksheet.Cells[rowNum, 2] = kvp.Value;
            }

            workbook.SaveAs(Filename: "c:\\PrintServices\\" + getSpreadsheetName());

            excel.DisplayAlerts = true;
            workbook.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
    }
}
