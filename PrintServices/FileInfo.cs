using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace PrintServices
{
    class FileInfo
    {
        public static void removeFilesToDelete()
        {
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + @"\Application Files\JobsToDelete.xlsx", ReadOnly: true, Editable: false);
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
            if (worksheet == null)
                return;

            Range range = worksheet.UsedRange;
            int totalRows = worksheet.UsedRange.Rows.Count;
            string pattern;
            for (int i = 1; i <= totalRows; i++)
            {
                pattern = range.Cells[i, 1].Value;
                foreach (string file in Directory.EnumerateFiles(Environment.CurrentDirectory, pattern))
                {
                    File.Delete(file);
                }
            }
            workbook.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
    }
}
