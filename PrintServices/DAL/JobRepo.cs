using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToneDownThatBackEnd.Models;
using Microsoft.Office.Interop.Excel;

namespace PrintServices.DAL
{
    public class JobRepo
    {
        public JobContext Context { get; set; }
        public JobRepo()
        {
            Context = new JobContext();
        }
        public JobRepo(JobContext _context)
        {
            Context = _context;
        }
        public void ImportMasterSpreadsheet()
        {
            Console.WriteLine("I'm trying to read the MasterList.");
            Application excel = new Application();
            //excel.DisplayAlerts = false;
            Workbook workbook = excel.Workbooks.Open("C:/PrintServices\\TestMasterList.xlsx", ReadOnly: false, Editable: false);
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
            if (worksheet == null)
                return;

            Range fileNames = worksheet.Columns["A"];
            Range pageCounts = worksheet.Columns["B"];
            Range PrintQueues = worksheet.Columns["C"];
            Range Boards = worksheet.Columns["D"];
            Range Dispositions = worksheet.Columns["E"];
            Range range = worksheet.UsedRange;
            int totalColumns = worksheet.UsedRange.Columns.Count;
            int totalRows = worksheet.UsedRange.Rows.Count;
            Console.WriteLine(totalColumns);
            Console.WriteLine(totalRows);

            for (int i = 2; i <= totalRows; i++)
            {
                Console.WriteLine(range.Cells[i, 1].Value);
                Console.WriteLine(range.Cells[i, 2].Value);
                Console.WriteLine(range.Cells[i, 3].Value);
                Console.WriteLine(range.Cells[i, 4].Value);
                Console.WriteLine(range.Cells[i, 5].Value);

            }


            workbook.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
        public List<Job> GetJobs()
        {
            int i = 1;
            return Context.Jobs.ToList();
        }
    }
}
