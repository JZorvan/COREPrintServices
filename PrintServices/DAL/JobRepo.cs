using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PrintServices.Models;
using Microsoft.Office.Interop.Excel;
using System.Data.Entity;

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
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open("C:/PrintServices\\TestMasterList.xlsx", ReadOnly: true, Editable: false);
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
            if (worksheet == null)
                return;

            Range range = worksheet.UsedRange;
            int totalRows = worksheet.UsedRange.Rows.Count;

            for (int i = 2; i <= totalRows; i++)
            {
                Job job = new Job();
                job.FileName = range.Cells[i, 1].Value.Trim();
                job.PageCount = Convert.ToInt32(range.Cells[i, 2].Value);
                job.PrintQueue = range.Cells[i, 3].Value.Trim();
                job.Board = range.Cells[i, 4].Value.Trim();
                job.Disposition = range.Cells[i, 5].Value.Trim();
                Context.Jobs.Add(job);
                Context.SaveChanges();    
            }
            workbook.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
        public List<Job> GetJobs()
        {
            //int i = 1;
            return Context.Jobs.ToList();
        }
        public void ClearRepository()
        {
            Context.Jobs.RemoveRange(Context.Jobs);
            Context.SaveChanges();
        }
    }
}
