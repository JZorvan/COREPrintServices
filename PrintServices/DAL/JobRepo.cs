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
        public void ImportMasterSpreadsheet()  // Pulls the information from the Master spreadsheet and to build the database
        {
            ConsoleHandler.Print("import");
            Application excel = new Application();
            excel.DisplayAlerts = false;
            try
            {
                Workbook workbook = excel.Workbooks.Open(@"C:\PrintServices\Application Files\MasterList.xlsx", ReadOnly: true, Editable: false);
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
                ConsoleHandler.Print("importsuccess");
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
        public List<Job> GetJobs()  // Pulls the jobs from the repo as a List
        {
            return Context.Jobs.ToList();
        }
        public void UpdatePageCount(string filename, int pagecount)  // Divides the pagecount by two if it is supposed to be a double-sided job
        {
            Job foundJob = Context.Jobs.SingleOrDefault(j => j.FileName == filename);
            if (foundJob.PrintQueue == "ce5793_RB-2_pdf" ||
                foundJob.PrintQueue == "ce5786_RB-2_pdf" ||
                foundJob.PrintQueue == "ce5793_FIRE-2_pdf")
            {
                foundJob.PageCount = pagecount/2;
            } else
            {
                foundJob.PageCount = pagecount;
            }
            Context.SaveChanges();
        }
        public void ClearRepository()  // The repo is cleared at the beginning and end to avoid residual data causing errors
        {
            if (GetJobs().Count() > 0)  
            {
                Context.Jobs.RemoveRange(Context.Jobs);  
                Context.SaveChanges();
                ConsoleHandler.Print("clear");  
            }
        }
    }
}
