using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using PrintServices.DAL;
using PrintServices.Models;

namespace PrintServices
{
    class Program
    {
        static void Main(string[] args)
        {
            JobRepo db = new JobRepo();
            db.ImportMasterSpreadsheet();
            List<Job> jobs = db.GetJobs();
            foreach (Job job in jobs)
            {
                Console.WriteLine(job.FileName);
            }
            db.ClearRepository(jobs);
            Console.WriteLine(jobs.Count);




                //Notes on making the path automatic for other users:
                //string path = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory);
                //string filepath = Path.GetFullPath(Path.Combine(path, ".."));
                //Console.WriteLine(filepath);

                //FileInfo.remove3502AA();

                //List<Task> tasks = new List<Task>()
                //{
                //    new Task(() => WordHandler.convertToPdf()),
                //    new Task(() => PdfHandler.findDuplicates()),
                //    new Task(() => PdfHandler.renameFiles()),
                //    new Task(() => ExcelHandler.populateSpreadsheet()),
                //    new Task(() => ConsoleHandler.printToConsole())
                //};

                //foreach (Task t in tasks)
                //{
                //    t.Start();
                //    t.Wait();
                //}

                Console.ReadKey();

        }
    }
}
