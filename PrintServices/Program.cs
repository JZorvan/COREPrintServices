using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using PrintServices.DAL;
using PrintServices.Models;
using System.Data.Entity;

namespace PrintServices
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Task> tasks = new List<Task>()
            {
                new Task(() => WordHandler.convertToPdf()),
                new Task(() => PdfHandler.findDuplicates()),
                new Task(() => PdfHandler.renameFiles()),
                new Task(() => ExcelHandler.populateSpreadsheet()),
                new Task(() => ConsoleHandler.printToConsole())
            };

            FileInfo.removeFilesToDelete();
            tasks[0].Start();
                tasks[0].Wait();
            //JobRepo db = new JobRepo();
            //db.ImportMasterSpreadsheet();
            //List<Job> jobs = db.GetJobs();
            //db.ClearRepository();

            //Notes on making the path automatic for other users:
            //string path = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory);
            //string filepath = Path.GetFullPath(Path.Combine(path, ".."));
            //Console.WriteLine(filepath);



            //foreach (Task t in tasks)
            //{
            //    t.Start();
            //    t.Wait();
            //}

            Console.ReadKey();

        }
    }
}
