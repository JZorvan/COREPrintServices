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
            JobRepo db = new JobRepo();
            Task ImportMasterSpreadsheet = new Task(() => db.ImportMasterSpreadsheet());
            Task ConvertToPdf = new Task(() => WordHandler.convertToPdf());
            
            Task PopulateSpreadsheet = new Task(() => ExcelHandler.populateSpreadsheet());
            Task PrintToConsole = new Task(() => ConsoleHandler.printToConsole());

            ImportMasterSpreadsheet.Start();
            ImportMasterSpreadsheet.Wait();
            List<Job> jobs = db.GetJobs();
            FileInfo.removeFilesToDelete();
            ConvertToPdf.Start();
            ConvertToPdf.Wait();
            Task handleDuplicates = new Task(() => PdfHandler.handleDuplicates(jobs));
            handleDuplicates.Start();
            handleDuplicates.Wait();
            Task RenameFiles = new Task(() => PdfHandler.handleRenaming(jobs));
            RenameFiles.Start();
            RenameFiles.Wait();


            db.ClearRepository();

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
