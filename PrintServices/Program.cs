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
using static System.Net.Mime.MediaTypeNames;

namespace PrintServices
{
    class Program
    {
        static void Main(string[] args)
        {
            ConsoleHandler.Print("greeting");

            JobRepo db = new JobRepo(new JobContext());
            db.ClearRepository();

            Task ImportMasterSpreadsheet = new Task(() => db.ImportMasterSpreadsheet());
            Task ConvertToPdf = new Task(() => WordHandler.convertToPdf());
            Task TransferFiles  = new Task(() => FileInfo.transferFiles());
            TransferFiles.Start();
            TransferFiles.Wait();
            ImportMasterSpreadsheet.Start();
            ImportMasterSpreadsheet.Wait();
            List<Job> jobs = db.GetJobs();
            FileInfo.removeFilesToDelete();
            ConvertToPdf.Start();
            ConvertToPdf.Wait();
            Task handleDuplicates = new Task(() => PdfHandler.handleDuplicates(jobs));
            handleDuplicates.Start();
            handleDuplicates.Wait();
            Task RenameFiles = new Task(() => PdfHandler.renameFiles(jobs));
            RenameFiles.Start();
            RenameFiles.Wait();
            jobs = Counter.assignSheetCount(jobs);

            foreach (Job job in jobs)
            {
                if (job.PageCount != 0)
                {
                    db.UpdatePageCount(job.FileName, job.PageCount);
                }
            }
            ConsoleHandler.Print("counted");
            jobs = db.GetJobs();
            Task PopulateSpreadsheet = new Task(() => ExcelHandler.populateSpreadsheet(jobs));
            PopulateSpreadsheet.Start();
            PopulateSpreadsheet.Wait();
            Task GenerateBatchFile = new Task(() => BatchHandler.generateBatchFile(jobs));
            GenerateBatchFile.Start();
            GenerateBatchFile.Wait();

            db.ClearRepository();
            ConsoleHandler.Print("done");
            ConsoleHandler.Print("exit");
        }
    }
}
