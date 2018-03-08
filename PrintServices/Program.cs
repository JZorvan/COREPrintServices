using System.Collections.Generic;
using System.Threading.Tasks;
using static PrintServices.ExcelHandler;

namespace PrintServices
{
    class Program
    {
        static void Main(string[] args)
        {
            ConsoleHandler.Print("greeting");

            Task TransferFiles = new Task(() => FileInfo.transferFiles());
            TransferFiles.Start();
            TransferFiles.Wait();

            List<KeyValuePair<string, ValuePair>> jobsAndCounts = ReadSpreadsheet();

            FileInfo.removeFilesToDelete();

            //FileInfo.checkForUnknowns(jobsAndCounts);

            Task ConvertToPdf = new Task(() => WordHandler.convertToPdf());
            ConvertToPdf.Start();
            ConvertToPdf.Wait();

            Task handleDuplicates = new Task(() => PdfHandler.handleDuplicates(jobsAndCounts));
            handleDuplicates.Start();
            handleDuplicates.Wait();

            Task RenameFiles = new Task(() => PdfHandler.renameFiles(jobsAndCounts));
            RenameFiles.Start();
            RenameFiles.Wait();

            FileInfo.checkForUnknowns(jobsAndCounts);

            jobsAndCounts = Counter.assignSheetCount(jobsAndCounts);
            jobsAndCounts = Counter.UpdatePageCount(jobsAndCounts);

            ConsoleHandler.Print("counted");
            
            Task PopulateSpreadsheet = new Task(() => populateSpreadsheet(jobsAndCounts));
            PopulateSpreadsheet.Start();
            PopulateSpreadsheet.Wait();

            Task GenerateBatchFile = new Task(() => BatchHandler.generateBatchFile(jobsAndCounts));
            GenerateBatchFile.Start();
            GenerateBatchFile.Wait();

            ConsoleHandler.Print("done");
            ConsoleHandler.Print("exit");
        }
    }
}
