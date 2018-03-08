using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using static PrintServices.ExcelHandler;

namespace PrintServices
{
    class FileInfo
    {
        public static void transferFiles()  // Prompts User to ready files for processing
        {
            ConsoleHandler.Print("question");
            string response = Console.ReadLine();
            checkFilePresence();
        
        }
        public static void checkFilePresence() // Checks there are files ready to be processed
        {
            List<string> letterFiles = new List<string>();
            foreach (string file in Directory.EnumerateFiles(@"C:\PrintServices", "*.pdf"))
            {
                if (Path.GetFileName(file).StartsWith("~") == false)
                {
                    letterFiles.Add(file);
                }
            }
            foreach (string file in Directory.EnumerateFiles(@"C:\PrintServices", "*.doc"))
            {
                if (Path.GetFileName(file).StartsWith("~") == false)
                {
                    letterFiles.Add(file);
                }
            }
            if (letterFiles.Count == 0)
            {
                ConsoleHandler.Print("liar");
            }

        }
        public static void checkForUnknowns(List<KeyValuePair<string, ValuePair>> jobs)  // Checks if there are files that aren't referenced on the Masterlist
        {
            List<string> acceptedFiles = new List<string>();
            foreach (KeyValuePair<string, ValuePair> job in jobs)
            {
                acceptedFiles.Add(job.Key);
            }

            List<string> files = PdfHandler.getFiles();
            List<string> unknownFiles = new List<string>();
            foreach (string file in files)
            {
                if (!acceptedFiles.Contains(PdfHandler.filenameTrimmer(file) + ".pdf") && !PdfHandler.filenameTrimmer(file).StartsWith("~"))
                {
                    unknownFiles.Add(PdfHandler.filenameTrimmer(file) + ".pdf");
                }
            }
            foreach (string file in Directory.EnumerateFiles(@"C:\PrintServices", "*.docx"))
            {
                if (!acceptedFiles.Contains(WordHandler.filenameTrimmer(file) + ".pdf") && !WordHandler.filenameTrimmer(file).StartsWith("~"))
                {
                    unknownFiles.Add(file.Substring(17));
                }
            }
            if (unknownFiles.Count > 0)
            {
                ConsoleHandler.Print("error");
                Console.WriteLine(" {0} files were found that are not referenced on the Master Spreadsheet:", unknownFiles.Count);
                foreach (string file in unknownFiles)
                {
                    Console.WriteLine("     {0}", file);
                }
                ConsoleHandler.Print("exit");
            }
        }
        public static void removeFilesToDelete() // Read a spreadsheet of file patterns that need to be deleted and deletes them
        {
            Application excel = new Application();
            excel.DisplayAlerts = false;
            try
            {
                ConsoleHandler.Print("finddeletes");
                Workbook workbook = excel.Workbooks.Open(@"C:\PrintServices\Application Files\JobsToDelete.xlsx", ReadOnly: true, Editable: false);
                Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
                if (worksheet == null)
                    return;

                Range range = worksheet.UsedRange;
                int totalRows = worksheet.UsedRange.Rows.Count;
                string pattern;
                List<string> filesToDelete = new List<string>();

                for (int i = 1; i <= totalRows; i++)
                {
                    pattern = range.Cells[i, 1].Value;
                    foreach (string file in Directory.EnumerateFiles(@"C:\PrintServices", pattern))
                    {
                        filesToDelete.Add(file);
                    }
                }

                if (filesToDelete.Count > 0)
                {
                    Console.WriteLine();
                    foreach (string file in filesToDelete)
                    {
                        File.Delete(file);
                        Console.Write(file);
                        ConsoleHandler.Print("deleted");
                    }
                    Console.WriteLine();
                }
                else
                {
                    ConsoleHandler.Print("none");
                }

                workbook.Close();
                excel.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
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
    }
}
