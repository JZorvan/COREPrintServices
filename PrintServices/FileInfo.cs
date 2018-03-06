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
