using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using WordToPDF;

namespace PrintServices
{
    class WordHandler
    {
        public static void convertToPdf()  // Converts Word files to PDF files
        {
            ConsoleHandler.Print("wordfiles");
            try
            {
                var wordFiles = Directory.EnumerateFiles(@"C:\PrintServices", "*.docx");
                foreach (string file in wordFiles)
                {
                    if (Path.GetFileName(file).StartsWith("~") == false)
                    {
                        Word2Pdf converter = new Word2Pdf();
                        converter.InputLocation = file;
                        converter.OutputLocation = Path.ChangeExtension(file, ".pdf");
                        converter.Word2PdfCOnversion();
                        File.Delete(file);
                    }
                }
                ConsoleHandler.Print("wordfilessuccess");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine();
                Console.WriteLine();
                ConsoleHandler.Print("error");
                ConsoleHandler.Print("open");
                ConsoleHandler.Print("exit");
            }
        }
    }
}
