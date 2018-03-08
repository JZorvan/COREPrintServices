using System;
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
        public static string filenameTrimmer(string file) // Helper method to remove portions of a filename
        {
            char[] pdfString = { '.', 'd', 'o', 'c', 'x'};
            string trimmedFilename = "";

            trimmedFilename = file.Substring(17).TrimEnd(pdfString);

            return trimmedFilename;
        }
    }

}
