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
        public static void convertToPdf()
        {
            var wordFiles = Directory.EnumerateFiles("C:/PrintServices", "*.docx");
            int i = 0;
            foreach (string file in wordFiles)
            {
                if (Path.GetFileName(file).StartsWith("~") == false)
                {
                    Word2Pdf converter = new Word2Pdf();
                    converter.InputLocation = file;
                    converter.OutputLocation = Path.ChangeExtension(file, ".pdf");
                    converter.Word2PdfCOnversion();
                    i++;
                    File.Delete(file);
                }
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\r\n" + i + " files have been converted to Pdf format.\r\n");
        }
    }
}
