using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;
using SautinSoft;
using WordToPDF;

namespace PrintServices
{
    class WordHandler
    {
        public static void convertToPdf()
        {
            int i = 0;
            foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.docx"))
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
            Console.WriteLine(i + " files have been converted to Pdf format.");
        }
    }
}
