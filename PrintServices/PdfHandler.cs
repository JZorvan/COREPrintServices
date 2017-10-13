using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintServices
{
    class PdfHandler
    {
        public static void findDuplicates()
        {
            string[] pdfFiles = Directory.GetFiles("C:/PrintServices", "*.pdf");
            foreach (string filename in FileInfo.allFiles)
            {
                var result = Array.FindAll(pdfFiles, s => s.Contains(filename));
                if (result.Count() > 1)
                {
                    Console.WriteLine(filename + " has " + result.Count() + " files that need to be combined.");
                } 
            }
        }
        public static void combinePdfs()
        {

        }
        public static void renameFiles()
        {
            foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.pdf"))
            {
                File.Move(file, getNewFilename(file));
            }
        }
        public static string getNewFilename(string file)
        {
            if (file.Contains("lp32-15C"))
            {
                return file.Substring(0, 25) + ".pdf";
            }
            else if (file.Contains("lp32-19C"))
            {
                return file.Substring(0, 25) + ".pdf";
            }
            else if (file.Contains("lp59-15C"))
            {
                return file.Substring(0, 25) + ".pdf";
            }
            else if (file.Contains("lp59-1601B"))
            {
                return file.Substring(0, 27) + ".pdf";
            }
            else if (file.Contains("lp59-3502B"))
            {
                return file.Substring(0, 27) + ".pdf";
            }
            else if (file.Contains("25CAN001"))
            {
                return "C:/PrintServices\\25CAN001.pdf";
            }
            else if (file.Contains("25CAA001"))
            {
                return "C:/PrintServices\\25CAA001.pdf";
            }
            else if (file.Contains("tndc06"))
            {
                return file.Substring(0, 28) + ".pdf";
            }
            else
            {
                return file.Substring(0, 26) + ".pdf";
            }
        }
        public static string filenameTrimmer(string file)
        {
            char[] pdfString = { '.', 'p', 'd', 'f' };
            string trimmedFilename = "";

            trimmedFilename = file.Substring(17).TrimEnd(pdfString);

            return trimmedFilename;
        }
    }
}
