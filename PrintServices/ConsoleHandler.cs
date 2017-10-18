using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintServices
{
    class ConsoleHandler
    {
        public static void printToConsole()
        {
            //foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.pdf"))
            //{
            //    string fileName = PdfHandler.filenameTrimmer(file);
            //    string pageCount = "";
            //    Dictionary<string, string> countDictionary = new Dictionary<string, string>();
            //    if (FileInfo.TwoSidedFiles.Contains(file))
            //    {
            //        pageCount = (Counter.getNumberOfPages(file) / 2).ToString();
            //    }
            //    else
            //    {
            //        pageCount = (Counter.getNumberOfPages(file) / 2).ToString();
            //    }
            //    countDictionary.Add(fileName, pageCount);
            //    Counter.AddToTotalCount(file, Counter.getNumberOfPages(file));
            //}
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Cyan;
            foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.pdf"))
            {
                if (FileInfo.TwoSidedFiles.Contains(file))
                {
                    Console.WriteLine(PdfHandler.filenameTrimmer(file) + " - " + (Counter.getNumberOfPages(file) / 2));
                }
                else
                {
                    Console.WriteLine(PdfHandler.filenameTrimmer(file) + " - " + Counter.getNumberOfPages(file));
                }
                //Counter.AddToTotalCount(file, Counter.getNumberOfPages(file));
            }
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Total Entries: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(Counter.totalDocCount);
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Total Pages: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(Counter.totalPageCount);
        }
    }
}
