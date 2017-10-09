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
                Counter.AddToTotalCount(file, Counter.getNumberOfPages(file));
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
