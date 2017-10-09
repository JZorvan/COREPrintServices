using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintServices
{
    class ConsoleHandler
    {
        public static void PrintToConsole()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.pdf"))
            {
                if (Total.TwoSidedFiles.Contains(file))
                {
                    Console.WriteLine(Filename.Trimmer(file) + " - " + (PdfReader.getNumberOfPages(file) / 2));
                }
                else
                {
                    Console.WriteLine(Filename.Trimmer(file) + " - " + PdfReader.getNumberOfPages(file));
                }
                Total.AddToTotalCount(file, PdfReader.getNumberOfPages(file));
            }
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Total Entries: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(Total.totalDocCount);
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("Total Pages: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(Total.totalPageCount);
        }
    }
}
