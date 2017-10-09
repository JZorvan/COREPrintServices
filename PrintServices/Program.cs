using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintServices
{
    class Program
    {
        static void Main(string[] args)
        {
            WordHandler.convertToPdf();
            //PdfHandler.renameFiles();
            //ConsoleHandler.PrintToConsole();
            Console.ReadKey();
        }
    }
}
