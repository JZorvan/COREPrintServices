using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace PrintServices
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo.remove3502AA();

            List<Task> tasks = new List<Task>()
            {
                new Task(() => WordHandler.convertToPdf()),
                new Task(() => PdfHandler.findDuplicates()),
                new Task(() => PdfHandler.renameFiles()),
                new Task(() => ConsoleHandler.printToConsole())
            };

            foreach (Task t in tasks)
            {
                t.Start();
                t.Wait();
            }

            Console.ReadKey();
        }
    }
}
