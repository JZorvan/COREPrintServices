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
            List<Task> tasks = new List<Task>();
            Task a = new Task(() => { WordHandler.convertToPdf(); });
            Task b = new Task(() => { PdfHandler.renameFiles(); });
            Task c = new Task(() => { ConsoleHandler.printToConsole(); });
            tasks.Add(a);
            tasks.Add(b);
            tasks.Add(c);
            foreach (Task t in tasks)
            {
                t.Start();
                t.Wait();
            }
            Console.ReadKey();
        }
    }
}
