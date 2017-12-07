using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintServices
{
    class ConsoleHandler
    {
        public static void Print(string input)
        {
            switch (input)
            {
                case "greeting":
                    Console.WriteLine("Welcome to the Print and Mail Services Application.");
                    break;
                case "no files":
                    Console.WriteLine("There are no files in your folder.");
                    Console.WriteLine("Please copy today's files from JBoss to your PrintServices folder.");
                    break;
                
            }
        }
    }
}
