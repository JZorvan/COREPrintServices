﻿using System;
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
            PdfHandler.renameFiles();
            ConsoleHandler.PrintToConsole();
            Console.Read();
        }
    }
}
