using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PrintServices
{
    class ConsoleHandler
    {
        public static void Print(string input)
        {
            switch (input)
            {
                case "greeting":  // Program Start
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("      ____ Welcome to the Print and Mail Services Application ____");
                    Console.WriteLine();
                    break;
                case "question":  // Prompts the User to double check that they have the right files ready
                    Console.BackgroundColor = ConsoleColor.Magenta;
                    Console.ForegroundColor = ConsoleColor.Black;
                    Console.WriteLine(" Before continuing, please make sure you have done the following: ");
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine(@"     1. Ensure you have deleted any Generated Letters in C:\PrintServices that don't need to be sent out today.");
                    Console.WriteLine(@"     2. Get any Generated Letters that need to be processed today from JBoss and add them to C:\PrintServices.");
                    Console.WriteLine(@"     3. Close any of those Generated Letters that you may have opened.");
                    Console.WriteLine();
                    Console.WriteLine(" Have you done both of these?  Press Enter to continue..");
                    Console.ForegroundColor = ConsoleColor.White;
                    break;
                case "liar":  // Tells User that no applicable files were found and closes program.
                    Console.BackgroundColor = ConsoleColor.Red;
                    Console.WriteLine(" No .pdf or .doc files were found! ");
                    Console.WriteLine(" Please restart the application after you have completed the above steps. ");
                    Thread.Sleep(10000);
                    Environment.Exit(0);
                    break;
                case "clear":  // ClearRepository() action
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(" The database has been cleared of any residual data.");
                    Console.WriteLine();
                    break;
                case "import":  // ImportMasterSpreadsheet() start
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.Write(" Reading and importing the data from the MasterList...");
                    break;
                case "importsuccess":  // ImportMasterSpreadsheet() success
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(" Masterlist was successfully imported!");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine();
                    Console.WriteLine();
                    break;
                case "finddeletes": // removeFilesToDelete() start
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.Write(" Checking to see if any of these files need to be deleted...");
                    break;
                case "deleted": // removeFilesToDelete() successful, files deleted
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write(" was deleted.  We don't process that one.");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine();
                    break;
                case "none": // removeFilesToDelete() successful, no files to delete
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(" but there aren't any of those today.");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine();
                    Console.WriteLine();
                    break;
                case "wordfiles":  // convertToPdf() start
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.Write(" Converting Word files into PDFs...");
                    break;
                case "wordfilessuccess": // convertToPdf() success
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(" Complete!");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine();
                    Console.WriteLine();
                    break;
                case "duplicates":  // handleDuplicates() start
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine(" Combining multiple PDFs...");
                    Console.WriteLine();
                    break;
                case "duplicatessuccess": // handleDuplicates() success
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(" ..File Merge Complete!");
                    Console.WriteLine();
                    break;
                case "rename":  // renameFiles() start
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.Write(" Renaming files in the simplified format...");
                    break;
                case "renamed": // renameFiles() success
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(" Complete!");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine();
                    Console.WriteLine();
                    break;
                case "count":  // assignSheetCount() start
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.Write(" Counting the number of pages in each file...");
                    break;
                case "counted": // UpdatePageCount() success
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(" Complete!");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine();
                    Console.WriteLine();
                    break;
                case "batch":  // generateBatchFile() start
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.Write(" Creating a batch file for you to run...");
                    break;
                case "batched": // generateBatchFile() success
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.Write(" Complete!");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine();
                    Console.WriteLine();
                    break;
                case "done":
                    Console.BackgroundColor = ConsoleColor.Magenta;
                    Console.ForegroundColor = ConsoleColor.Black;
                    Console.WriteLine(" The CORE PrintServices Application has completed successfully! ");
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine(@"     Next - You should do the following: ");
                    Console.WriteLine(@"     1. Run the 'print.bat' file that generated in C:\PrintServices.");
                    Console.WriteLine(@"     2. Email today's 'MasterList.xlsx' to the Print Services, Postal, and MailRoom groups.");
                    Console.WriteLine(@"     3. Once the batch has successfully completed, you can delete today's PDF files.");
                    Console.WriteLine();
                    break;
                case "open":  // Program Error occuring because a file is in use
                    Console.WriteLine(" The file the Application is trying to manipulate is already open.  Please close it before restarting the Application.");
                    Console.WriteLine();
                    break;
                case "error":  // Prefaces exceptions, directing the user to the guide
                    Console.BackgroundColor = ConsoleColor.Red;
                    Console.WriteLine(" See the error message below and refer to your PrintServices guide before restarting the Application: ");
                    Console.WriteLine();
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
                case "exit":  // Informs User they can Exit
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine(" Press any key to exit..");
                    Console.ReadKey();
                    Environment.Exit(0);
                    break;
                //case "spreadsheet":
                //    Console.ForegroundColor = ConsoleColor.White;
                //    Console.WriteLine("Today's MasterList spreadsheet has been created.");
                //    Console.WriteLine("Please email this file when ready.");
                //    Console.WriteLine();
                //    break;
                //case "batch":
                //    Console.ForegroundColor = ConsoleColor.White;
                //    Console.WriteLine(@"The batch file, 'print.bat', has been created.");
                //    Console.WriteLine("Please run the batch file when ready.");
                //    Console.WriteLine();
                //    break;
            }
        }
    }
}
