using System.Collections.Generic;
using System.IO;
using System.Text;
using static PrintServices.ExcelHandler;

namespace PrintServices
{
    class BatchHandler
    {
        public static void generateBatchFile(List<KeyValuePair<string, ValuePair>> jobs)
        {
            ConsoleHandler.Print("batch");
            string batchFile = @"C:\PrintServices\print.bat";
            
            if (File.Exists(batchFile))
            {
                File.Delete(batchFile);
            }

            using (StreamWriter writer = File.CreateText(batchFile))
            {
                foreach (KeyValuePair<string, ValuePair> pair in jobs)
                {
                    if (pair.Value.pageCount > 0)
                    {
                        StringBuilder command = new StringBuilder();
                        command.Append("LPR -S emtexvip1.nash.tenn -P ");
                        command.Append(pair.Value.printQueue);
                        command.Append(@" C:\PrintServices\");
                        command.Append(pair.Key);
                        writer.WriteLine(command.ToString());
                    }
                }
            }
            ConsoleHandler.Print("batched");
        }
    }
}
