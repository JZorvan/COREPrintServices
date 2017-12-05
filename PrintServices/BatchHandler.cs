using Microsoft.Office.Interop.Excel;
using PrintServices.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PrintServices
{
    class BatchHandler
    {
        public static void generateBatchFile(List<Job> jobs)
        {
            string batchFile = AppDomain.CurrentDomain.BaseDirectory + "test.bat";
            
            if (File.Exists(batchFile))
            {
                File.Delete(batchFile);
            }

            using (StreamWriter writer = File.CreateText(batchFile))
            {
                foreach (Job job in jobs)
                {
                    if (job.PageCount > 0)
                    {
                        StringBuilder command = new StringBuilder();
                        command.Append("LPR -S emtexvip1.nash.tenn -P ");
                        command.Append(job.PrintQueue);
                        command.Append(@" C:\PrintServices\");
                        command.Append(job.FileName);
                        writer.WriteLine(command.ToString());
                    }
                }
            }
        }
    }
}
