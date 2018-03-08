using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using static PrintServices.ExcelHandler;

namespace PrintServices
{
    class Counter
    {
        public static int readNumberOfPages(string filename)  // reads the pagecount a PDF
        {
            PdfReader pdfReader = new PdfReader(filename);
            return pdfReader.NumberOfPages;
        }
        public static List<KeyValuePair<string, ValuePair>> assignSheetCount(List<KeyValuePair<string, ValuePair>> jobsAndCounts) // adds the pagecount of each file to the list of jobs
        {
            ConsoleHandler.Print("count");
            List<string> files = PdfHandler.getFiles();

            foreach (string file in files)
            {
                string filename = PdfHandler.filenameTrimmer(file) + ".pdf";
                KeyValuePair<string, ValuePair> pair = jobsAndCounts.Find(j => j.Key == filename);

                int index = jobsAndCounts.IndexOf(pair);
                if (index != -1)
                    jobsAndCounts[index] = new KeyValuePair<string, ValuePair>(pair.Key, new ValuePair { pageCount = readNumberOfPages(file), printQueue = pair.Value.printQueue });
            }
            return jobsAndCounts;
        }
        public static List<KeyValuePair<string, ValuePair>> UpdatePageCount(List<KeyValuePair<string, ValuePair>> jobsAndCounts)  // Divides the pagecount by two if it is supposed to be a double-sided job
        {
            List<KeyValuePair<string, ValuePair>> UpdatedList = new List<KeyValuePair<string, ValuePair>>();
            foreach (KeyValuePair<string, ValuePair> pair in jobsAndCounts)
            {
                int updatedCount;
                if (pair.Value.printQueue == "ce5793_RB-2_pdf" ||
                    pair.Value.printQueue == "ce5786_RB-2_pdf" ||
                    pair.Value.printQueue == "ce5793_FIRE-2_pdf")
                {
                    updatedCount = pair.Value.pageCount / 2;
                }
                else
                {
                    updatedCount = pair.Value.pageCount;
                }
               
                UpdatedList.Add(new KeyValuePair<string, ValuePair>(pair.Key, new ValuePair { pageCount = updatedCount, printQueue = pair.Value.printQueue }));               
            }
            return UpdatedList;
        }
    }
}
