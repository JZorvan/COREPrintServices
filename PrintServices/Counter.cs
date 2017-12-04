using iTextSharp.text.pdf;
using PrintServices.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PrintServices
{
    class Counter
    {
        public static int readNumberOfPages(string filename)
        {
            PdfReader pdfReader = new PdfReader(filename);
            return pdfReader.NumberOfPages;
        }
        public static List<Job> assignSheetCount(List<Job> jobs)
        {
            List<string> files = PdfHandler.getFiles();
            foreach (string file in files)
            {
                string filename = PdfHandler.filenameTrimmer(file) + ".pdf";
                Job result = jobs.Find(j => j.FileName == filename);
                if (result != null)
                {
                    result.PageCount = readNumberOfPages(file);
                }
            }
            return jobs;
        }
    }
}
