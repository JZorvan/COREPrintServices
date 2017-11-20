using iTextSharp.text.pdf;
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
        public static int totalPageCount = 0;
        public static int totalDocCount = 0;
        public static int AddToTotalCount(string file, int pageCount)
        {
            totalDocCount++;
            if (FileInfo.TwoSidedFiles.Contains(file))
            {
                totalPageCount += (pageCount / 2);
            }
            else
            {
                totalPageCount += pageCount;
            }
            return totalPageCount;
        }
        public static int readNumberOfPages(string filename)
        {
            PdfReader pdfReader = new PdfReader(filename);
            Console.WriteLine(pdfReader.NumberOfPages);
            return pdfReader.NumberOfPages;
        }
        //public static int getNumberOfPages(string fileName)
        //{
        //    using (StreamReader streamReader = new StreamReader(File.OpenRead(fileName)))
        //    {
        //        Regex regex = new Regex(@"/Type\s*/Page[^s]");
        //        MatchCollection matches = regex.Matches(streamReader.ReadToEnd());
        //        return matches.Count;
        //    }
        //}
    }
}
