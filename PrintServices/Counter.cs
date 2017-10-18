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
        public static int getNumberOfPages(string fileName)
        {
            using (StreamReader streamReader = new StreamReader(File.OpenRead(fileName)))
            {
                Regex regex = new Regex(@"/Type\s*/Page[^s]");
                MatchCollection matches = regex.Matches(streamReader.ReadToEnd());
                return matches.Count;
            }
        }
        //public static void createCountDictionary()
        //{
        //    Dictionary<string, string> countDictionary = new Dictionary<string, string>();
        //    string fileName = "";
        //    string pageCount = "";
        //    foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.pdf"))
        //    {
        //        fileName = PdfHandler.filenameTrimmer(file);

        //        if (FileInfo.TwoSidedFiles.Contains(file))
        //        {
        //            pageCount = (Counter.getNumberOfPages(file) / 2).ToString();
        //        }
        //        else
        //        {
        //            pageCount = (Counter.getNumberOfPages(file) / 2).ToString();
        //        }

        //        countDictionary.Add(fileName, pageCount);
        //        Counter.AddToTotalCount(file, Counter.getNumberOfPages(file));

        //    }
        //    return countDictionary;
        //}
        //public static Dictionary<string, string> createCountDictionary()
        //{
        //    Dictionary<string, string> countDictionary = new Dictionary<string, string>();
        //    string fileName = "";
        //    string pageCount = "";
        //    foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.pdf"))
        //    {
        //        fileName = PdfHandler.filenameTrimmer(file);

        //        if (FileInfo.TwoSidedFiles.Contains(file))
        //        {
        //            pageCount = (Counter.getNumberOfPages(file) / 2).ToString();
        //        }
        //        else
        //        {
        //            pageCount = (Counter.getNumberOfPages(file) / 2).ToString();
        //        }

        //        countDictionary.Add(fileName, pageCount);
        //        Counter.AddToTotalCount(file, Counter.getNumberOfPages(file));

        //    }
        //    return countDictionary;
        //}
    }
}
