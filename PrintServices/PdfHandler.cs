using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using PrintServices.Models;

namespace PrintServices
{
    class PdfHandler
    {

        public static void findDuplicates(List<Job> joblist)
        {
            string[] pdfFiles = Directory.GetFiles("C:/PrintServices", "*.pdf");
            List<string> duplicates = new List<string>();
            foreach (Job job in joblist)
            {
                var result = Array.FindAll(pdfFiles, s => s.Contains(filenameTrimmer(job.FileName)));
                if (result.Count() > 1)
                {
                    duplicates = result.ToList();
                    mergePdfs(duplicates, filenameTrimmer(job.FileName), result.Count());
                } 
            }
        }
        public static void mergePdfs(List<string> duplicateFiles, string jobName, int numDocs)
        {
            string mergedFile = "C:/PrintServices\\" + jobName + ".pdf";
            using (FileStream stream = new FileStream(mergedFile, FileMode.Create))
            using (Document doc = new Document())
            using (PdfCopy pdf = new PdfCopy(doc, stream))
            {
                doc.Open();

                PdfReader reader = null;
                PdfImportedPage page = null;

                duplicateFiles.ForEach(file =>
                {
                    reader = new PdfReader(file);

                    for (int i = 0; i < reader.NumberOfPages; i++)
                    {
                        page = pdf.GetImportedPage(reader, i + 1);
                        pdf.AddPage(page);
                    }

                    pdf.FreeReader(reader);
                    reader.Close();
                    File.Delete(file);
                });
                Console.WriteLine("There were " + numDocs + " files for the " + jobName + " job that were merged into one file.");
            }
        }
        public static void renameFiles()
        {
            foreach (string file in Directory.EnumerateFiles("C:/PrintServices", "*.pdf"))
            {
                File.Move(file, getNewFilename(file));
            }
        }
        public static string getNewFilename(string file)
        {
            if (file.Contains("lp32-15C"))
            {
                return file.Substring(0, 25) + ".pdf";
            }
            else if (file.Contains("lp32-19C"))
            {
                return file.Substring(0, 25) + ".pdf";
            }
            else if (file.Contains("lp59-15C"))
            {
                return file.Substring(0, 25) + ".pdf";
            }
            else if (file.Contains("lp59-1601B"))
            {
                return file.Substring(0, 27) + ".pdf";
            }
            else if (file.Contains("lp59-3502B"))
            {
                return file.Substring(0, 27) + ".pdf";
            }
            else if (file.Contains("25CAN001"))
            {
                return "C:/PrintServices\\25CAN001.pdf";
            }
            else if (file.Contains("25CAA001"))
            {
                return "C:/PrintServices\\25CAA001.pdf";
            }
            else if (file.Contains("tndc06"))
            {
                return file.Substring(0, 28) + ".pdf";
            }
            else
            {
                return file.Substring(0, 26) + ".pdf";
            }
        }
        public static string filenameTrimmer(string file)
        {
            char[] pdfString = { '.', 'p', 'd', 'f' };
            string trimmedFilename = "";

            trimmedFilename = file.Substring(17).TrimEnd(pdfString);

            return trimmedFilename;
        }
        public static string filenameTrimmer(Job job)
        {
            char[] pdfString = { '.', 'p', 'd', 'f' };
            string trimmedFilename = "";

            trimmedFilename = job.FileName.TrimEnd(pdfString);

            return trimmedFilename;
        }
    }
}
