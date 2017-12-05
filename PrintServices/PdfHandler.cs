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
        public static List<string> getFiles()
        {
            List<string> files = new List<string>();
            foreach (string file in Directory.EnumerateFiles(AppDomain.CurrentDomain.BaseDirectory, "*.pdf"))
            {
                files.Add(file);
            }
            return files;
        }
        public static void handleDuplicates(List<Job> joblist)
        {
            List<string> files = getFiles();
            List<string> duplicates = new List<string>();
            foreach (Job job in joblist)
            {
                var result = files.FindAll(f => f.Contains(filenameTrimmer(job)));
                if (result.Count() > 1)
                {
                    duplicates = result.ToList();
                    mergePdfs(duplicates, filenameTrimmer(job));
                }
            }
        }
        public static void mergePdfs(List<string> duplicates, string jobName)
        {
            string mergedFile = AppDomain.CurrentDomain.BaseDirectory + jobName + ".pdf";
            using (FileStream stream = new FileStream(mergedFile, FileMode.Create))
            using (Document doc = new Document())
            using (PdfCopy pdf = new PdfCopy(doc, stream))
            {
                doc.Open();

                PdfReader reader = null;
                PdfImportedPage page = null;

                duplicates.ForEach(file =>
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
            }
        }
        public static void handleRenaming(List<Job> jobs)
        {
            List<string> files = getFiles();
            foreach (Job job in jobs)
            {
                string result = files.Find(f => f.Contains(filenameTrimmer(job)));
                if (result != null)
                {
                    renameFile(result, AppDomain.CurrentDomain.BaseDirectory + job.FileName);
                } else { continue; }
            }
        }
        public static void renameFile(string file, string newFilename)
        {
            File.Move(file, newFilename);
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
