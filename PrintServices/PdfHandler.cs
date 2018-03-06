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
        public static List<string> getFiles()  // Gets PDF files and puts them in a list for easy perusal.
        {
            List<string> files = new List<string>();
            foreach (string file in Directory.EnumerateFiles(@"C:\PrintServices", "*.pdf"))
            {
                files.Add(file);
            }
            return files;
        }
        public static void handleDuplicates(List<Job> joblist)  // Finds the duplicate files
        {
            ConsoleHandler.Print("duplicates");
            List<string> files = getFiles();
            List<string> duplicates = new List<string>();

            foreach (Job job in joblist)
            {
                List<string> result = files.FindAll(f => f.Contains(filenameTrimmer(job)));
                if (result.Count() > 1)
                {
                    duplicates = result.ToList();
                    mergePdfs(duplicates, job);
                }
            }
            ConsoleHandler.Print("duplicatessuccess");
        }
        public static void mergePdfs(List<string> duplicates, Job job)  // Merges any duplicates into a new file, deleting the originals
        {
            string mergedFile = @"C:\PrintServices\" + job.FileName;
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
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("     {0} files for the {1} Job were combined into one file.", duplicates.Count, filenameTrimmer(job));
        }
        public static void renameFiles(List<Job> jobs)  // Renames PDF to the simplified format
        {
            ConsoleHandler.Print("rename");
            List<string> files = getFiles();
            foreach (Job job in jobs)
            {
                string result = files.Find(f => f.Contains(filenameTrimmer(job)));
                if (result != null)
                {
                    File.Move(result, @"C:\PrintServices\" + job.FileName);
                } else { continue; }
            }
            ConsoleHandler.Print("renamed");
        }
        public static string filenameTrimmer(string file) // Helper method to remove portions of a filename
        {
            char[] pdfString = { '.', 'p', 'd', 'f' };
            string trimmedFilename = "";

            trimmedFilename = file.Substring(17).TrimEnd(pdfString);

            return trimmedFilename;
        }
        public static string filenameTrimmer(Job job) // Helper method to remove portions of a filename
        {
            char[] pdfString = { '.', 'p', 'd', 'f' };
            string trimmedFilename = "";

            trimmedFilename = job.FileName.TrimEnd(pdfString);

            return trimmedFilename;
        }
    }
}
