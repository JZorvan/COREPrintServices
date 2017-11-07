using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToneDownThatBackEnd.Models;

namespace PrintServices.DAL
{
    public class JobRepo
    {
        public JobContext Context { get; set; }
        public JobRepo()
        {
            Context = new JobContext();
        }
        public JobRepo(JobContext _context)
        {
            Context = _context;
        }
        public void ImportMasterSpreadsheet()
        {
            // Create db from the master spreadsheet with zero page counts
        }
        public List<Job> GetJobs()
        {
            int i = 1;
            return Context.Jobs.ToList();
        }
    }
}
