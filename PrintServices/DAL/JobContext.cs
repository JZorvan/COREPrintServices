using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ToneDownThatBackEnd.Models;

namespace PrintServices.DAL
{
    public class JobContext : DbContext
    {
        public virtual DbSet<Job> Jobs { get; set; }
    }
}
