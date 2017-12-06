using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PrintServices.Models;

namespace PrintServices.DAL
{
    public class JobContext : DbContext
    {
        public virtual DbSet<Job> Jobs { get; set; }
    }
}
