using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace ToneDownThatBackEnd.Models
{
    public class Job
    {
        [Key]
        public int EntryId { get; set; }
        public string FileName { get; set; }
        public int PageCount { get; set; }
        public string PrintQueue { get; set; }
        public string Board { get; set; }
        public string Disposition { get; set; }
    }
}