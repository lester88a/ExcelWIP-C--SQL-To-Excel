using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWIP.Models
{
    public class DateInModel
    {
        public int RefNumber { get; set; }
        public string ModelNumber { get; set; }
        public string DealerName { get; set; }
        public string FTRMA { get; set; }
        public DateTime? DateIn { get; set; } 
        public DateTime? DateDockIn { get; set; }
        public int? DealerID { get; set; }
        public bool Warranty { get; set; }
    }
}
