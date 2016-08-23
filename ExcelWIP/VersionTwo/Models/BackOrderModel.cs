using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWIP.VersionTwo.Models
{
    public class BackOrderModel
    {
        public int RefNumber { get; set; }
        public int Aging { get; set; }
        public string ModelNumber { get; set; }
        public DateTime? DateIn { get; set; }
        public DateTime? DateDockIn { get; set; }
        public string FuturetelLocation { get; set; }
        public string DelayReason { get; set; }
        public string Technician { get; set; }
        public string Status { get; set; }
    }
}
