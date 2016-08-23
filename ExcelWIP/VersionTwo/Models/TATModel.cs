using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWIP.Models
{
    public class TATModel
    {
        public string TATByDateIn { get; set; }
        public string TATByDockIn { get; set; }
        public int RefNumber { get; set; }
        public string DealerName { get; set; }
        public DateTime? DateIn { get; set; }
        public DateTime? DateDockIn { get; set; }
        public DateTime? DateComplete { get; set; }
    }
}
