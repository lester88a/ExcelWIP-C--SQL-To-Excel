using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWIP.VersionTwo.Models
{
    public class PriorityModel
    {
        public int RefNumber { get; set; }
        public int Aging { get; set; }
        public string FuturetelLocation { get; set; }
        public string Program { get; set; }
        public bool Warranty { get; set; }
    }
}
