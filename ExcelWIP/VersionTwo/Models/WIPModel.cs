using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWIP.VersionTwo.Models
{
    public class WIPModel
    {
        public int RefNumber { get; set; }
        public string DealerName { get; set; }
        public int? DealerID { get; set; }
        public string GSPNTicketNo { get; set; }
        public string FuturetelLocation { get; set; }
        public bool Warranty { get; set; }
        public DateTime? DateDockIn { get; set; }
        public DateTime? DateIn { get; set; }
        public string SVP { get; set; }
        public string ToFactoryWayBill { get; set; }
        public string Program { get; set; }
        public string ModelNumber { get; set; }
        public string FTRMA { get; set; }
        public string ShipWayBill { get; set; }
        public string DelayReason { get; set; }
        public string Type { get; set; }
        public string Technician { get; set; }
        public int Aging { get; set; }
    }
}
