using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWIP.VersionTwo.Models
{
    public class Repair
    {
        public int RefNumber { get; set; }
        public int Aging { get; set; }
        public string TATByDateIn { get; set; }
        public string TATByDockIn { get; set; }
        public string ESN { get; set; }
        public string MCN { get; set; }
        public string GSPNTicketNo { get; set; }
        public DateTime? DateIn { get; set; }
        public DateTime? DateComplete { get; set; }
        public DateTime? DateDockIn { get; set; }
        public DateTime? DateDockOut { get; set; }
        public DateTime? DateEstimation { get; set; }
        public DateTime? DateApproved { get; set; }
        public DateTime? DateReject { get; set; }
        public DateTime? DateBackorder { get; set; }
        public DateTime? DateFinish { get; set; }
        public bool Warranty { get; set; }
        public string Program { get; set; }
        public int DealerID { get; set; }
        public string DealerName { get; set; }
        public string FuturetelLocation { get; set; }
        public string SVP { get; set; }
        public string ModelNumber { get; set; }
        public string DelayReason { get; set; }
        public string Technician { get; set; }
        public string DealerRefNumber { get; set; }
        public string Status { get; set; }
        public string Manufacturer { get; set; }
        public string FTRMA { get; set; }
        public string ShipWayBill { get; set; }
        public string Type { get; set; }
        public string ToFactoryWayBill { get; set; }
    }
}
