using ExcelWIP.Models;
using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using ExcelWIP.VersionTwo.Models;


namespace ExcelWIP.VersionTwo.Controllers
{
    public class WIPController
    {
        //instance variables
        private ObservableCollection<WIPModel> WIPDataModel;
        private SQL.DataTable WIPDataTable;

        //Get DataTable
        public SQL.DataTable GetWIPDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            WIPDataModel = new ObservableCollection<WIPModel>();
            WIPDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;
            
            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.Status !="C"
                         orderby i.Aging ascending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                WIPDataModel.Add(new WIPModel
                {
                    RefNumber = item.RefNumber,
                    DealerName = item.DealerName,
                    DealerID = item.DealerID,
                    GSPNTicketNo = item.GSPNTicketNo,
                    FuturetelLocation = item.FuturetelLocation,
                    Warranty = item.Warranty,
                    DateDockIn = item.DateDockIn,
                    DateIn = item.DateIn,
                    SVP = item.SVP,
                    ToFactoryWayBill = item.ToFactoryWayBill,
                    Program = item.Program,
                    ModelNumber = item.ModelNumber,
                    FTRMA = item.FTRMA,
                    ShipWayBill = item.ShipWayBill,
                    DelayReason = item.DelayReason,
                    Type = item.Type,
                    Technician = item.Technician,
                    Aging = item.Aging,
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(WIPDataModel))
            {
                WIPDataTable.Load(reader);
            }

            //order the datatable's columns
            WIPDataTable.Columns["RefNumber"].SetOrdinal(0);
            WIPDataTable.Columns["DealerName"].SetOrdinal(1);
            WIPDataTable.Columns["DealerID"].SetOrdinal(2);
            WIPDataTable.Columns["GSPNTicketNo"].SetOrdinal(3);
            WIPDataTable.Columns["FuturetelLocation"].SetOrdinal(4);
            WIPDataTable.Columns["Warranty"].SetOrdinal(5);
            WIPDataTable.Columns["DateDockIn"].SetOrdinal(6);
            WIPDataTable.Columns["DateIn"].SetOrdinal(7);
            WIPDataTable.Columns["SVP"].SetOrdinal(8);
            WIPDataTable.Columns["ToFactoryWayBill"].SetOrdinal(9);
            WIPDataTable.Columns["Program"].SetOrdinal(10);
            WIPDataTable.Columns["ModelNumber"].SetOrdinal(11);
            WIPDataTable.Columns["FTRMA"].SetOrdinal(12);
            WIPDataTable.Columns["ShipWayBill"].SetOrdinal(13);
            WIPDataTable.Columns["DelayReason"].SetOrdinal(14);
            WIPDataTable.Columns["Type"].SetOrdinal(15);
            WIPDataTable.Columns["Technician"].SetOrdinal(16);
            WIPDataTable.Columns["Aging"].SetOrdinal(17);

            return WIPDataTable;
        }
    }
}
