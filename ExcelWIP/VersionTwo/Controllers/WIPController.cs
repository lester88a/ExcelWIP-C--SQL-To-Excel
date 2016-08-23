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
                    Aging = item.Aging,
                    DealerName = item.DealerName,
                    DealerID = item.DealerID,
                    GSPNTicketNo = item.GSPNTicketNo,
                    DateIn = item.DateIn,
                    DateComplete = item.DateComplete,
                    DateDockIn = item.DateDockIn,
                    DateDockOut = item.DateDockOut,
                    DateEstimation = item.DateEstimation,
                    DateApproved = item.DateApproved,
                    DateReject = item.DateReject,
                    DateBackorder = item.DateBackorder,
                    Warranty = item.Warranty,
                    Program = item.Program,
                    FuturetelLocation = item.FuturetelLocation,
                    ESN = item.ESN,
                    MCN = item.MCN,
                    SVP = item.SVP,
                    ModelNumber = item.ModelNumber,
                    DelayReason=item.DelayReason,
                    Technician = item.Technician,
                    DealerRefNumber=item.DealerRefNumber,
                    Status = item.Status,
                    Manufacturer = item.Manufacturer
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(WIPDataModel))
            {
                WIPDataTable.Load(reader);
            }

            //order the datatable's columns
            WIPDataTable.Columns["RefNumber"].SetOrdinal(0);
            WIPDataTable.Columns["Aging"].SetOrdinal(1);
            WIPDataTable.Columns["DealerName"].SetOrdinal(2);
            WIPDataTable.Columns["DealerID"].SetOrdinal(3);
            WIPDataTable.Columns["GSPNTicketNo"].SetOrdinal(4);
            WIPDataTable.Columns["DateIn"].SetOrdinal(5);
            WIPDataTable.Columns["DateComplete"].SetOrdinal(6);
            WIPDataTable.Columns["DateDockIn"].SetOrdinal(7);
            WIPDataTable.Columns["DateDockOut"].SetOrdinal(8);
            WIPDataTable.Columns["DateEstimation"].SetOrdinal(9);
            WIPDataTable.Columns["DateApproved"].SetOrdinal(10);
            WIPDataTable.Columns["DateReject"].SetOrdinal(11);
            WIPDataTable.Columns["DateBackorder"].SetOrdinal(12);
            WIPDataTable.Columns["Warranty"].SetOrdinal(13);
            WIPDataTable.Columns["Program"].SetOrdinal(14);
            WIPDataTable.Columns["FuturetelLocation"].SetOrdinal(15);
            WIPDataTable.Columns["ESN"].SetOrdinal(16);
            WIPDataTable.Columns["MCN"].SetOrdinal(17);
            WIPDataTable.Columns["SVP"].SetOrdinal(18);
            WIPDataTable.Columns["ModelNumber"].SetOrdinal(19);
            WIPDataTable.Columns["DelayReason"].SetOrdinal(20);
            WIPDataTable.Columns["Technician"].SetOrdinal(21);
            WIPDataTable.Columns["DealerRefNumber"].SetOrdinal(22);

            return WIPDataTable;
        }
    }
}
