using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using ExcelWIP.VersionTwo.Models;

namespace ExcelWIP.VersionTwo.Controllers
{
    public class NonRogersController
    {
        //instance variables
        private ObservableCollection<RogersModel> RogersDataModel;
        private SQL.DataTable RogersDataTable;

        //Get DataTable
        public SQL.DataTable GetNonRogersDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            RogersDataModel = new ObservableCollection<RogersModel>();
            RogersDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;

            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.DealerID != 517 && (i.SVP != "FIDO" && i.SVP != "ROGERS")
                         orderby i.Aging ascending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                RogersDataModel.Add(new RogersModel
                {
                    RefNumber = item.RefNumber,
                    Aging = item.Aging,
                    FuturetelLocation = item.FuturetelLocation,
                    Program = item.Program
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(RogersDataModel))
            {
                RogersDataTable.Load(reader);
            }

            //order the datatable's columns
            RogersDataTable.Columns["RefNumber"].SetOrdinal(0);
            RogersDataTable.Columns["Aging"].SetOrdinal(1);
            RogersDataTable.Columns["FuturetelLocation"].SetOrdinal(2);
            RogersDataTable.Columns["Program"].SetOrdinal(3);

            return RogersDataTable;
        }
    }
}
