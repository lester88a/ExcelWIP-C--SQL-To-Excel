using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using ExcelWIP.VersionTwo.Models;

namespace ExcelWIP.VersionTwo.Controllers
{
    public class BulkController
    {
        //instance variables
        private ObservableCollection<BulkModel> BulkDataModel;
        private SQL.DataTable BulkDataTable;

        //Get DataTable
        public SQL.DataTable GetBulkDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            BulkDataModel = new ObservableCollection<BulkModel>();
            BulkDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;
            
            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.Program == "BULK"
                         orderby i.Aging ascending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                BulkDataModel.Add(new BulkModel
                {
                    RefNumber = item.RefNumber,
                    Aging = item.Aging,
                    FuturetelLocation = item.FuturetelLocation,
                    Program = item.Program
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(BulkDataModel))
            {
                BulkDataTable.Load(reader);
            }

            //order the datatable's columns
            BulkDataTable.Columns["RefNumber"].SetOrdinal(0);
            BulkDataTable.Columns["Aging"].SetOrdinal(1);
            BulkDataTable.Columns["FuturetelLocation"].SetOrdinal(2);
            BulkDataTable.Columns["Program"].SetOrdinal(3);

            return BulkDataTable;
        }
    }
}
