using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using ExcelWIP.VersionTwo.Models;

namespace ExcelWIP.VersionTwo.Controllers
{
    public class PriorityController
    {
        //instance variables
        private ObservableCollection<PriorityModel> PriorityDataModel;
        private SQL.DataTable PriorityDataTable;

        //Get DataTable
        public SQL.DataTable GetPriorityDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            PriorityDataModel = new ObservableCollection<PriorityModel>();
            PriorityDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;

            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.FuturetelLocation !="E" && i.FuturetelLocation !="SA" && i.Status!="C"
                         orderby i.Aging ascending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                PriorityDataModel.Add(new PriorityModel
                {
                    RefNumber = item.RefNumber,
                    Aging = item.Aging,
                    FuturetelLocation = item.FuturetelLocation,
                    Program = item.Program,
                    Warranty = item.Warranty
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(PriorityDataModel))
            {
                PriorityDataTable.Load(reader);
            }

            //order the datatable's columns
            PriorityDataTable.Columns["RefNumber"].SetOrdinal(0);
            PriorityDataTable.Columns["Aging"].SetOrdinal(1);
            PriorityDataTable.Columns["FuturetelLocation"].SetOrdinal(2);
            PriorityDataTable.Columns["Program"].SetOrdinal(3);
            PriorityDataTable.Columns["Warranty"].SetOrdinal(4);

            return PriorityDataTable;
        }
    }
}
