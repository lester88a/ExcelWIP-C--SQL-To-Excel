using ExcelWIP.Models;
using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using ExcelWIP.VersionTwo.Models;

namespace ExcelWIP.Controllers
{
    public class DateInController
    {
        //instance variables
        private ObservableCollection<DateInModel> DateInDataModel;
        private SQL.DataTable DateInDataTable;

        //Get DataIn DataTable
        public SQL.DataTable GetDateInDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            DateInDataModel = new ObservableCollection<DateInModel>();
            DateInDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;

            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.Aging == 0 && i.Status != "C"
                         orderby i.Aging descending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                DateInDataModel.Add(new DateInModel
                {
                    RefNumber = item.RefNumber,
                    ModelNumber = item.ModelNumber,
                    DealerName = item.DealerName,
                    DateIn = item.DateIn,
                    DateDockIn = item.DateDockIn,
                    DealerID = item.DealerID,
                    Warranty = item.Warranty,
                    FTRMA = item.FTRMA
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(DateInDataModel))
            {
                DateInDataTable.Load(reader);
            }

            //order the datatable's columns
            DateInDataTable.Columns["RefNumber"].SetOrdinal(0);
            DateInDataTable.Columns["ModelNumber"].SetOrdinal(1);
            DateInDataTable.Columns["DealerName"].SetOrdinal(2);
            DateInDataTable.Columns["FTRMA"].SetOrdinal(3);
            DateInDataTable.Columns["DateIn"].SetOrdinal(4);
            DateInDataTable.Columns["DateDockIn"].SetOrdinal(5);
            DateInDataTable.Columns["DealerID"].SetOrdinal(6);
            DateInDataTable.Columns["Warranty"].SetOrdinal(7);

            return DateInDataTable;
        }
    }
}
