using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using System;
using ExcelWIP.VersionTwo.Models;

namespace ExcelWIP.VersionTwo.Controllers
{
    public class RetailController
    {
        //instance variables
        private ObservableCollection<RetailModel> RetailDataModel;
        private SQL.DataTable RetailDataTable;

        //Get DataTable
        public SQL.DataTable GetRetailDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            RetailDataModel = new ObservableCollection<RetailModel>();
            RetailDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;

           
            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.DealerID == 517
                         orderby i.Aging ascending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                RetailDataModel.Add(new RetailModel
                {
                    RefNumber = item.RefNumber,
                    Aging = item.Aging,
                    FuturetelLocation = item.FuturetelLocation,
                    Program = item.Program
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(RetailDataModel))
            {
                RetailDataTable.Load(reader);
            }

            //order the datatable's columns
            RetailDataTable.Columns["RefNumber"].SetOrdinal(0);
            RetailDataTable.Columns["Aging"].SetOrdinal(1);
            RetailDataTable.Columns["FuturetelLocation"].SetOrdinal(2);
            RetailDataTable.Columns["Program"].SetOrdinal(3);

            return RetailDataTable;
        }
    }
}
