using ExcelWIP.Models;
using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using System;
using ExcelWIP.VersionTwo.Models;

namespace ExcelWIP.Controllers
{
    public class TATController
    {
        //instance variables
        private ObservableCollection<TATModel> TATDataModel;
        private SQL.DataTable TATDataTable;

        //Get DataTable
        public SQL.DataTable GetTATDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            TATDataModel = new ObservableCollection<TATModel>();
            TATDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;

            //date time: today's date at 7:00 AM
            DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 7, 0, 0);

            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.DateComplete > today && i.DelayReason ==""
                         orderby i.DateComplete ascending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                TATDataModel.Add(new TATModel
                {
                    TATByDateIn = item.TATByDateIn,
                    TATByDockIn = item.TATByDockIn,
                    RefNumber = item.RefNumber,
                    DealerName = item.DealerName,
                    DateIn = item.DateIn,
                    DateDockIn = item.DateDockIn,
                    DateComplete = item.DateComplete,
                    DelayReason = item.DelayReason
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(TATDataModel))
            {
                TATDataTable.Load(reader);
            }

            //order the datatable's columns
            TATDataTable.Columns["TATByDateIn"].SetOrdinal(0);
            TATDataTable.Columns["TATByDockIn"].SetOrdinal(1);
            TATDataTable.Columns["DealerName"].SetOrdinal(2);
            TATDataTable.Columns["DateDockIn"].SetOrdinal(3);
            TATDataTable.Columns["DateComplete"].SetOrdinal(4);
            TATDataTable.Columns["DateIn"].SetOrdinal(5);
            TATDataTable.Columns["RefNumber"].SetOrdinal(6);
            TATDataTable.Columns["DelayReason"].SetOrdinal(7);

            return TATDataTable;
        }
    }
}
