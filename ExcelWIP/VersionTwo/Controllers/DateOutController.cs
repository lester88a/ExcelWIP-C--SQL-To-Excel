using ExcelWIP.Models;
using System.Data.SqlClient;
using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using Newtonsoft.Json;
using FastMember;
using System;
using ExcelWIP.VersionTwo.Models;

namespace ExcelWIP.Controllers
{
    public class DateOutController
    {
        //instance variables
        private ObservableCollection<DateOutModel> DateOutDataModel;
        private SQL.DataTable DateInDataTable;

        //Get DataTable
        public SQL.DataTable GetDateOutDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            DateOutDataModel = new ObservableCollection<DateOutModel>();
            DateInDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;

            //date time: today's date at 7:00 AM
            DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 7, 0, 0);

            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.DateComplete > today && i.Manufacturer=="SAMSUNG"
                         orderby i.DateComplete ascending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                DateOutDataModel.Add(new DateOutModel
                {
                    RefNumber = item.RefNumber,
                    ModelNumber = item.ModelNumber,
                    DealerName = item.DealerName,
                    DealerID = item.DealerID,
                    DateIn = item.DateIn,
                    DateDockIn = item.DateDockIn,
                    DateComplete = item.DateComplete,
                    FTRMA = item.FTRMA
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(DateOutDataModel))
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
            DateInDataTable.Columns["DateComplete"].SetOrdinal(7);

            return DateInDataTable;
        }
    }
}
