﻿using ExcelWIP.Models;
using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using ExcelWIP.VersionTwo.Models;


namespace ExcelWIP.VersionTwo.Controllers
{
    public class BackOrderController
    {
        //instance variables
        private ObservableCollection<BackOrderModel> BackOrderDataModel;
        private SQL.DataTable BackOrderDataTable;

        //Get DataTable
        public SQL.DataTable GetBackOrderDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            BackOrderDataModel = new ObservableCollection<BackOrderModel>();
            BackOrderDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;

            //LINQ query to get DataIn data from all data
            var dataIn = from i in allData.AsEnumerable()
                         where i.FuturetelLocation == "B"
                         orderby i.Aging descending, i.RefNumber ascending
                         select i;
            //assign data to connection
            foreach (var item in dataIn)
            {
                BackOrderDataModel.Add(new BackOrderModel
                {
                    RefNumber = item.RefNumber,
                    Aging = item.Aging,
                    ModelNumber = item.ModelNumber,
                    DateIn = item.DateIn,
                    DateDockIn = item.DateDockIn,
                    DelayReason = item.DelayReason,
                    Technician = item.Technician,
                    Status = item.Status
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(BackOrderDataModel))
            {
                BackOrderDataTable.Load(reader);
            }

            //order the datatable's columns
            BackOrderDataTable.Columns["RefNumber"].SetOrdinal(0);
            BackOrderDataTable.Columns["Aging"].SetOrdinal(1);
            BackOrderDataTable.Columns["ModelNumber"].SetOrdinal(2);
            BackOrderDataTable.Columns["DateIn"].SetOrdinal(3);
            BackOrderDataTable.Columns["DateDockIn"].SetOrdinal(4);
            BackOrderDataTable.Columns["DelayReason"].SetOrdinal(5);
            BackOrderDataTable.Columns["Technician"].SetOrdinal(6);
            BackOrderDataTable.Columns["Status"].SetOrdinal(5);

            return BackOrderDataTable;
        }

    }
}
