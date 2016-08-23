using System.Linq;
using SQL = System.Data;
using System.Collections.ObjectModel;
using FastMember;
using ExcelWIP.VersionTwo.Models;
using System;

namespace ExcelWIP.VersionTwo.Controllers
{
    public class TechnicianController
    {
        //instance variables
        private ObservableCollection<TechModel> TechDataModel;
        private SQL.DataTable TechDataTable;

        //Get DataTable
        public SQL.DataTable GetTechDataTable(ObservableCollection<Repair> AllRepairData)
        {
            //initialize objects
            TechDataModel = new ObservableCollection<TechModel>();
            TechDataTable = new SQL.DataTable();

            //get all data
            var allData = AllRepairData;

            //date time: today's date at 7:00 AM
            DateTime today = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 7, 0, 0);

            //LINQ query to get DataIn data from all data
            var techData = from i in allData.AsEnumerable()
                         where i.DateFinish >= today
                         orderby i.Aging ascending, i.RefNumber ascending
                         select i;

            //useing linq query to get total outputs of each tech
            var groups = (techData.GroupBy(n => n.Technician)
                .Select(n => new { MetricName = n.Key, MetricCount = n.Count() })
                .OrderByDescending(n => n.MetricCount));

            //assign data to connection
            foreach (var item in groups)
            {
                TechDataModel.Add(new TechModel
                {
                    Technician = item.MetricName,
                    TotalOutput = item.MetricCount
                });
            }

            //convert connection to DataTable using FastMember reference
            using (var reader = ObjectReader.Create(TechDataModel))
            {
                TechDataTable.Load(reader);
            }

            //order the datatable's columns
            TechDataTable.Columns["Technician"].SetOrdinal(0);
            TechDataTable.Columns["TotalOutput"].SetOrdinal(1);

            return TechDataTable;
        }
    }
}
