using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWIP
{
    public class ExcelSheet:DataTable
    {
        //constructor
        public ExcelSheet()
        {
            //begin monitoring for execution time
            var watch = System.Diagnostics.Stopwatch.StartNew();
            //get connection string
            Connection con = new Connection();
            //get query
            string queryWIP = con.GetQueryWIP();
            string queryIn = con.GetQueryIn();
            string queryOut = con.GetQueryOut();
            string queryTAT = con.GetQueryTAT();
            string queryBackOrder = con.GetQueryBackOrder();
            string queryBulk = con.GetQueryBulk();
            string queryRetail = con.GetQueryRetail();
            //set worksheet name
            string sheetWIP = "WIP";
            string sheetIn = "In";
            string sheetOut = "Out";
            string sheetTAT = "TAT";
            string sheetBackOrder = "BackOrder";
            string sheetBulk = "Bulk";
            string sheetRetail = "Retail";

            //create excel sheet
            CreateExcel();

            //create sheet sheet by name
            CreatSheet(queryRetail, sheetRetail);
            CreatSheet(queryBulk, sheetBulk);
            CreatSheet(queryBackOrder, sheetBackOrder);
            CreatSheet(queryTAT, sheetTAT);
            CreatSheet(queryOut, sheetOut);
            CreatSheet(queryIn, sheetIn);
            CreatSheet(queryWIP, sheetWIP);

            //save excel sheet
            SaveExcel();

            //stop monitoring execution time
            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            //display execution time
            Console.WriteLine("Execution time: "+ elapsedMs/1000 + " seconds");
        }

        public override void CreateExcel()
        {
            base.CreateExcel();
        }

        public override void CreatSheet(string query, string sheetName)
        {
            base.CreatSheet(query, sheetName);
        }

        public override void SaveExcel()
        {
            base.SaveExcel();
        }


    }
}
