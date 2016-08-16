using System;
using System.Data.SqlClient;
using System.Linq;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelWIP
{
    class Program
    {
        private static string FileName;
        private static string ConnectionString;
        private static string QueryWIP;
        /*------------------------------------------------------------------------------------------------*/
        static void Main(string[] args)
        {
            Connection co = new Connection();
            FileName = co.GetFileName();
            ConnectionString = co.GetConnectionString();
            QueryWIP = co.GetQueryWIP();
            
            try
            {

                WIP wip = new WIP();
                wip.SQLToExcel(FileName, ConnectionString, QueryWIP);

                Task.Delay(1000);

                PivotTableBulk pt = new PivotTableBulk();
                pt.GetPiovtTableBulk(FileName);

            }
            catch (Exception e)
            {

                Console.WriteLine("Error: \n"+e);
            }

            

            Console.ReadKey();
        }

       
    }
}
