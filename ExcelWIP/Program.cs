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
        static void Main(string[] args)
        {
            try
            {
                //Creat new excel sheet object
                ExcelSheet excelSheet = new ExcelSheet();
            }
            catch (Exception e)
            {
                //display error info
                Console.WriteLine("Error: \n"+e);
            }
            //press any key to continue
            Console.ReadKey();
        }
    }
}
