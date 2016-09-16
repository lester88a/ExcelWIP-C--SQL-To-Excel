using System;
using ExcelWIP.VersionTwo;
using System.Threading;
using ExcelWIP.VersionTwo.PivotTable;
using ExcelWIP.VersionTwo.EmailService;
using System.Collections;
using SQL = System.Data;

namespace ExcelWIP
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //begin monitoring for execution time
                var watch = System.Diagnostics.Stopwatch.StartNew();

                #region version one
                ////local variables
                //string fileName = null;
                //string sheetNameBulk = "Bulk";
                //string sheetNameRetail = "Retail";
                //string sheetNameRogers = "Rogers";
                //string sheetNameNonRogers = "NonRogers";

                ////Creat new excel sheet object
                //ExcelSheet excelSheet = new ExcelSheet(ref fileName);

                ////create pivot tables
                //PivotTable pt = new PivotTable();
                ////bulk piovt table
                //pt.CreatePiovtTable(fileName, sheetNameBulk);
                ////retail pivot table
                //pt.CreatePiovtTable(fileName, sheetNameRetail);
                ////rogers pivot table
                //pt.CreatePiovtTable(fileName, sheetNameRogers);
                #endregion


                #region version two
                string fileName = null;
                SQL.DataTable techDataTable = null;
                string sheetBulk = "_Bulk";
                string sheetRetail = "_Retail";
                string sheetRogers = "_RogersDealer";
                string sheetNonRogers = "_NonRogersDealer";
                string sheetProrityList = "_PriorityList";

                ExcelManager ex = new ExcelManager();
                ex.GetExcelSheet(ref fileName, ref techDataTable);

                //create pivot tables
                PivotTableForNormal pt = new PivotTableForNormal();
                pt.CreatePiovtTable(fileName, sheetNonRogers);
                pt.CreatePiovtTable(fileName, sheetRogers);
                pt.CreatePiovtTable(fileName, sheetRetail);
                pt.CreatePiovtTable(fileName, sheetBulk);
                
                PiovtTableForPriority ptPriority = new PiovtTableForPriority();
                ptPriority.CreatePiovtTable(fileName, sheetProrityList);

                #endregion

                //Email Service
                #region Email Service
                TechOutputString techString = new TechOutputString();
                string tString = techString.ConvertDataTableToString(techDataTable);

                EmailService es = new EmailService();
                es.SendEmailMethod(fileName, fileName.Remove(0, 52).Remove(fileName.Remove(0, 52).Length - 5), "*** This is an automatically generated email, please do not reply ***\n\n" + tString);
                Console.WriteLine("Email Send Success");
                #endregion

                //stop monitoring execution time
                #region Calculate Execution Time
                watch.Stop();
                var elapsedMs = watch.ElapsedMilliseconds;
                //display execution time
                Console.WriteLine("Execution time: " + elapsedMs / 1000 + " seconds");
                #endregion
            }
            catch (Exception e)
            {
                //display error info
                Console.WriteLine("Error: \n" + e);
            }
            Console.WriteLine("Application will exit within 5 seconds.");
            //delay 5 seconds for user to read detailed info
            Thread.Sleep(5000);


        }
    }
}
