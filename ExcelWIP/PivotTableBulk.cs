using System;
using System.Data.SqlClient;
using System.Linq;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace ExcelWIP
{
    public class PivotTableBulk
    {
        //convert Bulk sheet to piovt table
        public void GetPiovtTableBulk(string fileName)
        {
            //Create Excel objects
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workBook;

            Excel._Worksheet workSheet;

            excel = new Excel.Application();

            workBook = excel.Workbooks.Open(fileName);
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

            /*-------------------------*/
            //create work sheet name
            workSheet = (Excel._Worksheet)excel.Worksheets.Add();
            workSheet.Name = "Bulk PiovtTable";
            

            // specify first cell for pivot table
            Excel.Range oRange2 = workSheet.Cells[1, 1];
            

            // create Pivot Cache and Pivot Table
            Excel.PivotCache pivotCache = workBook.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, "Bulk!A1:D900");

            Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(TableDestination: oRange2, TableName: "BulkSummary");
            
            #region
            //Set up the Program field as the page field, and Aging as the row field
            Excel.PivotField pageField = (Excel.PivotField)pivotTable.PivotFields("Program");
            pageField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;

            Excel.PivotField rowField = (Excel.PivotField)pivotTable.PivotFields("Aging");
            rowField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

            Excel.PivotField columnField = (Excel.PivotField)pivotTable.PivotFields("FuturetelLocation");
            columnField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

            pivotTable.AddDataField(pivotTable.PivotFields("RefNumber"), "Count of RefNumber", Excel.XlConsolidationFunction.xlCount);
            #endregion
            
            workBook.Save();
            workBook.Close();
            excel.Quit();
            Console.WriteLine("Success saved!");

        }
    }
}
