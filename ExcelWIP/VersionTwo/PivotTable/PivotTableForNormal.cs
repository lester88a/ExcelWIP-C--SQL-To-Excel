using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWIP.VersionTwo.PivotTable
{
    public class PivotTableForNormal
    {
        //convert Bulk sheet to piovt table
        public void CreatePiovtTable(string fileName, string sheetName)
        {
            //Create Excel objects
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workBook;

            Excel._Worksheet workSheet;

            excel.Visible = false;
            excel.DisplayAlerts = false;

            workBook = excel.Workbooks.Open(fileName);
            workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);

            /*-------------------------*/
            //create work sheet name
            workSheet = (Excel._Worksheet)excel.Worksheets.Add();
            workSheet.Name = sheetName.Remove(0,1);


            // specify first cell for pivot table
            Excel.Range oRange2 = workSheet.Cells[1, 1];


            // create Pivot Cache and Pivot Table
            Excel.PivotCache pivotCache = workBook.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, sheetName + "!A1:D5000");

            Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(TableDestination: oRange2, TableName: sheetName + "Summary");

            /*-------------------------*/
            //set up fields for pivot table
            #region setup fields
            //Set up the Program field as the page field, and Aging as the row field
            Excel.PivotField pageField = (Excel.PivotField)pivotTable.PivotFields("Program");
            pageField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            //set up Aging field as row field
            Excel.PivotField rowField = (Excel.PivotField)pivotTable.PivotFields("Aging");
            rowField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            //set up FuturetelLocation field as column field
            Excel.PivotField columnField = (Excel.PivotField)pivotTable.PivotFields("FuturetelLocation");
            columnField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;

            pivotTable.AddDataField(pivotTable.PivotFields("RefNumber"), "Count of RefNumber", Excel.XlConsolidationFunction.xlCount);
            #endregion

            /*-------------------------*/
            //save and exit excel work book
            excel.Visible = false;
            excel.DisplayAlerts = false;

            //hide the row data sheet
            Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
            worksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Console.WriteLine("workSheet [" + sheetName + "] hided.");
            //save
            workBook.Save();

            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(workSheet);
            workBook.Close();
            Marshal.ReleaseComObject(workBook);
            excel.Quit();
            Marshal.ReleaseComObject(excel);
            Console.WriteLine("Pivot table [" + sheetName + " ] successfuly saved!");

        }
    }
}
