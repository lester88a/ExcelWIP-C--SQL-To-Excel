using System;
using System.Data.SqlClient;
using System.Linq;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelWIP
{
    public class BackOrder
    {
        /*------------------------------------------------------------------------------------------------*/
        //convert sql to excel
        public void SQLToExcel(string fileName, string connectionString, string queryTAT, Excel.Application oExcel, Excel._Workbook oWorkBook)
        {
            /*-------------------------*/
            //Read data from SQL Server
            SQL.DataTable dtData = new SQL.DataTable();

            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(queryTAT.ToString(), sqlConnection))
                {
                    sqlDataAdapter.Fill(dtData);
                }
            }
            /*-------------------------*/
            //Create Excel objects
            Excel._Worksheet workSheet;

            workSheet = oWorkBook.ActiveSheet;

            /*-------------------------*/
            //create work sheet name
            workSheet = (Excel._Worksheet)oExcel.Worksheets.Add();
            workSheet.Name = "BackOrder";

            /*-------------------------*/
            //Add column names to excel sheet
            string[] colNames = new string[dtData.Columns.Count];
            int col = 0;
            //fetch column names from dtData
            foreach (SQL.DataColumn dc in dtData.Columns)
            {
                colNames[col++] = dc.ColumnName;
            }
            //last column for english alphabet
            char lastColumn = (char)(65 + dtData.Columns.Count - 1);
            Console.WriteLine(lastColumn);
            //assign column names
            workSheet.get_Range("A1", lastColumn + "1").Value2 = colNames;
            //set width of columns
            workSheet.get_Range("A1", "A1").ColumnWidth = 12;
            workSheet.get_Range("B1", "B1").ColumnWidth = 8;
            workSheet.get_Range("C1", "C1").ColumnWidth = 16;
            workSheet.get_Range("D1", "D1").ColumnWidth = 21;
            workSheet.get_Range("E1", "E1").ColumnWidth = 21;
            workSheet.get_Range("F1", "F1").ColumnWidth = 15;
            workSheet.get_Range("G1", "G1").ColumnWidth = 15;
            workSheet.get_Range("H1", "H1").ColumnWidth = 15;
            workSheet.get_Range("I1", "I1").ColumnWidth = 8;

            //column header's font - from first to last
            workSheet.get_Range("A1", lastColumn + "1").Font.Bold = true;

            /*-------------------------*/
            //Add DataRows data to Excel
            SQL.DataRow[] dataRow = dtData.Select();
            //two dimensional string array
            string[,] rowData = new string[dataRow.Count<SQL.DataRow>(), dtData.Columns.Count];
            int rowCount = 0;

            foreach (SQL.DataRow row in dataRow)
            {
                for (int i = 0; i < dtData.Columns.Count; i++)
                {
                    //two dimensional string array
                    rowData[rowCount, i] = row[i].ToString();
                }
                //increase rowCount by 1
                rowCount++;
            }
            //make sure the query has at lest one row
            if (rowCount>0)
            {
                //assign row data
                workSheet.get_Range("A2", lastColumn + rowCount.ToString()).Value2 = rowData;
            }
            

        }
    }
}
