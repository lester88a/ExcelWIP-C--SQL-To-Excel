using System;
using System.Data.SqlClient;
using System.Linq;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWIP
{
    public class Bulk
    {
        /*------------------------------------------------------------------------------------------------*/
        //convert sql to excel
        public void SQLToExcel(string fileName, string connectionString, string queryBulk, Excel.Application oExcel, Excel._Workbook oWorkBook)
        {
            /*-------------------------*/
            //Read data from SQL Server
            SQL.DataTable dataTable = new SQL.DataTable();

            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(queryBulk.ToString(), sqlConnection))
                {
                    sqlDataAdapter.Fill(dataTable);
                }
            }
            /*-------------------------*/
            //Create Excel objects
            Excel._Worksheet workSheet;

            workSheet = oWorkBook.ActiveSheet;

            /*-------------------------*/
            //create work sheet name
            workSheet = (Excel._Worksheet)oExcel.Worksheets.Add();
            workSheet.Name = "Bulk";

            /*-------------------------*/
            //Add column names to excel sheet
            string[] colNames = new string[dataTable.Columns.Count];
            int col = 0;
            //fetch column names from dtData
            foreach (SQL.DataColumn dc in dataTable.Columns)
            {
                colNames[col++] = dc.ColumnName;
            }
            //last column for english alphabet
            char lastColumn = (char)(65 + dataTable.Columns.Count - 1);
            Console.WriteLine(lastColumn);
            //assign column names
            workSheet.get_Range("A1", lastColumn + "1").Value2 = colNames;
            //set width of columns
            workSheet.get_Range("A1", "A1").ColumnWidth = 12;
            workSheet.get_Range("B1", "B1").ColumnWidth = 8;
            workSheet.get_Range("C1", "C1").ColumnWidth = 12;
            workSheet.get_Range("D1", "D1").ColumnWidth = 15;

            //column name's font
            workSheet.get_Range("A1", lastColumn + "1").Font.Bold = true;

            /*-------------------------*/
            //Add DataRows data to Excel
            SQL.DataRow[] dataRow = dataTable.Select();
            //two dimensional string array
            string[,] rowData = new string[dataRow.Count<SQL.DataRow>(), dataTable.Columns.Count];
            int rowCount = 0;

            foreach (SQL.DataRow row in dataRow)
            {
                for (int i = 0; i < dataTable.Columns.Count; i++)
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
