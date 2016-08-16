using System;
using System.Data.SqlClient;
using System.Linq;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelWIP
{
    public class WIP
    {
        /*------------------------------------------------------------------------------------------------*/
        //convert sql to excel
        public void SQLToExcel(string fileName, string connectionString, string query)
        {
            /*-------------------------*/
            //Read data from SQL Server
            SQL.DataTable dtData = new SQL.DataTable();

            using (SqlConnection sqlConnection = new SqlConnection(connectionString))
            {
                using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(query.ToString(), sqlConnection))
                {
                    sqlDataAdapter.Fill(dtData);
                }
            }
            /*-------------------------*/
            //Create Excel objects
            Excel.Application excel;
            Excel._Workbook workBook;
            Excel._Worksheet workSheet;

            excel = new Excel.Application();
            excel.Visible = true;

            workBook = excel.Workbooks.Add(Missing.Value);
            workSheet = workBook.ActiveSheet;

            /*-------------------------*/
            //create work sheet name
            workSheet = (Excel._Worksheet)excel.Worksheets.Add();
            workSheet.Name = "WIP";

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
            workSheet.get_Range("B1", "B1").ColumnWidth = 7;
            workSheet.get_Range("C1", "C1").ColumnWidth = 20;
            workSheet.get_Range("D1", "D1").ColumnWidth = 8;
            workSheet.get_Range("E1", "E1").ColumnWidth = 13;
            workSheet.get_Range("F1", "F1").ColumnWidth = 21;
            workSheet.get_Range("G1", "G1").ColumnWidth = 13;
            workSheet.get_Range("H1", "H1").ColumnWidth = 21;
            workSheet.get_Range("I1", "I1").ColumnWidth = 13;
            workSheet.get_Range("J1", "J1").ColumnWidth = 21;
            workSheet.get_Range("K1", "K1").ColumnWidth = 21;
            workSheet.get_Range("L1", "L1").ColumnWidth = 21;
            workSheet.get_Range("M1", "M1").ColumnWidth = 21;
            workSheet.get_Range("N1", "N1").ColumnWidth = 10;
            workSheet.get_Range("O1", "O1").ColumnWidth = 10;
            workSheet.get_Range("P1", "P1").ColumnWidth = 17;
            workSheet.get_Range("Q1", "Q1").ColumnWidth = 16;
            workSheet.get_Range("R1", "R1").ColumnWidth = 13;
            workSheet.get_Range("S1", "S1").ColumnWidth = 12;
            workSheet.get_Range("T1", "T1").ColumnWidth = 18;
            workSheet.get_Range("U1", "U1").ColumnWidth = 13;
            workSheet.get_Range("V1", "V1").ColumnWidth = 13;
            workSheet.get_Range("W1", "W1").ColumnWidth = 17;

            //column name's font
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
            //make sure the result of query has at lest one row
            if (true)
            {
                //assign row data
                workSheet.get_Range("A2", lastColumn + rowCount.ToString()).Value2 = rowData;
            }
            

            /*-------------------------*/
            //get dateIn tab
            DateIn dataIn = new DateIn();
            Connection co = new Connection();
            dataIn.SQLToExcel(fileName, connectionString, co.GetQueryIn(), excel, workBook);
            /*-------------------------*/
            //get dateOut tab
            DateOut dateOut = new DateOut();
            dateOut.SQLToExcel(fileName, connectionString, co.GetQueryOut(), excel, workBook);
            /*-------------------------*/
            //get TAT tab
            TAT tat = new TAT();
            tat.SQLToExcel(fileName, connectionString, co.GetQueryTAT(), excel, workBook);
            /*-------------------------*/
            //get BackOrder tab
            BackOrder backOrder = new BackOrder();
            backOrder.SQLToExcel(fileName, connectionString, co.GetQueryBackOrder(), excel, workBook);
            /*-------------------------*/
            //get Bulk tab
            Bulk bulk = new Bulk();
            bulk.SQLToExcel(fileName, connectionString, co.GetQueryBulk(), excel, workBook);

           


            /*-------------------------*/
            //Save Data Excel sheet
            excel.Visible = true;
            excel.DisplayAlerts = false;


            //save without prompt
            excel.UserControl = true;
            workBook.SaveAs(fileName, AccessMode: Excel.XlSaveAsAccessMode.xlExclusive);
            workBook.Close();
            excel.Quit();
            Console.WriteLine("Success saved!");
        }

    }
}
