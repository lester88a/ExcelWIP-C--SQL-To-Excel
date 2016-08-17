using System;
using System.Data.SqlClient;
using System.Linq;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelWIP
{
    public class DataTable
    {
        //public & private instance variables
        private string FileName { get; set; }
        private string _ConnectionString { get; set; }
        private Excel.Application _Excel { get; set; }
        private Excel._Workbook _WorkBook { get; set; }
        private Excel._Worksheet WorkSheet { get; set; }
        private SQL.DataTable dataTable { get; set; }

        public virtual void CreateExcel()
        {
            //get connection
            Connection con = new Connection();
            this.FileName = con.GetFileName();
            _ConnectionString = con.GetConnectionString();
            
            /*-------------------------*/
            //Create Excel objects
            #region Create Excel objects

            _Excel = new Excel.Application();
            _Excel.Visible = true;

            _WorkBook = _Excel.Workbooks.Add(Missing.Value);

            WorkSheet = _WorkBook.ActiveSheet;
            #endregion
            
        }
        
        //virtual method
        public virtual void CreatSheet(string query, string sheetName)
        {
            /*-------------------------*/
            #region Read data from SQL Server
            dataTable = new SQL.DataTable();

            using (SqlConnection sqlConnection = new SqlConnection(_ConnectionString))
            {
                using (SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(query.ToString(), sqlConnection))
                {
                    sqlDataAdapter.Fill(dataTable);
                }
            }
            #endregion

            /*-------------------------*/
            //create work sheet name
            #region create a woork sheet by name
            WorkSheet = (Excel._Worksheet)_Excel.Worksheets.Add();
            WorkSheet.Name = sheetName;
            #endregion

            /*-------------------------*/
            //Add column names to excel sheet
            #region add sheet header and set sheet width and font
            string[] colNames = new string[dataTable.Columns.Count];
            int col = 0;
            //fetch column names from dtData
            foreach (SQL.DataColumn dc in dataTable.Columns)
            {
                colNames[col++] = dc.ColumnName;
            }

            //last column for english alphabet
            char lastColumn = (char)(65 + dataTable.Columns.Count - 1);

            //assign data to column headers
            WorkSheet.get_Range("A1", lastColumn + "1").Value2 = colNames;

            //set width
            WorkSheet.Columns.AutoFit();

            //set column headers' font to bold
            WorkSheet.get_Range("A1", lastColumn + "1").Font.Bold = true;
            #endregion

            /*-------------------------*/
            // Add DataRows data to Excel
            #region Add DataRows data to Excel
            string data = null;

            for (int i = 0; i <= dataTable.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dataTable.Columns.Count - 1; j++)
                {
                    data = dataTable.Rows[i].ItemArray[j].ToString();
                    WorkSheet.Cells[i + 2, j + 1] = data;
                }
            }
            #endregion
            Console.WriteLine("Sheet [{0}] created!",sheetName);
            
        }

        public virtual void SaveExcel()
        {
            /*-------------------------*/
            //Save Data Excel sheet
            _Excel.Visible = true;
            _Excel.DisplayAlerts = false;

            //save without prompt
            _Excel.UserControl = true;
            _WorkBook.SaveAs(this.FileName, AccessMode: Excel.XlSaveAsAccessMode.xlExclusive);
            _WorkBook.Close();
            _Excel.Quit();
            Console.WriteLine("Success saved Excel to:" + FileName);
        }

    }
}
