﻿using System;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelWIP.VersionTwo
{
    public class DataTableManager
    {
        //public & private instance variables
        public string FileName { get; set; }
        private Excel.Application _Excel { get; set; }
        private Excel._Workbook _WorkBook { get; set; }
        private Excel._Worksheet WorkSheet { get; set; }
        
        //virtual method
        public virtual void CreateExcel(ref string fileName)
        {
            //get fileName by connection
            Connection con = new Connection();
            this.FileName = con.GetFileName();
            fileName = this.FileName;
            /*-------------------------*/
            //Create Excel objects
            #region Create Excel objects

            _Excel = new Excel.Application();
            //hide excel windwow when generating
            _Excel.Visible = false;
            _Excel.DisplayAlerts = false;

            _WorkBook = _Excel.Workbooks.Add(Missing.Value);

            WorkSheet = _WorkBook.ActiveSheet;
            #endregion

        }

        //virtual method
        public virtual void CreatSheet(SQL.DataTable dataTable, string sheetName)
        {
            /*-------------------------*/
            //create work sheet name
            #region create a woork sheet by name
            WorkSheet = (Excel._Worksheet)_Excel.Worksheets.Add();
            WorkSheet.Name = sheetName;
            Console.WriteLine("\n----------------------------------------");
            Console.WriteLine("Create a new work sheet[{0}]", sheetName);
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
            if (false)
            {
                #region Old method for assign data to worksheet --- Faster but missing last row

                SQL.DataTable dtData = dataTable;

                SQL.DataRow[] dataRow = dtData.Select();
                
                int totalRows = dataRow.Count<SQL.DataRow>();
                int totalColumns = dtData.Columns.Count;

                Console.WriteLine("totalRows:" + totalRows);
                Console.WriteLine("totalColumns:" + totalColumns);

                //two dimensional string array
                string[,] rowData = new string[totalRows, dtData.Columns.Count];
               
                int rowCount = 0;
                for (int r = 0; r < totalRows; r++)
                {
                    for (int c = 0; c < totalColumns; c++)
                    {
                        //two dimensional string array
                        rowData[rowCount, c] = dataRow[r][c].ToString();
                    }
                    //increase rowCount by 1
                    rowCount++;
                }

                //foreach (SQL.DataRow row in dataRow)
                //{
                //    for (int i = 0; i < dtData.Columns.Count; i++)
                //    {
                //        //two dimensional string array
                //        rowData[rowCount, i] = row[i].ToString();
                //    }
                //    //increase rowCount by 1
                //    rowCount++;
                //}
                //make sure the query has result
                if (rowCount > 0)
                {
                    //assign row data
                    WorkSheet.get_Range("A2", lastColumn + rowCount.ToString()).Value2 = rowData;
                }
                #endregion
            }
            else
            {
                #region New method for add DataRows data to Excel --- Slower

                Console.WriteLine("Generating data for Sheet[{0}]...",sheetName);
                if (sheetName == "WIP")
                {
                    Console.WriteLine("This will take 2 to 5 minutes!\nDo not close it!");
                }
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
            }
            
            Console.WriteLine("Sheet [{0}] created!", sheetName);
            Console.WriteLine("----------------------------------------\n");
        }

        public virtual void SaveExcel()
        {
            /*-------------------------*/
            //Save Data Excel sheet
            #region Save Data to Excel sheet
            _Excel.Visible = false;
            _Excel.DisplayAlerts = false;

            //save without prompt
            _Excel.UserControl = true;
            _WorkBook.SaveAs(this.FileName, AccessMode: Excel.XlSaveAsAccessMode.xlExclusive);
            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(WorkSheet);
            _WorkBook.Close();
            Marshal.ReleaseComObject(_WorkBook);
            _Excel.Quit();
            Marshal.ReleaseComObject(_Excel);
            Console.WriteLine("Success saved Excel to:" + FileName);
            #endregion
        }
    }
}
