using System;
using System.Data.SqlClient;
using System.Linq;
using SQL = System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelWIP
{
    public class CSVFile
    {
        /*------------------------------------------------------------------------------------------------*/
        private static void SQLToCSV(string query, string fileName)
        {

            SqlConnection conn = new SqlConnection("ConnectionString");
            conn.Open();
            SqlCommand cmd = new SqlCommand(query, conn);
            SqlDataReader dr = cmd.ExecuteReader();

            using (System.IO.StreamWriter fs = new System.IO.StreamWriter(fileName))
            {
                // Loop through the fields and add headers
                for (int i = 0; i < dr.FieldCount; i++)
                {
                    string name = dr.GetName(i);
                    if (name.Contains(","))
                        name = "\"" + name + "\"";

                    fs.Write(name + ",");
                }
                fs.WriteLine();

                // Loop through the rows and output the data
                while (dr.Read())
                {
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(","))
                            value = "\"" + value + "\"";

                        fs.Write(value + ",");
                    }
                    fs.WriteLine();
                }



                Console.WriteLine("Success!");
                fs.Close();
            }
        }
        /*------------------------------------------------------------------------------------------------*/
    }
}
