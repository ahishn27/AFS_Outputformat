using System;
using System.Collections.Generic;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text;
using System.Reflection;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;


namespace AFS_Excel_Format
{
     class DBConnect
    {


        static void Main()
        {
             MySqlConnection connection;

              connection = new MySqlConnection("Database = testdb; Port = 3306; Data Source = 127.0.0.1; User Id = root; Password = ; SslMode = none; ");

              string sql = "select Material, Material_Description,Days,Requirements,Closing_Stock,PR_Receipts,PO_Receipts from sleeve;";

            // MySqlCommand cmd = new MySqlCommand(sql, connection);
            MySqlDataAdapter returnVal = new MySqlDataAdapter(sql, connection);

            DataTable dt = new DataTable("ProductInfo");
            returnVal.Fill(dt);

            for (int j = 0; j < dt.Rows.Count; j++)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    Console.Write(dt.Columns[i].ColumnName + " ");
                    Console.WriteLine(dt.Rows[j].ItemArray[i]);
                }
            }

            //Create a XML file
            dt.WriteXml(@"D:\MyDataset.xml");
          

            /*
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            */


            System.IO.StringWriter writer = new System.IO.StringWriter();
            dt.WriteXml(writer, XmlWriteMode.IgnoreSchema, false);
            string result = writer.ToString();
            Console.WriteLine(result);
            

            /* try { 
             using (MySqlDataReader reader = cmd.ExecuteReader())
             {
                 while (reader.Read())
                 {
                     // access your record colums by using reader
                     Console.WriteLine(reader["Product_Name"]);
                 }
             }
         }
         catch (Exception ex)
         {
          // handle exception here
         }
         finally
         {
           connection.Close();
         }       */


            //initializing excel applicaton
            /*Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                           //Check Null
                            if (xlApp == null)
                            {
                            MessageBox.Show("Excel is not properly installed!!");
                             return;
                             }
                              Excel.Workbook xlWorkBook;
                              Excel.Worksheet xlWorkSheet;
                              object misValue = System.Reflection.Missing.Value;
                              xlWorkBook = xlApp.Workbooks.Add(misValue);
                              xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                              xlWorkSheet.Cells[1, 1] = "Product_Id";
                              xlWorkSheet.Cells[1, 2] = "Product_Name";
                              xlWorkSheet.Cells[2, 1] = num;
                              xlWorkSheet.Cells[2, 2] = name;

                               xlWorkBook.SaveAs("d:\\csharp-Exceltest2.xlt", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                               xlWorkBook.Close(true, misValue, misValue);
                               xlApp.Quit();

                               MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.xls");
            */
            /*
                        Excel._Application xlapp;
                        Excel.Workbooks workbooks;
                        Excel.Workbook workbook;
                        object misValue = System.Reflection.Missing.Value;

                        xlapp = new Microsoft.Office.Interop.Excel.Application();
                        xlapp.Visible = true;
                        workbooks = xlapp.Workbooks;
                        String Filename = "D:\\csharp-Exceltest2.xls";
                        workbook = workbooks.Open(Filename);
                        Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)workbook.ActiveSheet;*/

        }

    }
    }
    

