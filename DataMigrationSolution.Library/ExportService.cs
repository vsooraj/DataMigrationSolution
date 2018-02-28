using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.IO;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat;

namespace DataMigrationSolution.Library
{
    public class ExportService
    {
        public void ExportToExcel()
        {
            // SET THE CONNECTION STRING.
            string sCon = "server=localhost;user=root;database=crm;port=3306;password=c@b0t1234";

            using (MySqlConnection con = new MySqlConnection(sCon))
            {
                using (MySqlCommand cmd = new MySqlCommand("SELECT id,user_name FROM vtiger_users"))
                {
                    MySqlDataAdapter sda = new MySqlDataAdapter();
                    try
                    {
                        cmd.Connection = con;
                        con.Open();
                        sda.SelectCommand = cmd;

                        var dt = new System.Data.DataTable();
                        sda.Fill(dt);

                        if (dt.Rows.Count > 0)
                        {
                            string path = "C:\\Users\\sooraj.v\\Desktop\\CRM\\";

                            if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
                            {
                                Directory.CreateDirectory(path);
                            }

                            File.Delete(path + "User.xlsx"); // DELETE THE FILE BEFORE CREATING A NEW ONE.

                            // ADD A WORKBOOK USING THE EXCEL APPLICATION.
                            Microsoft.Office.Interop.Excel.Application xlAppToExport = new Microsoft.Office.Interop.Excel.Application();
                            xlAppToExport.Workbooks.Add("");

                            // ADD A WORKSHEET.
                            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport = default(Microsoft.Office.Interop.Excel.Worksheet);
                            xlWorkSheetToExport = (Microsoft.Office.Interop.Excel.Worksheet)xlAppToExport.Sheets["Sheet1"];

                            // ROW ID FROM WHERE THE DATA STARTS SHOWING.
                            int iRowCnt = 4;

                            // SHOW THE HEADER.
                            xlWorkSheetToExport.Cells[1, 1] = "User Details";

                            Microsoft.Office.Interop.Excel.Range range = xlWorkSheetToExport.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                            range.EntireRow.Font.Name = "Calibri";
                            range.EntireRow.Font.Bold = true;
                            range.EntireRow.Font.Size = 20;

                            xlWorkSheetToExport.Range["A1:D1"].MergeCells = true;       // MERGE CELLS OF THE HEADER.

                            // SHOW COLUMNS ON THE TOP.
                            xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "ID";
                            xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "User";
                            //xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "PresentAddress";
                            //xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Email Address";

                            int i;
                            for (i = 0; i <= dt.Rows.Count - 1; i++)
                            {
                                xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i].Field<System.Int32>("id");
                                xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i].Field<string>("user_name");


                                iRowCnt = iRowCnt + 1;
                            }

                            // FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                            Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                            range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);

                            // SAVE THE FILE IN A FOLDER.
                            xlWorkSheetToExport.SaveAs(path + "User.xlsx");

                            // CLEAR.
                            xlAppToExport.Workbooks.Close();
                            xlAppToExport.Quit();
                            xlAppToExport = null;
                            xlWorkSheetToExport = null;


                            // Console.WriteLine("Data Exported Successfully.");


                        }
                    }
                    catch (Exception ex)
                    {

                    }
                    finally
                    {
                        sda.Dispose();
                        sda = null;
                    }
                }
            }
        }
    }
}
