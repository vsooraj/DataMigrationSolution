using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.IO;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat;

namespace DataMigrationSolution.Library
{
    public class ExportService
    {
        string query = "SELECT id, user_name FROM vtiger_users";
        string sCon = "server=localhost;user=root;database=crm;port=3306;password=c@b0t1234";
        public void Export(string name)
        {
            switch (name)
            {
                case "Users":
                    query = "SELECT id,user_name FROM vtiger_users";
                    break;
                case "Accounts":
                    query = "SELECT accountid,accountname,industry FROM vtiger_account";
                    break;
                default:
                    break;
            }
            // SET THE CONNECTION STRING.


            using (MySqlConnection con = new MySqlConnection(sCon))
            {
                using (MySqlCommand cmd = new MySqlCommand(query))
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

                            File.Delete(path + $"{name}.xlsx"); // DELETE THE FILE BEFORE CREATING A NEW ONE.

                            // ADD A WORKBOOK USING THE EXCEL APPLICATION.
                            Microsoft.Office.Interop.Excel.Application xlAppToExport = new Microsoft.Office.Interop.Excel.Application();
                            xlAppToExport.Workbooks.Add("");

                            // ADD A WORKSHEET.
                            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport = default(Microsoft.Office.Interop.Excel.Worksheet);
                            xlWorkSheetToExport = (Microsoft.Office.Interop.Excel.Worksheet)xlAppToExport.Sheets["Sheet1"];

                            // ROW ID FROM WHERE THE DATA STARTS SHOWING.
                            int iRowCnt = 4;

                            // SHOW THE HEADER.
                            xlWorkSheetToExport.Cells[1, 1] = $"{name} Details";

                            Microsoft.Office.Interop.Excel.Range range = xlWorkSheetToExport.Cells[1, 1] as Microsoft.Office.Interop.Excel.Range;
                            range.EntireRow.Font.Name = "Calibri";
                            range.EntireRow.Font.Bold = true;
                            range.EntireRow.Font.Size = 20;

                            xlWorkSheetToExport.Range["A1:D1"].MergeCells = true;       // MERGE CELLS OF THE HEADER.

                            // SHOW COLUMNS ON THE TOP.
                            switch (name)
                            {
                                case "Users":
                                    iRowCnt = getUser(dt, xlWorkSheetToExport, iRowCnt);
                                    break;
                                case "Accounts":
                                    iRowCnt = getAccount(dt, xlWorkSheetToExport, iRowCnt);
                                    break;
                                default:
                                    break;
                            }


                            // FINALLY, FORMAT THE EXCEL SHEET USING EXCEL'S AUTOFORMAT FUNCTION.
                            Microsoft.Office.Interop.Excel.Range range1 = xlAppToExport.ActiveCell.Worksheet.Cells[4, 1] as Microsoft.Office.Interop.Excel.Range;
                            range1.AutoFormat(ExcelAutoFormat.xlRangeAutoFormatList3);

                            // SAVE THE FILE IN A FOLDER.
                            xlWorkSheetToExport.SaveAs(path + $"{name}.xlsx");

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

        private static int getUser(DataTable dt, Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport, int iRowCnt)
        {
            xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "ID";
            xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "User";

            int i;
            for (i = 0; i <= dt.Rows.Count - 1; i++)
            {
                xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i].Field<System.Int32>("id");
                xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i].Field<string>("user_name");


                iRowCnt = iRowCnt + 1;
            }

            return iRowCnt;
        }
        private static int getAccount(DataTable dt, Microsoft.Office.Interop.Excel.Worksheet xlWorkSheetToExport, int iRowCnt)
        {
            xlWorkSheetToExport.Cells[iRowCnt - 1, 1] = "ID";
            xlWorkSheetToExport.Cells[iRowCnt - 1, 2] = "Account";
            xlWorkSheetToExport.Cells[iRowCnt - 1, 3] = "Industry";
            //xlWorkSheetToExport.Cells[iRowCnt - 1, 4] = "Email Address";

            int i;
            for (i = 0; i <= dt.Rows.Count - 1; i++)
            {
                xlWorkSheetToExport.Cells[iRowCnt, 1] = dt.Rows[i].Field<System.Int32>("accountid");
                xlWorkSheetToExport.Cells[iRowCnt, 2] = dt.Rows[i].Field<string>("accountname");
                xlWorkSheetToExport.Cells[iRowCnt, 3] = dt.Rows[i].Field<string>("industry");


                iRowCnt = iRowCnt + 1;
            }

            return iRowCnt;
        }
    }
}
