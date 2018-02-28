using MySql.Data.MySqlClient;
using System.Collections.Generic;

namespace DataMigrationSolution.BL
{
    public class AccountRepository
    {
        static string connStr = "server=localhost;user=root;database=crm;port=3306;password=c@b0t1234";

        public IEnumerable<Account> LoadAll()
        {
            var accounts = new List<Account>();
            using (var conn = new MySqlConnection(connStr))
            {
                conn.Open();
                string sql = "SELECT * FROM vtiger_account";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    accounts.Add(
                    new Account
                    {
                        accountid = (int)rdr[0],
                        accountname = rdr[1].ToString(),
                        industry = rdr[4].ToString()
                    });

                }
                rdr.Close();
            }
            return accounts;
        }

    }


}