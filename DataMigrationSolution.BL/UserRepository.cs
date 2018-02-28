using MySql.Data.MySqlClient;
using System.Collections.Generic;

namespace DataMigrationSolution.BL
{
    public class UserRepository
    {
        static string connStr = "server=localhost;user=root;database=crm;port=3306;password=c@b0t1234";

        public IEnumerable<User> LoadAll()
        {
            var users = new List<User>();
            using (var conn = new MySqlConnection(connStr))
            {
                //Console.WriteLine("Connecting to MySQL...");
                conn.Open();
                string sql = "SELECT * FROM vtiger_users";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    users.Add(
                    new User
                    {
                        Id = (int)rdr[0],
                        FirstName = rdr[1].ToString()
                    });

                }
                rdr.Close();
            }
            return users;
        }

    }


}
