using DataMigrationSolution.BL;
using DataMigrationSolution.Library;
using System;
using System.Windows.Forms;

namespace DataMigrationSolution.Windows
{
    public partial class FrmMain : Form
    {
        public string exportName;
        public FrmMain()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        //private void btnAccounts_Click(object sender, EventArgs e)
        //{
        //    exportName = "Accounts";
        //    var accountRepository = new AccountRepository();
        //    dataGridView1.DataSource = accountRepository.LoadAll();
        //}

        private void btnUsers_Click(object sender, EventArgs e)
        {
            exportName = "Users";
            var userRepository = new UserRepository();
            dataGridView1.DataSource = userRepository.LoadAll();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            var exportService = new ExportService();
            if (exportName != null)
                exportService.Export(exportName);
            else
                MessageBox.Show($"{exportName} is null,please check");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            exportName = "Accounts";
            var accountRepository = new AccountRepository();
            dataGridView1.DataSource = accountRepository.LoadAll();
        }
    }
}
