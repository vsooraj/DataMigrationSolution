using DataMigrationSolution.BL;
using DataMigrationSolution.Library;
using System;
using System.Windows.Forms;

namespace DataMigrationSolution.Windows
{
    public partial class FrmMain : Form
    {
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

        private void btnAccounts_Click(object sender, EventArgs e)
        {

        }

        private void btnUsers_Click(object sender, EventArgs e)
        {
            var UserRepository = new UserRepository();
            dataGridView1.DataSource = UserRepository.LoadAll();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            var exportService = new ExportService();
            exportService.ExportToExcel();
        }
    }
}
