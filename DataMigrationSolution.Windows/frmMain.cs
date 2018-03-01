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
            LoadUsers();
            HideSubItems();

        }

        private void HideSubItems()
        {
            btnUsers.Visible = false;
            button17.Visible = false;
            button15.Visible = false;
            button16.Visible = false;
            btnCampaigns.Visible = false;
            button14.Visible = false;
            button12.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button11.Visible = false;
            button10.Visible = false;
            button9.Visible = false;
            button8.Visible = false;
            button2.Visible = false;
            button1.Visible = false;
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void btnUsers_Click(object sender, EventArgs e)
        {
            LoadUsers();
        }

        private void LoadUsers()
        {
            exportName = "Users";
            displaymessage(exportName);
            var userRepository = new UserRepository();
            dataGridView1.DataSource = userRepository.LoadAll();
        }

        private void displaymessage(string info)
        {
            lblMessage.Text = $"{info} data loading";
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
            displaymessage(exportName);
            var accountRepository = new AccountRepository();
            dataGridView1.DataSource = accountRepository.LoadAll();
        }

        private void btnSales_Click(object sender, EventArgs e)
        {
            button2.Visible = true;
            button1.Visible = true;
            btnUsers.Visible = true;
            btnUsers.Location = new System.Drawing.Point(17, 85);
        }

        private void btnMarketing_Click(object sender, EventArgs e)
        {
            HideSubItems();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            HideSubItems();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            HideSubItems();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            MessageBox.Show("This functionality yet to implement");
        }
    }
}
