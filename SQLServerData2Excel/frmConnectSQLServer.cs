using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SQLServerData2Excel
{
    public partial class frmConnectSQLServer : Form
    {
        public frmConnectSQLServer()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string newConnectionString = textBox1.Text;
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var connectionStringsSection = (ConnectionStringsSection)config.GetSection("connectionStrings");
            connectionStringsSection.ConnectionStrings["ConnectionStringBDEMIS_Common"].ConnectionString = newConnectionString;
            config.Save();
            ConfigurationManager.RefreshSection("connectionStrings");
            frmConnectSQLServer_Load(sender, e);
            this.Hide();
            new frmMain().Show();
        }

        private void frmConnectSQLServer_Load(object sender, EventArgs e)
        {
            textBox1.Text = ConfigurationManager.ConnectionStrings["ConnectionStringBDEMIS_Common"].ConnectionString;
        }
    }
}
