using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SQLServerData2Excel
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
                Connecting.openCon();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Application.Exit();
                return;
            }
            //Get a DataTable
            string sqlGetAllTable = @"SELECT TABLE_NAME
                                    FROM INFORMATION_SCHEMA.TABLES
                                    WHERE TABLE_TYPE = 'BASE TABLE' 
                                    AND TABLE_CATALOG = 'BDEMIS_Common'";
                DataTable allTableName = Connecting.GetDataTable(sqlGetAllTable);

                //Fill data to combobox
                cbxTableName.DataSource = allTableName;
                cbxTableName.DisplayMember = "TABLE_NAME";
                cbxTableName.ValueMember = "TABLE_NAME";

                //Config combobox
                cbxTableName.DropDownStyle = ComboBoxStyle.DropDownList;
                cbxTableName.SelectedIndex = 1;
                cbxTableName.SelectedIndex = 0;

                //Config datagridview
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                //Config Progressbar
                progressBar1.Maximum = cbxTableName.Items.Count;
            
        }

        private void cbxTableName_SelectedIndexChanged(object sender, EventArgs e)
        {
            string tableName = cbxTableName.SelectedValue.ToString();
            string sqlGetData = string.Format(@"SELECT * FROM {0}",tableName);
            DataTable availableData = Connecting.GetDataTable(sqlGetData);
            dataGridView1.DataSource = availableData;
        }

        private void btnExportSingleTable_Click(object sender, EventArgs e)
        {
            string tableName = cbxTableName.Text;
            ExportToXlsx.SimpleMode(dataGridView1, tableName, progressBar2);
        }

        private void btnExportAllTable_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            FolderBrowserDialog saveDialog = new FolderBrowserDialog();
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                int numberOfTable = cbxTableName.Items.Count;
                for (int i = 0; i < numberOfTable; i++)
                {
                    cbxTableName.SelectedIndex = i;
                    string tableName = cbxTableName.SelectedValue.ToString();
                    ExportToXlsx.SimpleModeMultiFile(dataGridView1, tableName, saveDialog.SelectedPath, progressBar2);
                    progressBar1.Increment(1);
                }
                MessageBox.Show("Export done!");
            }
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
    }
}
