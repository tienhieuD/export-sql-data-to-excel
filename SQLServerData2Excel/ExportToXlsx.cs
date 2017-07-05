using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace SQLServerData2Excel
{
    class ExportToXlsx
    {
        public static void SimpleMode(DataGridView dgvName, string tableName, ProgressBar pbar)
        {
            // Creating a Excel object.
            _Application excel = new Microsoft.Office.Interop.Excel.Application();
            _Workbook workbook = excel.Workbooks.Add(Type.Missing);
            _Worksheet worksheet = null;
            pbar.Value = 0;
            pbar.Maximum = dgvName.RowCount;

            try
            {
                worksheet = workbook.ActiveSheet;

                worksheet.Name = tableName;

                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column.
                for (int i = 0; i < dgvName.Rows.Count+1; i++)
                {
                    for (int j = 0; j < dgvName.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dgvName.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dgvName.Rows[i-1].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                    pbar.Increment(1);
                }

                //Getting the location and file name of the excel to save from user.
                FolderBrowserDialog saveDialog = new FolderBrowserDialog();

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    string savePath = string.Format(@"{0}\{1}{2}", saveDialog.SelectedPath, tableName, ".xlsx");
                    workbook.SaveAs(savePath);
                   
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }

        public static void SimpleModeMultiFile(DataGridView dgvName, string tableName, string folderPath, ProgressBar pbar)
        {
            // Creating a Excel object.
            _Application excel = new Microsoft.Office.Interop.Excel.Application();
            _Workbook workbook = excel.Workbooks.Add(Type.Missing);
            _Worksheet worksheet = null;
            pbar.Value = 0;
            pbar.Maximum = dgvName.RowCount;

            try
            {
                worksheet = workbook.ActiveSheet;

                worksheet.Name = tableName;

                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column.
                for (int i = 0; i < dgvName.Rows.Count + 1; i++)
                {
                    for (int j = 0; j < dgvName.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check.
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dgvName.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dgvName.Rows[i - 1].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                    pbar.Increment(1);
                }

                //Getting the location and file name of the excel to save from user.

                
                    string savePath = string.Format(@"{0}\{1}{2}", folderPath, tableName, ".xlsx");
                    workbook.SaveAs(savePath);
                
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }
        }
    }
}
