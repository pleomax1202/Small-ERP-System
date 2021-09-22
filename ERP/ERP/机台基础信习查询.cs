using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace Combination
{
    public partial class 机台基础信习查询 : Form
    {
        Sql sql = new Sql();

        public 机台基础信习查询()
        {
            InitializeComponent();
        }

        public void Load_Data()
        {
            dataGridView1.Rows.Clear();
            DataTable dt = new DataTable();
            dt = sql.getQuery(@"SELECT * FROM [dbo].[Machine] Order by Mcode");

            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = item["Mcode"].ToString();
                dataGridView1.Rows[n].Cells[1].Value = item["Mname"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = item["MInUnit"].ToString();
                dataGridView1.Rows[n].Cells[3].Value = item["MOutUnit"].ToString();
                dataGridView1.Rows[n].Cells[4].Value = item["MUnit"].ToString();
                dataGridView1.Rows[n].Cells[5].Value = item["MSpeed"].ToString();
                dataGridView1.Rows[n].Cells[6].Value = item["MWUnit"].ToString();
                dataGridView1.Rows[n].Cells[7].Value = item["MHour"].ToString();
                dataGridView1.Rows[n].Cells[9].Value = item["MOrder"].ToString();

                DataTable dtWFlow = new DataTable();
                dtWFlow = sql.getQuery(@"SELECT WName FROM [ChengyiYuntech].[dbo].[WorkFlow] WHERE ID = '" + item["MOrder"].ToString() + "'");

                dataGridView1.Rows[n].Cells[8].Value = dtWFlow.Rows[0][0].ToString();
            }  
        }

        public void button1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            if (textBox1.Text == "" && textBox2.Text == "")
            {
                dataGridView1.Rows.Clear();
                Load_Data();
            }
            else
            {
                dataGridView1.Rows.Clear();
                DataTable dt = new DataTable();
                dt = sql.getQuery(@"SELECT * FROM [dbo].[Machine] WHERE [Mcode] like '%" + textBox1.Text + "%' And [Mname] like '%" + textBox2.Text + "%' Order by Mcode");

                foreach (DataRow item in dt.Rows)
                {
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = item["Mcode"].ToString();
                    dataGridView1.Rows[n].Cells[1].Value = item["Mname"].ToString();
                    dataGridView1.Rows[n].Cells[2].Value = item["MInUnit"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["MOutUnit"].ToString();
                    dataGridView1.Rows[n].Cells[4].Value = item["MUnit"].ToString();
                    dataGridView1.Rows[n].Cells[5].Value = item["MSpeed"].ToString();
                    dataGridView1.Rows[n].Cells[6].Value = item["MWUnit"].ToString();
                    dataGridView1.Rows[n].Cells[7].Value = item["MHour"].ToString();
                }
            }

            Cursor = Cursors.Default;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                int index = dataGridView1.CurrentRow.Index;    //取得选中行的索引
                machineInfoEdit mie = new machineInfoEdit(dataGridView1, index, button1);
                mie.Show();
            }
            else
            {
                MessageBox.Show("请选择要编辑的行");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            machineInfoEdit2 mie2 = new machineInfoEdit2(button1);
            mie2.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                DialogResult result = MessageBox.Show("请确认是否删除该行", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    int index = dataGridView1.CurrentRow.Index;
                    var SqlQuery = @"DELETE FROM [dbo].[Machine] WHERE [Mcode] = '" + dataGridView1.Rows[index].Cells[0].Value + "'";
                    sql.sqlCmd(SqlQuery);
                    MessageBox.Show("该行已成功移除");
                    Load_Data();
                }
            }
            else
            {
                MessageBox.Show("请选择要删除的行");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Export_Data();
            }
            catch (Exception)
            {
                MessageBox.Show("请将先前导出关闭");
            }
        }

        private void Export_Data()
        {
            if (dataGridView1.Rows.Count != 0)
            {
                Excel.Application excelApp;
                Excel._Workbook wBook;
                Excel._Worksheet wSheet;
                Excel.Range wRange;

                string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string inputPath = System.Environment.CurrentDirectory;
                string exportPath = path + @"\机台基础信息导出";
                string filePath = inputPath + @"\机台基础信息";

                // 開啟一個新的應用程式
                excelApp = new Excel.Application();

                // 讓Excel文件可見
                excelApp.Visible = false;

                // 停用警告訊息
                excelApp.DisplayAlerts = false;

                // 加入新的活頁簿
                excelApp.Workbooks.Add(Type.Missing);

                wBook = excelApp.Workbooks.Open(filePath,
                        0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                        true, false, 0, true, false, false);

                // 設定活頁簿焦點
                wBook.Activate();

                wSheet = (Excel._Worksheet)wBook.Worksheets[1];

                wSheet.Name = "机台基础信息";

                // storing Each row and column value to excel sheet
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count - 1; j++)
                    {
                        wSheet.Cells[i + 4, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                    }
                }

                Excel.Range last = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range allRange = wSheet.get_Range("A1", last);
                allRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;        //格線
                allRange.Font.Size = "14";
                allRange.Columns.AutoFit();

                //Save Excel
                wBook.SaveAs(exportPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                wBook = null;
                wSheet = null;
                wRange = null;
                excelApp = null;

                GC.Collect();
                MessageBox.Show("导出成功");
            }
            else
            {
                MessageBox.Show("请确认是否有资料");
            }
        }
    }
}
