using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace Combination
{
    public partial class 温控报表 : Form
    {
        public 温控报表()
        {
            InitializeComponent();
        }

        Sql sql = new Sql();

        private void Load_Data()
        {
            //this.dataGridView1.Rows.Clear();

            string showTable = "Select b.area,a.checktime,b.temperature,a.temperature,  " +
                "case when ABS(cast(b.temperature as float)-cast(a.temperature as float) ) > cast(b.temperaturechange as float) then '1' else '0'end as 溫度檢測,  case when ABS(cast(b.wet as float)-cast(a.wet as float)) > cast(b.wetchange as float) then '1' else '0'end as 濕度檢測,b.wet,a.wet  FROM [chengyifbsscctest].[dbo].[temperature] a,[chengyifbsscctest].[dbo].[temperatureinfo] b  where a.Machineno = b.Machineno " +
                "and convert(varchar(10),a.checktime,120) = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "' order by checktime";

            DataSet dsMainTable;
            SqlServer sql = new SqlServer();
            sql.Connect("ChengyiYuntech");
            dsMainTable = sql.SqlCmd(showTable);
            dataGridView1.DataSource = dsMainTable.Tables[0];
            dataGridView1.Columns[0].HeaderText = "监控区域";
            dataGridView1.Columns[1].HeaderText = "日期";
            dataGridView1.Columns[2].HeaderText = "标准温度";
            dataGridView1.Columns[3].HeaderText = "检测温度";
            dataGridView1.Columns[6].HeaderText = "标准湿度";
            dataGridView1.Columns[7].HeaderText = "检测湿度";

            int tem;
            int wet;
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                tem = int.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString());
                wet = int.Parse(dataGridView1.Rows[i].Cells[5].Value.ToString());
                if (tem == 1)
                    dataGridView1.Rows[i].Cells[3].Style.BackColor = Color.LightGoldenrodYellow;
                if (wet == 1)
                    dataGridView1.Rows[i].Cells[7].Style.BackColor = Color.LightGoldenrodYellow;
            }
            dataGridView1.Columns[4].Visible = false;
            dataGridView1.Columns[5].Visible = false;
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Load_Data();
        }

        private void btnSetting_Click(object sender, EventArgs e)
        {
            tempBasicSetting tbsc = new tempBasicSetting();
            tbsc.Show();
        }

        private void btnExport_Click(object sender, EventArgs e)
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
                string exportPath = path + @"\温控报表导出";
                string filePath = inputPath + @"\温控报表";

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

                wSheet.Name = "温控报表";

                wSheet.Cells[1, 2] = dateTimePicker1.Value.ToString("yyyy/MM/dd");
                // storing Each row and column value to excel sheet
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    wSheet.Cells[i + 7, 1] = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                    wSheet.Cells[i + 7, 2] = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                    wSheet.Cells[i + 7, 3] = Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
                    wSheet.Cells[i + 7, 4] = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
                    wSheet.Cells[i + 7, 5] = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);
                    wSheet.Cells[i + 7, 6] = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value);
                }

                Excel.Range last = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range allRange = wSheet.get_Range("A7", last);
                allRange.Font.Size = "14";
                allRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;        //格線
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

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                string execute = Convert.ToString(this.dataGridView1.Rows[e.RowIndex].Cells["Column5"].Value);
                if (Convert.ToString(execute) == "异常")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
                string execute2 = Convert.ToString(this.dataGridView1.Rows[e.RowIndex].Cells["Column8"].Value);
                if (Convert.ToString(execute2) == "异常")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
        }
    }
}
