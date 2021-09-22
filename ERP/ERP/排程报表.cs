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
using Combination.Detail;

namespace Combination
{
    public partial class 排程报表 : Form
    {
        string sRole;
        string ID;
        string temp;
        DataSet dsMainTable;
        DataSet dsProductDetail;
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;
        string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string inputPath = System.Environment.CurrentDirectory;
        string exportPath;
        string filePath;
        Sql sql = new Sql();

        public 排程报表(string role)
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
            this.dataGridView2.ShowCellToolTips = false;
            this.toolTip1.AutomaticDelay = 0;
            this.toolTip1.OwnerDraw = true;
            this.toolTip1.ShowAlways = true;
            this.toolTip1.ToolTipTitle = " ";
            this.toolTip1.UseAnimation = true;
            this.toolTip1.UseFading = true;
            sRole = role;
        }

        private void 排程报表_Load(object sender, EventArgs e)
        {
            if (sRole == "权限管理员" || sRole == "系统管理员")
            {
                this.btnEdit.Enabled = true;
            }
            else
            {
                this.btnEdit.Enabled = false;
            }
        }

        private void Load_Data()
        {
            string dateTimeNow = DateTime.Now.ToString("yyyyMMdd");
            dataGridView1.Rows.Clear();
            DataTable dt = new DataTable();
            dt = sql.getQuery(@"SELECT a.*, b.Mname FROM [ChengyiYuntech].[dbo].[ProduceOrder] a, [dbo].[Machine] b WHERE a.[OMachineCode] = b.[Mcode] and a.OStatus = 0 and a.[ODate] = '" + dateTimeNow + "'");

            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = Convert.ToDateTime(item["ODate"]).ToString("yyyy/MM/dd");
                dataGridView1.Rows[n].Cells[1].Value = item["OOrder"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = item["OMachineCode"].ToString();
                dataGridView1.Rows[n].Cells[3].Value = item["Mname"].ToString();
                dataGridView1.Rows[n].Cells[4].Value = item["OID"].ToString();
                dataGridView1.Rows[n].Cells[5].Value = item["Ohour"].ToString();
                dataGridView1.Rows[n].Cells[6].Value = item["OPName"].ToString();
                dataGridView1.Rows[n].Cells[9].Value = item["ID"].ToString();
                if (item["OStatus"].ToString() == "0")
                {
                    dataGridView1.Rows[n].Cells[8].Value = "未执行";
                }
                else if (item["OStatus"].ToString() == "1")
                {
                    dataGridView1.Rows[n].Cells[8].Value = "已执行";
                }


                DataTable dtStaff = new DataTable();
                dtStaff = sql.getQuery("SELECT [SID] ,[SName] FROM [ChengyiYuntech].[dbo].[Staff] WHERE SID = '" + item["OStaff"].ToString() + "'");

                dataGridView1.Rows[n].Cells[7].Value = dtStaff.Rows[0][1].ToString();
            }
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            if (checkBox.Checked == false)
            {
                dataGridView1.Rows.Clear();
                if (comboBox1.SelectedIndex == 2)
                {
                    temp = "%班%";
                    SeekData(temp);
                }

                if (comboBox1.SelectedIndex == 0 || comboBox1.SelectedIndex == 1)
                {
                    SeekData(comboBox1.Text);
                }
            }
            else
            {
                format2();
            }
            Cursor = Cursors.Default;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if(checkBox.Checked == false)
            {
                if(dataGridView1.Rows.Count != 0)
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
                else if(dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("无信息导出");
                }
            }
            else if (checkBox.Checked == true)
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
        }

        private void Export_Data()
        {
            path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            inputPath = System.Environment.CurrentDirectory;

            if (checkBox.Checked == false)
            {
                exportPath = path + @"\排程报表导出";
                filePath = inputPath + @"\排程报表";
            }
            else if (checkBox.Checked == true)
            {
                exportPath = path + @"\排程报表-格式2导出";
                filePath = inputPath + @"\排程报表-格式2";
            }


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

            wSheet.Name = "排程报表";

            // storing Each row and column value to excel sheet
            if (checkBox.Checked == true)
            {
                for (int i = 0; i < dataGridView2.Columns.Count; i++)
                {
                    wSheet.Cells[1, i + 1] = dataGridView2.Columns[i].HeaderText;
                }
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        wSheet.Cells[i + 2, j + 1] = Convert.ToString(dataGridView2.Rows[i].Cells[j].Value);
                    }
                }
            }
            else if (checkBox.Checked == false)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count - 1; j++)
                    {
                        wSheet.Cells[i + 4, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                    }
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

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                int index = dataGridView1.CurrentRow.Index;
                if(dataGridView1.Rows[index].Cells[8].Value.ToString() == "已执行")
                {
                    DialogResult result = MessageBox.Show("确定是否删除该行?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (result == DialogResult.OK)
                    {
                        string SqlQuery = @"DELETE FROM [ChengyiYuntech].[dbo].[ProduceOrder] WHERE [ID] = '" + dataGridView1.Rows[index].Cells[10].Value + "'";
                        string SqlQuery2 = @"DELETE FROM [ChengyiYuntech].[dbo].[ScanRecord] WHERE [POID] = '" + dataGridView1.Rows[index].Cells[10].Value + "'";

                        sql.sqlCmd(SqlQuery);
                        sql.sqlCmd(SqlQuery2);
                        MessageBox.Show("该行已成功移除");
                        Load_Data();
                    }
                }
                else
                {
                    string SqlQuery = @"DELETE FROM [ChengyiYuntech].[dbo].[ProduceOrder] WHERE [ID] = '" + dataGridView1.Rows[index].Cells[10].Value + "'";
                    string SqlQuery2 = @"DELETE FROM [ChengyiYuntech].[dbo].[ScanRecord] WHERE [POID] = '" + dataGridView1.Rows[index].Cells[10].Value + "'";

                    sql.sqlCmd(SqlQuery);
                    sql.sqlCmd(SqlQuery2);
                    MessageBox.Show("该行已成功移除");
                    Load_Data();
                }
            }
            else
            {
                MessageBox.Show("请选择要删除的行");
            }
        }

        private void SeekData(string comboValue)
        {
            DataTable dt = new DataTable();
            dt = sql.getQuery(@"SELECT a.*, b.Mname FROM [ChengyiYuntech].[dbo].[ProduceOrder] a, [dbo].[Machine] b WHERE a.[OMachineCode] = b.[Mcode] AND a.[OID] like '%" + textBox1.Text + "%' " +
                                                    "AND a.[OMachineCode] like '%" + textBox2.Text + "%' AND a.[ODate] BETWEEN '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' AND '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "'" +
                                                    "AND a.[OOrder]  like '" + comboValue + "'");
            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = Convert.ToDateTime(item["ODate"]).ToString("yyyy/MM/dd");
                dataGridView1.Rows[n].Cells[1].Value = item["OOrder"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = item["OMachineCode"].ToString();
                dataGridView1.Rows[n].Cells[3].Value = item["Mname"].ToString();
                dataGridView1.Rows[n].Cells[4].Value = item["OID"].ToString();
                dataGridView1.Rows[n].Cells[5].Value = item["Ohour"].ToString();
                dataGridView1.Rows[n].Cells[6].Value = item["OPName"].ToString();
                dataGridView1.Rows[n].Cells[10].Value = item["ID"].ToString();
                if (item["OStatus"].ToString() == "0")
                {
                    dataGridView1.Rows[n].Cells[8].Value = "未执行";
                }
                else if (item["OStatus"].ToString() == "1")
                {
                    dataGridView1.Rows[n].Cells[8].Value = "已执行";
                }

                DataTable dtStaff = new DataTable();
                dtStaff = sql.getQuery("SELECT [SID] ,[SName] FROM [ChengyiYuntech].[dbo].[Staff] WHERE SID = '" + item["OStaff"].ToString() + "'");

                dataGridView1.Rows[n].Cells[7].Value = dtStaff.Rows[0][1].ToString();

                DataTable dtWFlow = new DataTable();
                dtWFlow = sql.getQuery(@"SELECT WName FROM [ChengyiYuntech].[dbo].[WorkFlow] WHERE ID = '" + item["OWFlow"].ToString() + "'");

                dataGridView1.Rows[n].Cells[9].Value = dtWFlow.Rows[0][0].ToString();
            }

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("查无资料");
            }
        }

        private void format2()
        {
            DateTime dt = dateTimePicker1.Value;//當天，日期若要改為搜尋條件請改這邊
            string dt1 = string.Format("{0:M}", dt);
            string dt2 = string.Format("{0:M}", dt.AddDays(1));
            string dt3 = string.Format("{0:M}", dt.AddDays(2));
            string dt1_2 = string.Format("{0:yyyyMMdd}", dt);
            string dt2_2 = string.Format("{0:yyyyMMdd}", dt.AddDays(1));
            string dt3_2 = string.Format("{0:yyyyMMdd}", dt.AddDays(2));

            string showTable = "select CASE WHEN RN > 1 THEN '' ELSE mname END AS 機台,isnull(aoid,'') as N'" + dt1 + "白班',isnull(convert(varchar,ahour),'')as 工時," +
                 "isnull(convert(varchar,ahour*60*mspeed),'') as 數量,case when ahour is null then '' else munit end as 單位,isnull(boid,'') as N'" + dt1 + "晚班'," +
                 "isnull(convert(varchar,bhour),'')as 工時,isnull(convert(varchar,bhour*60*mspeed),'') as 數量,case when bhour  is null then '' else munit end as 單位," +
                 "isnull(coid,'') as N'" + dt2 + "白班',isnull(convert(varchar,chour),'')as 工時,isnull(convert(varchar,chour*60*mspeed),'') as 數量," +
                 "case when chour  is null then '' else munit end as 單位,isnull(doid,'') as N'" + dt2 + "晚班',isnull(convert(varchar,dhour),'')as 工時," +
                 "isnull(convert(varchar,dhour*60*mspeed),'') as 數量,case when dhour  is null then '' else munit end as 單位,isnull(eoid,'') as N'" + dt3 + "白班'," +
                 "isnull(convert(varchar,ehour),'')as 工時,isnull(convert(varchar,ehour*60*mspeed),'') as 數量," +
                 "case when ehour  is null then '' else munit end as 單位,isnull(foid,'') as  N'" + dt3 + "晚班',isnull(convert(varchar,fhour),'')as 工時," +
                 "isnull(convert(varchar,fhour*60*mspeed),'') as 數量,case when fhour  is null then '' else munit end as 單位 " +
                 "from(select RN = ROW_NUMBER() OVER ( PARTITION BY a.mname order by a.mname),a.mname,a.mcode,a.mspeed,a.minunit as munit,b.*  " +
                 "from [ChengyiYuntech].[dbo].[Machine]  a left join (select *   from(select * from(select * from(select * from (select * from  " +
                 "(select [omachinecode] as amcode,[oid] as aoid,[ohour] as ahour from [ChengyiYuntech].[dbo].[ProduceOrder]  " +
                 "where ODATE = '" + dt1_2 + "' and [oorder]='白班' )a full outer join (select [omachinecode] as bmcode,[oid] as boid,[ohour] as bhour " +
                 "from [ChengyiYuntech].[dbo].[ProduceOrder]  where ODATE = '" + dt1_2 + "'and [oorder]='晚班' )b " +
                 "on a.aoid = b.boid and a.amcode=b.bmcode)v1 full outer join (select [omachinecode] as cmcode,[oid] as coid,[ohour] as chour " +
                 "from [ChengyiYuntech].[dbo].[ProduceOrder]  where ODATE = '" + dt2_2 + "' and [oorder]='白班' )c " +
                 "on (v1.aoid = c.coid or v1.boid = c.coid) and (v1.amcode=c.cmcode or v1.bmcode=c.cmcode))v2 " +
                 "full outer join (select [omachinecode] as dmcode,[oid] as doid,[ohour] as dhour from [ChengyiYuntech].[dbo].[ProduceOrder]  " +
                 "where ODATE = '" + dt2_2 + "'and [oorder]='晚班' )d " +
                 "on (v2.aoid = d.doid or v2.boid = d.doid or v2.coid = d.doid ) " +
                 "and (v2.amcode=d.dmcode or v2.bmcode=d.dmcode or v2.cmcode=d.dmcode))v3 " +
                 "full outer join  (select [omachinecode] as emcode,[oid] as eoid,[ohour] as ehour from [ChengyiYuntech].[dbo].[ProduceOrder]   " +
                 "where ODATE = '" + dt3_2 + "' and [oorder]='白班')e on (v3.aoid = e.eoid or v3.boid = e.eoid or v3.coid = e.eoid or v3.doid = e.eoid ) " +
                 "and (v3.amcode=e.emcode or v3.bmcode=e.emcode or v3.cmcode=e.emcode or v3.dmcode=e.emcode))v4 " +
                 "full outer join  (select [omachinecode] as fmcode,[oid] as foid,[ohour] as fhour from [ChengyiYuntech].[dbo].[ProduceOrder]   " +
                 "where ODATE =  '" + dt3_2 + "' and [oorder]='晚班' )f on (v4.aoid = f.foid or v4.boid = f.foid or v4.coid =f.foid or v4.doid = f.foid or v4.eoid = f.foid) " +
                 "and (v4.amcode=f.fmcode or v4.bmcode=f.fmcode or v4.cmcode=f.fmcode or v4.dmcode=f.fmcode or v4.emcode=f.fmcode ))b " +
                 "on a.mcode = b.amcode or a.mcode = b.bmcode or a.mcode = b.cmcode or a.mcode = b.dmcode or a.mcode = b.emcode or a.mcode = b.fmcode) v " +
                 "order by mcode";

            SqlServer sql = new SqlServer();
            sql.Connect("ChengyiYuntech");
            dsMainTable = sql.SqlCmd(showTable);
            dataGridView2.DataSource = dsMainTable.Tables[0];
            dataGridView2.Columns[0].Frozen = true;

            string mcode = dataGridView2.Rows[0].Cells[0].Value.ToString();
            string mcode2;
            Color[] rowColor = new Color[2];
            rowColor[0] = Color.AliceBlue;
            rowColor[1] = Color.LightYellow;
            bool colorSwitch = false;
            dataGridView2.Rows[0].DefaultCellStyle.BackColor = rowColor[0];
            for (int i = 1; i < dataGridView2.RowCount - 1; i++)
            {
                mcode2 = dataGridView2.Rows[i].Cells[0].Value.ToString();
                if (mcode2 != "")
                {
                    colorSwitch = !colorSwitch;
                }
                dataGridView2.Rows[i].DefaultCellStyle.BackColor = rowColor[Convert.ToInt32(colorSwitch)];
            }
            sql.Close();
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                string execute = Convert.ToString(this.dataGridView1.Rows[e.RowIndex].Cells["Column8"].Value);
                if (execute == "未执行")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
        }

        private void bunifuCheckbox1_OnChange(object sender, EventArgs e)
        {
            try
            {
                if (checkBox.Checked == true)
                {
                    this.label2.Visible = false;
                    this.label3.Visible = false;
                    this.label4.Visible = false;
                    this.label5.Visible = false;
                    this.textBox1.Visible = false;
                    this.textBox2.Visible = false;
                    this.dateTimePicker2.Visible = false;
                    this.comboBox1.Visible = false;
                    this.label5.Visible = false;
                    this.dataGridView2.Visible = true;
                    this.dataGridView1.Visible = false;
                    btnExport.Visible = false;

                    format2();
                }
                else
                {
                    this.label2.Visible = true;
                    this.label3.Visible = true;
                    this.label4.Visible = true;
                    this.label5.Visible = true;
                    this.textBox1.Visible = true;
                    this.textBox2.Visible = true;
                    this.dateTimePicker2.Visible = true;
                    this.comboBox1.Visible = true;
                    this.label5.Visible = true;
                    this.dataGridView1.Visible = true;
                    this.dataGridView2.Visible = false;
                    btnExport.Visible = true;
                }
            }
            catch (Exception)
            {

            } 
        }

        private void toolTip1_Draw(object sender, DrawToolTipEventArgs e)
        {
            e.Graphics.FillRectangle(Brushes.AliceBlue, e.Bounds);
            e.Graphics.DrawRectangle(Pens.Chocolate, new System.Drawing.Rectangle(0, 0, e.Bounds.Width - 1, e.Bounds.Height - 1));
            e.Graphics.DrawString(this.toolTip1.ToolTipTitle + e.ToolTipText, e.Font, Brushes.Red, e.Bounds);
        }

        private void dataGridView2_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            System.Drawing.Point mousePos = PointToClient(MousePosition);
            SqlServer sql = new SqlServer();
            sql.Connect("ChengyiYuntech");
            int i = e.ColumnIndex;
            int ei = e.RowIndex;


            if (i < 0 || ei < 0)
            {
                return;
            }
            else
            {
                try
                {
                    string billNo = dataGridView2.Rows[ei].Cells[i].Value.ToString();
                    string productDetailCmd = "select opname from  [ChengyiYuntech].[dbo].[ProduceOrder] where oid = '" + billNo + "'";
                    dsProductDetail = sql.SqlCmd(productDetailCmd);
                    string tip;
                    tip = dsProductDetail.Tables[0].Rows[0]["opname"].ToString();
                    this.toolTip1.Hide(this.dataGridView2);
                    this.toolTip1.Show(tip, this.dataGridView2, new System.Drawing.Point(mousePos.X, mousePos.Y - 200));
                }
                catch
                {
                    toolTip1.Hide(dataGridView2);
                }
            }
        }

        private void dataGridView2_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            this.toolTip1.Hide(this.dataGridView2);
        }

        private void dataGridView1_CellPainting_1(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                string execute = Convert.ToString(this.dataGridView1.Rows[e.RowIndex].Cells["Column8"].Value);
                if (execute == "未执行")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count != 0)
            {
                int index;
                index = dataGridView1.CurrentRow.Index;
                Order_Edit order_Edit = new Order_Edit(dataGridView1, index);
                order_Edit.Show();
            }
            else
            {
                MessageBox.Show("请选择编辑行");
            }
        }
    }
}
