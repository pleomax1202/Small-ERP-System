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
using Combination.Detail;
using Excel = Microsoft.Office.Interop.Excel;

namespace Combination
{
    public partial class FactoryInput : Form
    {

        string Sid, sRole;
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;
        string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string inputPath = System.Environment.CurrentDirectory;
        string exportPath;
        string filePath;
        Sql sql = new Sql();

        public FactoryInput(string SID, string role)
        {
            InitializeComponent();
            Sid = SID;
            sRole = role;
        }

        private void FactoryInput2_Load(object sender, EventArgs e)
        {
            this.ActiveControl = textBox1;
            dgvSample.Visible = false;

            if(sRole == "权限管理员" || sRole == "系统管理员" || sRole == "最高管理员")
            {
                this.button3.Visible = true;
                this.button4.Visible = true;
                this.button5.Visible = true;
            }
            else if(sRole == "干部")
            {
                this.button3.Visible = false;
                this.button4.Visible = false;
            }
            else
            {
                this.button3.Visible = false;
                this.button4.Visible = false;
                this.button5.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            if (FscCheck.Checked == false && NFscCheck.Checked == false && SampleCheckbox.Checked == true)
            {
                string query = @"select * from  ((select convert(varchar(100), a.pdate,111) as 日期,upper(b.omachinecode) AS 機台編號,c.mname AS 機台名稱,b.oid as 製造單號,b.opname as 產品名稱,b.oorder as 班次,
                                a.phour as 花費工時, a.poqty as 領料,c.Minunit as 單位1, a.ppqty as 產出,c.Moutunit as 單位2,d.SName as 操作員  from [ChengyiYuntech].[dbo].[ScanRecord]a,[ChengyiYuntech].[dbo].[ProduceOrder]b,[ChengyiYuntech].[dbo].[Machine]c,[ChengyiYuntech].[dbo].[STAFF] d
                                where a.poid = b.id AND c.mcode = b.omachinecode and d.SID = a.PStaff and (b.oid like '%" + textBox1.Text + "' or  b.OPname like '%" + textBox3.Text + "%') and b.odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' " +
                                "and b.osample=1) union ( select convert(varchar(100), b.odate,111) as 日期,upper(b.omachinecode) AS 機台編號,c.mname AS 機台名稱,b.oid as 製造單號,b.opname as 產品名稱, b.oorder as 班次, " +
                                "0 as 花費工時,0 as 領料,c.Minunit as 單位1, 0 as 產出,c.Moutunit as 單位2,'無' as 操作員  " +
                                "from [ChengyiYuntech].[dbo].[ProduceOrder]b,[ChengyiYuntech].[dbo].[Machine]c where b.ostatus =0 AND c.mcode = b.omachinecode and (b.oid like '%" + textBox1.Text + "' or  b.OPname like '%" + textBox3.Text + "%' ) and b.odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' " +
                                "and b.osample=1) ) v1";

                dataGridView1.Columns[0].HeaderText = "日期";
                dataGridView1.Columns[1].HeaderText = "机台名称";
                dataGridView1.Columns[2].HeaderText = "制造单号";
                dataGridView1.Columns[3].HeaderText = "产品名称";
                dataGridView1.Columns[4].HeaderText = "机台编号";
                dataGridView1.Columns[5].HeaderText = "班次";
                dataGridView1.Columns[6].HeaderText = "花费工时";
                dataGridView1.Columns[7].HeaderText = "领料";
                dataGridView1.Columns[8].HeaderText = "单位";
                dataGridView1.Columns[9].HeaderText = "产出";
                dataGridView1.Columns[10].HeaderText = "单位";
                dataGridView1.Columns[11].HeaderText = "操作员";

                for (int i = 0; i < 12; i++)
                {
                    dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

                    if (i == 2)
                    {
                        dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                        continue;
                    }
                }

                for (int i = 12; i < dataGridView1.Columns.Count; i++)
                {
                    dataGridView1.Columns[i].Visible = false;
                }


                Load_Data(query);
            }
            else if (FscCheck.Checked == true && NFscCheck.Checked == true)
            {
                ResizeColumn();

                string query = @"select * from ((select a.POID,b.ODate,c.MCode as 機台編號,c.Mname as 機台名稱,b.OID  as 製造單號,b.OOrder as 班次,b.OPName as 產品名稱,a.PHour as 生產工時,a.POQty as 領料數量,C.MINUnit as 領料單位,(a.POPcs*(e.F_122+e.F_123))/1000 as 領料kg,a.PPQty as 產出數量,c.MoutUNit as 產出單位,(a.PPPcs*(e.F_122+e.F_123))/1000 as 產出kg,a.PWQty as 報廢數量,c.MWUnit as 報廢單位,a.PWWeight as 報廢kg,a.PNote1 工時原因,a.PNote2 as 報廢原因,case when e.FDefaultLoc = '20421' then '是' else '否' end as FSC,d.SName as 操作員 from 
                                                        [ChengyiYuntech].[dbo].[ScanRecord] a ,[ChengyiYuntech].[dbo].[Produceorder] b,[ChengyiYuntech].[dbo].[Machine] c, [ChengyiYuntech].[dbo].[Staff] d,["+ sql.CYDB +"].[dbo].[T_Icitem] e,["+ sql.CYDB +"].[dbo].[ICMO] f " +
                                                        "where a.POID = b.ID and b.OMachineCode = c.Mcode  and d.SID = a.PStaff and b.OID = f.Fbillno and e.FitemID = f.FitemID) union " +
                                                        "(select a.POID,b.ODate,c.MCode as 機台編號,c.Mname as 機台名稱,b.OID  as 製造單號,b.OOrder as 班次,b.OPName as 產品名稱,a.PHour as 生產工時,a.POQty as 領料數量,C.MINUnit as 領料單位,(a.POPcs*(e.F_122+e.F_123))/1000 as 領料kg,a.PPQty as 產出數量,c.MoutUNit as 產出單位,(a.PPPcs*(e.F_122+e.F_123))/1000 as 產出kg,a.PWQty as 報廢數量,c.MWUnit as 報廢單位,a.PWWeight as 報廢kg,a.PNote1 工時原因,a.PNote2 as 報廢原因,case when e.FDefaultLoc = '20421' then '是' else '否' end as FSC,d.SName as 操作員 from  " +
                                                        "[ChengyiYuntech].[dbo].[ScanRecord] a ,[ChengyiYuntech].[dbo].[Produceorder] b,[ChengyiYuntech].[dbo].[Machine] c, [ChengyiYuntech].[dbo].[Staff] d,["+ sql.CKDB +"].[dbo].[T_Icitem] e,["+ sql.CKDB +"].[dbo].[ICMO] f " +
                                                        "where a.POID = b.ID and b.OMachineCode = c.Mcode  and d.SID = a.PStaff and b.OID = f.Fbillno and e.FitemID = f.FitemID)) a  " +
                                                        "where a.製造單號 like '%" + textBox1.Text + "' and a.機台編號 like '%" + textBox3.Text + "%'and a.ODate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and a.產品名稱 like '%" + textBox2.Text + "%' Order by a.機台名稱, a.班次";
                Load_Data(query);
            }
            else if (FscCheck.Checked == true)
            {
                ResizeColumn();

                string query = @"select * from ((select a.POID,b.ODate,c.MCode as 機台編號,c.Mname as 機台名稱,b.OID  as 製造單號,b.OOrder as 班次,b.OPName as 產品名稱,a.PHour as 生產工時,a.POQty as 領料數量,C.MINUnit as 領料單位,(a.POPcs*(e.F_122+e.F_123))/1000 as 領料kg,a.PPQty as 產出數量,c.MoutUNit as 產出單位,(a.PPPcs*(e.F_122+e.F_123))/1000 as 產出kg,a.PWQty as 報廢數量,c.MWUnit as 報廢單位,a.PWWeight as 報廢kg,a.PNote1 工時原因,a.PNote2 as 報廢原因,case when e.FDefaultLoc = '20421' then '是' else '否' end as FSC,d.SName as 操作員 from 
                                                        [ChengyiYuntech].[dbo].[ScanRecord] a ,[ChengyiYuntech].[dbo].[Produceorder] b,[ChengyiYuntech].[dbo].[Machine] c, [ChengyiYuntech].[dbo].[Staff] d,["+ sql.CYDB +"].[dbo].[T_Icitem] e,["+ sql.CYDB +"].[dbo].[ICMO] f " +
                                                        "where a.POID = b.ID and b.OMachineCode = c.Mcode  and d.SID = a.PStaff and b.OID = f.Fbillno and e.FitemID = f.FitemID) union " +
                                                        "(select a.POID,b.ODate,c.MCode as 機台編號,c.Mname as 機台名稱,b.OID  as 製造單號,b.OOrder as 班次,b.OPName as 產品名稱,a.PHour as 生產工時,a.POQty as 領料數量,C.MINUnit as 領料單位,(a.POPcs*(e.F_122+e.F_123))/1000 as 領料kg,a.PPQty as 產出數量,c.MoutUNit as 產出單位,(a.PPPcs*(e.F_122+e.F_123))/1000 as 產出kg,a.PWQty as 報廢數量,c.MWUnit as 報廢單位,a.PWWeight as 報廢kg,a.PNote1 工時原因,a.PNote2 as 報廢原因,case when e.FDefaultLoc = '20421' then '是' else '否' end as FSC,d.SName as 操作員 from  " +
                                                        "[ChengyiYuntech].[dbo].[ScanRecord] a ,[ChengyiYuntech].[dbo].[Produceorder] b,[ChengyiYuntech].[dbo].[Machine] c, [ChengyiYuntech].[dbo].[Staff] d,["+ sql.CKDB +"].[dbo].[T_Icitem] e,["+ sql.CKDB +"].[dbo].[ICMO] f " +
                                                        "where a.POID = b.ID and b.OMachineCode = c.Mcode  and d.SID = a.PStaff and b.OID = f.Fbillno and e.FitemID = f.FitemID)) a  " +
                                                        "where a.製造單號 like '%" + textBox1.Text + "' and a.機台編號 like '%" + textBox3.Text + "%'and a.ODate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and a.FSC = '是' and a.產品名稱 like '%" + textBox2.Text + "%' Order by a.機台名稱, a.班次";
                Load_Data(query);
            }
            else if (NFscCheck.Checked == true)
            {
                ResizeColumn();

                string query = @"select * from ((select a.POID,b.ODate,c.MCode as 機台編號,c.Mname as 機台名稱,b.OID  as 製造單號,b.OOrder as 班次,b.OPName as 產品名稱,a.PHour as 生產工時,a.POQty as 領料數量,C.MINUnit as 領料單位,(a.POPcs*(e.F_122+e.F_123))/1000 as 領料kg,a.PPQty as 產出數量,c.MoutUNit as 產出單位,(a.PPPcs*(e.F_122+e.F_123))/1000 as 產出kg,a.PWQty as 報廢數量,c.MWUnit as 報廢單位,a.PWWeight as 報廢kg,a.PNote1 工時原因,a.PNote2 as 報廢原因,case when e.FDefaultLoc = '20421' then '是' else '否' end as FSC,d.SName as 操作員 from 
                                                        [ChengyiYuntech].[dbo].[ScanRecord] a ,[ChengyiYuntech].[dbo].[Produceorder] b,[ChengyiYuntech].[dbo].[Machine] c, [ChengyiYuntech].[dbo].[Staff] d,["+ sql.CYDB +"].[dbo].[T_Icitem] e,["+ sql.CYDB +"].[dbo].[ICMO] f " +
                                                        "where a.POID = b.ID and b.OMachineCode = c.Mcode  and d.SID = a.PStaff and b.OID = f.Fbillno and e.FitemID = f.FitemID) union " +
                                                        "(select a.POID,b.ODate,c.MCode as 機台編號,c.Mname as 機台名稱,b.OID  as 製造單號,b.OOrder as 班次,b.OPName as 產品名稱,a.PHour as 生產工時,a.POQty as 領料數量,C.MINUnit as 領料單位,(a.POPcs*(e.F_122+e.F_123))/1000 as 領料kg,a.PPQty as 產出數量,c.MoutUNit as 產出單位,(a.PPPcs*(e.F_122+e.F_123))/1000 as 產出kg,a.PWQty as 報廢數量,c.MWUnit as 報廢單位,a.PWWeight as 報廢kg,a.PNote1 工時原因,a.PNote2 as 報廢原因,case when e.FDefaultLoc = '20421' then '是' else '否' end as FSC,d.SName as 操作員 from  " +
                                                        "[ChengyiYuntech].[dbo].[ScanRecord] a ,[ChengyiYuntech].[dbo].[Produceorder] b,[ChengyiYuntech].[dbo].[Machine] c, [ChengyiYuntech].[dbo].[Staff] d,["+ sql.CKDB +"].[dbo].[T_Icitem] e,["+ sql.CKDB +"].[dbo].[ICMO] f " +
                                                        "where a.POID = b.ID and b.OMachineCode = c.Mcode  and d.SID = a.PStaff and b.OID = f.Fbillno and e.FitemID = f.FitemID)) a  " +
                                                        "where a.製造單號 like '%" + textBox1.Text + "' and a.機台編號 like '%" + textBox3.Text + "%'and a.ODate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and a.FSC = '否' and a.產品名稱 like '%" + textBox2.Text + "%' Order by a.機台名稱, a.班次";
                Load_Data(query);
            }
            else if (FscCheck.Checked == false && NFscCheck.Checked == false)
            {
                ResizeColumn();

                string query = @"select * from ((select a.POID,b.ODate,c.MCode as 機台編號,c.Mname as 機台名稱,b.OID  as 製造單號,b.OOrder as 班次,b.OPName as 產品名稱,a.PHour as 生產工時,a.POQty as 領料數量,C.MINUnit as 領料單位,(a.POPcs*(e.F_122+e.F_123))/1000 as 領料kg,a.PPQty as 產出數量,c.MoutUNit as 產出單位,(a.PPPcs*(e.F_122+e.F_123))/1000 as 產出kg,a.PWQty as 報廢數量,c.MWUnit as 報廢單位,a.PWWeight as 報廢kg,a.PNote1 工時原因,a.PNote2 as 報廢原因,case when e.FDefaultLoc = '20421' then '是' else '否' end as FSC,d.SName as 操作員 from 
                                                        [ChengyiYuntech].[dbo].[ScanRecord] a ,[ChengyiYuntech].[dbo].[Produceorder] b,[ChengyiYuntech].[dbo].[Machine] c, [ChengyiYuntech].[dbo].[Staff] d,["+ sql.CYDB +"].[dbo].[T_Icitem] e,["+ sql.CYDB +"].[dbo].[ICMO] f " +
                                                        "where a.POID = b.ID and b.OMachineCode = c.Mcode  and d.SID = a.PStaff and b.OID = f.Fbillno and e.FitemID = f.FitemID) union " +
                                                        "(select a.POID,b.ODate,c.MCode as 機台編號,c.Mname as 機台名稱,b.OID  as 製造單號,b.OOrder as 班次,b.OPName as 產品名稱,a.PHour as 生產工時,a.POQty as 領料數量,C.MINUnit as 領料單位,(a.POPcs*(e.F_122+e.F_123))/1000 as 領料kg,a.PPQty as 產出數量,c.MoutUNit as 產出單位,(a.PPPcs*(e.F_122+e.F_123))/1000 as 產出kg,a.PWQty as 報廢數量,c.MWUnit as 報廢單位,a.PWWeight as 報廢kg,a.PNote1 工時原因,a.PNote2 as 報廢原因,case when e.FDefaultLoc = '20421' then '是' else '否' end as FSC,d.SName as 操作員 from  " +
                                                        "[ChengyiYuntech].[dbo].[ScanRecord] a ,[ChengyiYuntech].[dbo].[Produceorder] b,[ChengyiYuntech].[dbo].[Machine] c, [ChengyiYuntech].[dbo].[Staff] d,["+ sql.CKDB +"].[dbo].[T_Icitem] e,["+ sql.CKDB +"].[dbo].[ICMO] f " +
                                                        "where a.POID = b.ID and b.OMachineCode = c.Mcode  and d.SID = a.PStaff and b.OID = f.Fbillno and e.FitemID = f.FitemID)) a  " +
                                                        "where a.製造單號 like '%" + textBox1.Text + "' and a.機台編號 like '%" + textBox3.Text + "%'and a.ODate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and a.產品名稱 like '%" + textBox2.Text + "%' Order by a.機台名稱, a.班次";
                Load_Data(query);
            }
            Cursor = Cursors.Default;
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

        private void button4_Click(object sender, EventArgs e)
        {

            if (dataGridView1.Rows.Count != 0)
            {
                DialogResult result = MessageBox.Show("清确认是否删除该行", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    int index = dataGridView1.CurrentRow.Index;

                    var SqlQuery = @"DELETE FROM [ChengyiYuntech].[dbo].[ScanRecord] WHERE[POID] = '" + dataGridView1.Rows[index].Cells[20].Value + "'";
                    string SqlQuery2 = @"UPDATE [ChengyiYuntech].[dbo].[ProduceOrder] SET [OStatus] = '0' WHERE ID = '" + dataGridView1.Rows[index].Cells[20].Value + "'";

                    sql.sqlCmd(SqlQuery);
                    sql.sqlCmd(SqlQuery2);
                    MessageBox.Show("该行已成功移除");
                }
            }
            else
            {
                MessageBox.Show("请选择资料行");
            }
        }

        private void Export_Data()
        {
            if (dataGridView1.Rows.Count != 0)
            {
                if (SampleCheckbox.Checked == true)
                {
                    exportPath = path + @"\车间打样输入报表导出";
                    filePath = inputPath + @"\车间打样输入报表";
                }
                else
                {
                    exportPath = path + @"\车间输入报表导出";
                    filePath = inputPath + @"\车间输入报表";
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

                if (SampleCheckbox.Checked == true)
                {
                    wSheet.Name = "车间打样输入报表";

                    wSheet.Cells[3, 1] = dateTimePicker1.Value.ToString("yyyy/MM/dd") + " - " + dateTimePicker2.Value.ToString("yyyy/MM/dd");
                    wRange = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 1]];
                    wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                    // storing Each row and column value to excel sheet
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count - 1; j++)
                        {
                            wSheet.Cells[i + 5, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                        }
                    }
                }
                else
                {
                    wSheet.Name = "车间输入报表";

                    wSheet.Cells[3, 1] = dateTimePicker1.Value.ToString("yyyy/MM/dd") + " - " + dateTimePicker2.Value.ToString("yyyy/MM/dd");
                    wRange = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 1]];
                    wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                    // storing Each row and column value to excel sheet
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count - 1; j++)
                        {
                            wSheet.Cells[i + 5, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                        }
                    }
                }


                Excel.Range last = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range allRange = wSheet.get_Range("A1", last);
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

        private void Load_Data(string query)
        {
            dataGridView1.Rows.Clear();

            DataTable dt = new DataTable();
            dt = sql.getQuery(query);

            if (SampleCheckbox.Checked == true)
            {
                foreach (DataRow item in dt.Rows)
                {
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = Convert.ToString(item["日期"]);
                    dataGridView1.Rows[n].Cells[1].Value = item["機台名稱"].ToString();
                    dataGridView1.Rows[n].Cells[2].Value = item["製造單號"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["產品名稱"].ToString();
                    dataGridView1.Rows[n].Cells[4].Value = item["機台編號"].ToString();
                    dataGridView1.Rows[n].Cells[5].Value = item["班次"].ToString();
                    dataGridView1.Rows[n].Cells[6].Value = item["花費工時"].ToString();
                    dataGridView1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["領料"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[8].Value = item["單位1"].ToString();
                    dataGridView1.Rows[n].Cells[9].Value = Convert.ToDecimal(item["產出"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[10].Value = item["單位2"].ToString();
                    dataGridView1.Rows[n].Cells[11].Value = item["操作員"].ToString();
                }
            }
            else
            {
                foreach (DataRow item in dt.Rows)
                {
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = Convert.ToDateTime(item["ODate"]).ToString("yyyy/MM/dd");
                    dataGridView1.Rows[n].Cells[1].Value = item["機台編號"].ToString();
                    dataGridView1.Rows[n].Cells[2].Value = item["製造單號"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["產品名稱"].ToString();
                    dataGridView1.Rows[n].Cells[4].Value = item["機台名稱"].ToString();
                    dataGridView1.Rows[n].Cells[5].Value = item["班次"].ToString();
                    dataGridView1.Rows[n].Cells[6].Value = item["生產工時"].ToString();
                    dataGridView1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["領料數量"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[8].Value = item["領料單位"].ToString();
                    dataGridView1.Rows[n].Cells[9].Value = Convert.ToDecimal(item["領料kg"]).ToString("N2");
                    dataGridView1.Rows[n].Cells[10].Value = Convert.ToDecimal(item["產出數量"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[11].Value = item["產出單位"].ToString();
                    dataGridView1.Rows[n].Cells[12].Value = Convert.ToDecimal(item["產出kg"]).ToString("N2");
                    dataGridView1.Rows[n].Cells[13].Value = item["報廢數量"].ToString();
                    dataGridView1.Rows[n].Cells[14].Value = item["報廢單位"].ToString();
                    dataGridView1.Rows[n].Cells[15].Value = Convert.ToDecimal(item["報廢kg"]).ToString("N2");
                    dataGridView1.Rows[n].Cells[16].Value = item["報廢原因"].ToString();
                    dataGridView1.Rows[n].Cells[17].Value = item["工時原因"].ToString();
                    dataGridView1.Rows[n].Cells[18].Value = item["FSC"].ToString();
                    dataGridView1.Rows[n].Cells[19].Value = item["操作員"].ToString();
                    dataGridView1.Rows[n].Cells[20].Value = item["POID"].ToString();
                }
            }

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("查无资料");
                this.ActiveControl = textBox1;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FI_Add fI_Add = new FI_Add(int.Parse(Sid));
            fI_Add.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                if (SampleCheckbox.Checked == true)
                {
                    int index = dataGridView1.CurrentRow.Index;
                    FactoryInput_SampleEdit fise = new FactoryInput_SampleEdit(dataGridView1, index, Sid, button1);
                    fise.Show();
                }
                else if (SampleCheckbox.Checked == false)
                {
                    int index = dataGridView1.CurrentRow.Index;
                    machineRPEdit mrp = new machineRPEdit(dataGridView1, index, Sid, button1);
                    mrp.Show();
                }
            }
            else
            {
                MessageBox.Show("请选择资料行");
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = textBox3;
            }
        }

        private void SampleCheckbox1_OnChange(object sender, EventArgs e)
        {
            if (SampleCheckbox.Checked == true)
            {
                FscCheck.Checked = false;
                NFscCheck.Checked = false;
                button5.Visible = false;
                button3.Visible = false;
                button4.Visible = false;
            }
            else
            {
                button5.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
            }
        }

        private void FscCheck_OnChange(object sender, EventArgs e)
        {
            if (FscCheck.Checked == true)
            {
                SampleCheckbox.Checked = false;
                button5.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
            }
        }

        private void NFscCheck_OnChange(object sender, EventArgs e)
        {
            if (NFscCheck.Checked == true)
            {
                SampleCheckbox.Checked = false;
                button5.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
            }
        }

        private void ResizeColumn()
        {
            for (int i = 1; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            }
            for (int i = 11; i < dataGridView1.Columns.Count - 1; i++)
            {
                dataGridView1.Columns[i].Visible = true;
            }

            dataGridView1.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            dataGridView1.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            dataGridView1.Columns[11].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
            dataGridView1.Columns[14].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader;
        }

    }
}
