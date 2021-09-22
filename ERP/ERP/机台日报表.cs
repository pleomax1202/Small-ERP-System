using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
namespace Combination
{
    public partial class 机台日报表 : Form
    {
        public 机台日报表()
        {
            InitializeComponent();
        }

        DataTable dtPO = new DataTable();
        DataTable dt = new DataTable();
        DataTable dtMachine = new DataTable();
        Sql sql = new Sql();
        int temp;

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                dataGridView1.Rows.Clear();
                string query = @"select a.ID,a.Mcode,a.MName,a.OOrder,a.OID,a.OPName,a.Ohour,a.PHour,a.POPcs,a.PPPcs,a.報廢kg,a.損耗率,case when a.預計產量 = 0 then 0 else a.PPPcs/a.預計產量 end as 績效,case when a.應有產量 = 0 then 0 else a.PPPcs/a.應有產量 end as 速度達成率,a.PNote1,a.PNote2 from
                                                        ((select a.ID,a.MCode,a.Mname,a.OOrder,a.OID,a.OPName,a.OHour,b.PHour,b.POPcs,b.PPPcs,case when a.FNumber like '12.C%' then (b.POPcs-b.PPPcs)*(a.F_122)/1000 else (b.POPcs-b.PPPcs)*(a.F_123+a.F_122)/1000 end as 報廢kg,
                                                        case when b.POPcs = b.PPPcs then '0' else (b.POPcs- b.PPPcs)/b.POPcs end  as 損耗率,b.PNote1,b.PNote2,
                                                        case when a.F_102 = '0' or a.F_108 = '0' or a.F_110 = '0' then '0' else (case when a.MUnit = 'KG' then (case when a.Fnumber  like '12.C%' then ((b.Phour*a.MSpeed*60*1000)/(a.F_122)) else ((b.Phour*a.MSpeed*60*1000)/(a.F_122+a.F_123)) end)
                                                        when a.MUnit = '张'  then (b.Phour*a.MSpeed*60*a.F_110) when a.MUnit = '箱'  then (b.Phour*a.MSpeed*60*a.F_102) when a.MUnit = '米'  then (b.Phour*(a.Mspeed*60*1000/a.F_108)*a.F_110) else (b.Phour*a.MSpeed*60) end) end as 應有產量,
                                                        case when a.F_102 = '0' or a.F_108 = '0' or a.F_110 = '0' then '0' else (case when a.MUnit = 'KG' then  (case when a.Fnumber  like '12.C%' then ((b.Phour*a.MSpeed*60*1000)/(a.F_122)) else ((b.Phour*a.MSpeed*60*1000)/(a.F_122+a.F_123)) end)
                                                        when a.MUnit = '张'  then (a.Ohour*a.MSpeed*60*a.F_110) when a.MUnit = '箱'  then (a.Ohour*a.MSpeed*60*a.F_102) when a.MUnit = '米'  then (a.Ohour*(a.Mspeed*60*1000/a.F_108)*a.F_110) else (a.Ohour*a.MSpeed*60) end) end as 預計產量 from 
                                                        (select e.MCode,a.ODate,e.MName,a.OOrder,a.OID,a.OPName,a.OHour,a.ID,c.F_123,c.F_122,c.FNumber,e.MUnit,e.MSpeed,c.F_102,c.F_108,c.F_110 from [ChengyiYuntech].[dbo].[ProduceOrder]	a
                                                        ,["+ sql.CYDB +"].[dbo].[T_Icitem] c,["+ sql.CYDB +"].[dbo].[ICMO] d,[ChengyiYuntech].[dbo].[Machine] e " +
                                                        "where a.OStatus  = '1' and a.OSample = '0' and a.OID = d.Fbillno and d.FitemID = c.FitemID and a.OMachineCode = e.MCode and CONVERT(varchar,a.Odate,23) = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "') a left join " +
                                                        "(select POID, PHour, POPcs, PPPcs, PNote1, PNote2 from[ChengyiYuntech].[dbo].[ScanRecord]) b on a.ID = b.POID) union " +
                                                        "(select a.ID, a.MCode, a.Mname, a.OOrder, a.OID, a.OPName, a.OHour, b.PHour, b.POPcs, b.PPPcs,case when a.FNumber like '12.C%' then(b.POPcs - b.PPPcs) * (a.F_122) / 1000 else (b.POPcs - b.PPPcs) * (a.F_123 + a.F_122) / 1000 end as 報廢kg, " +
                                                        "case when b.POPcs = b.PPPcs then '0' else (b.POPcs - b.PPPcs) / b.POPcs end as 損耗率, b.PNote1, b.PNote2, " +
                                                        "case when a.F_102 = '0' or a.F_108 = '0' or a.F_110 = '0' then '0' else (case when a.MUnit = 'KG' then(case when a.Fnumber  like '12.C%' then((b.Phour * a.MSpeed * 60 * 1000) / (a.F_122)) else ((b.Phour * a.MSpeed * 60 * 1000) / (a.F_122 + a.F_123)) end) " +
                                                        "when a.MUnit = '张'  then(b.Phour * a.MSpeed * 60 * a.F_110) when a.MUnit = '箱'  then(b.Phour * a.MSpeed * 60 * a.F_102) when a.MUnit = '米'  then(b.Phour * (a.Mspeed * 60 * 1000 / a.F_108) * a.F_110) else (b.Phour * a.MSpeed * 60) end) end as 應有產量, " +
                                                        "case when a.F_102 = '0' or a.F_108 = '0' or a.F_110 = '0' then '0' else (case when a.MUnit = 'KG' then(case when a.Fnumber  like '12.C%' then((b.Phour * a.MSpeed * 60 * 1000) / (a.F_122)) else ((b.Phour * a.MSpeed * 60 * 1000) / (a.F_122 + a.F_123)) end) " +
                                                        "when a.MUnit = '张'  then(a.Ohour * a.MSpeed * 60 * a.F_110) when a.MUnit = '箱'  then(a.Ohour * a.MSpeed * 60 * a.F_102) when a.MUnit = '米'  then(a.Ohour * (a.Mspeed * 60 * 1000 / a.F_108) * a.F_110) else (a.Ohour * a.MSpeed * 60) end) end as 預計產量 from " +
                                                        "(select e.MCode, a.ODate, e.MName, a.OOrder, a.OID, a.OPName, a.OHour, a.ID, c.F_123, c.F_122, c.FNumber, e.MUnit, e.MSpeed, c.F_102, c.F_108, c.F_110 from[ChengyiYuntech].[dbo].[ProduceOrder]  a " +
                                                        ",["+ sql.CKDB +"].[dbo].[T_Icitem] c,["+ sql.CKDB +"].[dbo].[ICMO] d,[ChengyiYuntech].[dbo].[Machine] e " +
                                                        "where a.OStatus = '1' and a.OSample = '0' and a.OID = d.Fbillno and d.FitemID = c.FitemID and a.OMachineCode = e.MCode and CONVERT(varchar, a.Odate, 23) = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "') a left join " +
                                                        "(select POID, PHour, POPcs, PPPcs, PNote1, PNote2 from[ChengyiYuntech].[dbo].[ScanRecord]) b on a.ID = b.POID) union " +
                                                        "(select a.ID, e.MCode, e.Mname, a.OOrder, a.OID, a.OPName, a.OHour, '' as PHour, '' as POPcs, '' as PPPcs, '' as 報廢kg, '' as 損耗率, '' as PNote1, '' as PNote2, '' as 應有產量, '' as 預計產量 " +
                                                        "from[ChengyiYuntech].[dbo].[ProduceOrder]  a,[ChengyiYuntech].[dbo].[Machine] e where a.OStatus = '0' and a.OSample = '0' and a.OMachineCode = e.MCode and CONVERT(varchar, a.Odate, 23) = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "') union " +
                                                        "(select '' as ID, 'ZZ' as Mcode, '' as MName, '' as OOrder, '' as OID, '小計' as OPName, isnull(Sum(a.OHour), 0) as Ohour, isnull(SUM(a.PHour), 0) as PHour, isnull(SUM(a.POPcs), 0) as POPcs, isNull(SUM(a.PPPcs), 0) as PPPcs, isnull(SUM(a.報廢kg), 0) as 報廢kg, '' as 損耗率, '' as PNote1, '' as Pnote2, '' as 應有產量, '' as 預計產量  from " +
                                                        "((select a.MCode, a.Mname, a.OOrder, a.OID, a.OPName, a.OHour, b.PHour, b.POPcs, b.PPPcs,case when a.FNumber like '12.C%' then(b.POPcs - b.PPPcs) * (a.F_122) / 1000 else (b.POPcs - b.PPPcs) * (a.F_123 + a.F_122) / 1000 end as 報廢kg, " +
                                                        "case when b.POPcs = b.PPPcs then '0' else (b.POPcs - b.PPPcs) / b.POPcs end as 損耗率, b.PNote1, b.PNote2 from " +
                                                        "(select e.MCode, a.ODate, e.MName, a.OOrder, a.OID, a.OPName, a.OHour, a.ID, c.F_123, c.F_122, c.FNumber from[ChengyiYuntech].[dbo].[ProduceOrder] a " +
                                                        ",["+ sql.CYDB +"].[dbo].[T_Icitem] c,["+ sql.CYDB +"].[dbo].[ICMO] d,[ChengyiYuntech].[dbo].[Machine] e " +
                                                        "where a.OStatus = '1' and a.OSample = '0' and a.OID = d.Fbillno and d.FitemID = c.FitemID and a.OMachineCode = e.MCode and CONVERT(varchar, a.Odate, 23) = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "') a left join " +
                                                        "(select POID, PHour, POPcs, PPPcs, PNote1, PNote2 from[ChengyiYuntech].[dbo].[ScanRecord]) b on a.ID = b.POID) union " +
                                                        "(select a.MCode, a.Mname, a.OOrder, a.OID, a.OPName, a.OHour, b.PHour, b.POPcs, b.PPPcs,case when a.FNumber like '12.C%' then(b.POPcs - b.PPPcs) * (a.F_122) / 1000 else (b.POPcs - b.PPPcs) * (a.F_123 + a.F_122) / 1000 end as 報廢kg, " +
                                                        "case when b.POPcs = b.PPPcs then '0' else (b.POPcs - b.PPPcs) / b.POPcs end as 損耗率, b.PNote1, b.PNote2 from " +
                                                        "(select e.MCode, a.ODate, e.MName, a.OOrder, a.OID, a.OPName, a.OHour, a.ID, c.F_123, c.F_122, c.FNumber from[ChengyiYuntech].[dbo].[ProduceOrder] a " +
                                                        ",["+ sql.CKDB +"].[dbo].[T_Icitem] c,["+ sql.CKDB +"].[dbo].[ICMO] d,[ChengyiYuntech].[dbo].[Machine] e " +
                                                        "where a.OStatus = '1' and a.OSample = '0' and a.OID = d.Fbillno and d.FitemID = c.FitemID and a.OMachineCode = e.MCode and CONVERT(varchar, a.Odate, 23) = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "') a left join " +
                                                        "(select POID, PHour, POPcs, PPPcs, PNote1, PNote2 from[ChengyiYuntech].[dbo].[ScanRecord]) b on a.ID = b.POID) union " +
                                                        "(select e.MCode, e.Mname, a.OOrder, a.OID, a.OPName, a.OHour, '' as PHour, '' as POPcs, '' as PPPcs, '' as 報廢kg, '' as 損耗率, '' as PNote1, '' as PNote2 " +
                                                        "from[ChengyiYuntech].[dbo].[ProduceOrder]  a,[ChengyiYuntech].[dbo].[Machine] e where a.OStatus = '0' and a.OSample = '0' and a.OMachineCode = e.MCode and CONVERT(varchar, a.Odate, 23) = '" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "')) a))a order by a.Mcode,a.OOrder,a.ID";
                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = dateTimePicker1.Value.ToString("yyyy/MM/dd");
                    dataGridView1.Rows[n].Cells[1].Value = item["Mcode"].ToString();
                    dataGridView1.Rows[n].Cells[2].Value = item["Mname"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["OOrder"].ToString();
                    dataGridView1.Rows[n].Cells[4].Value = item["OID"].ToString();
                    dataGridView1.Rows[n].Cells[5].Value = item["OPName"].ToString();
                    dataGridView1.Rows[n].Cells[6].Value = item["Ohour"].ToString();
                    dataGridView1.Rows[n].Cells[7].Value = item["PHour"].ToString();
                    dataGridView1.Rows[n].Cells[8].Value = Convert.ToDecimal(item["POPcs"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[9].Value = Convert.ToDecimal(item["PPPcs"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[10].Value = Convert.ToDecimal(item["報廢kg"]).ToString("N2");
                    dataGridView1.Rows[n].Cells[11].Value = Convert.ToDecimal(item["損耗率"]).ToString("p");
                    dataGridView1.Rows[n].Cells[12].Value = Convert.ToDecimal(item["績效"]).ToString("p");
                    dataGridView1.Rows[n].Cells[13].Value = Convert.ToDecimal(item["速度達成率"]).ToString("p");
                    dataGridView1.Rows[n].Cells[14].Value = Convert.ToString(item["PNote1"]);
                    dataGridView1.Rows[n].Cells[15].Value = Convert.ToString(item["PNote2"]);
                    temp = n;
                }

                int count = dataGridView1.Rows.Count;
                dataGridView1.Rows[count - 1].Cells[0].Value = "";
                dataGridView1.Rows[count - 1].Cells[1].Value = "";
                dataGridView1.Rows[count - 1].Cells[11].Value = "";

                dt.Clear();

                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("查无资料");
                }
            }
            catch (System.InvalidCastException)
            {
                dt.Clear();
                dataGridView1.Rows.Clear();
                MessageBox.Show("查无资料");
            }
            Cursor = Cursors.Default;
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
                string exportPath = path + @"\机台日报表导出" + dateTimePicker1.Value.ToString("yyMMdd");
                string filePath = inputPath + @"\机台日报表";

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

                wSheet.Name = "机台日报表";

                wSheet.Cells[2, 1] = "机台日报表     " + dateTimePicker1.Value.ToString("yyyy/MM/dd");
                wRange = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                // storing Each row and column value to excel sheet
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 1; j < dataGridView1.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 4, j] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);

                        if (Convert.ToDecimal(dataGridView1.Rows[i].Cells[7].Value) == 0)
                        {
                            wRange = wSheet.Range[wSheet.Cells[i + 4, 1], wSheet.Cells[i + 4, 15]];
                            wRange.Select();
                            wRange.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                        }
                    }
                }

                Excel.Range last = wSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range allRange = wSheet.get_Range("A1", last);
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
                string execute = Convert.ToString(this.dataGridView1.Rows[e.RowIndex].Cells["Column10"].Value);
                if (Convert.ToDecimal(execute) == 0)
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
        }

        private void btnMSpdA_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
            Process p = new Process();
            string inputPath = System.Environment.CurrentDirectory;
            string fileName = @"\SpeedAnalysis.exe";
            p.StartInfo.FileName = inputPath + fileName;
            p.Start();
            Cursor = Cursors.WaitCursor;
        }

        private void btnWasted_Click(object sender, EventArgs e)
        {
            Process p = new Process();
            string inputPath = System.Environment.CurrentDirectory;
            string fileName = @"\报废原因分析.exe";
            p.StartInfo.FileName = inputPath + fileName;
            p.Start();
        }
    }
}
