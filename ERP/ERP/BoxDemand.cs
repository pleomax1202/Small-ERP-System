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
    public partial class BoxDemand : Form
    {
        Sql sql = new Sql();
        string txtValue, query;

        public BoxDemand()
        {
            InitializeComponent();
        }

        private void BoxDemand_Load(object sender, EventArgs e)
        {
            cbbEnough.SelectedIndex = 0;
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            dataGridView1.Rows.Clear();

            if (cbbEnough.SelectedIndex == 0)
            {
                txtValue = "";
            }
            else
            {
                txtValue = " where a.是否足夠<0 ";
            }

            query = @"with a as(select a.FBillNo,b.FItemID as 產品內碼,b.Fnumber,b.Fmodel,a.FAuxQty,d.FitemID as 外箱編碼,c.FChildNumber,c.FChildModel from  
                    [" + sql.CYDB + "].[dbo].[ICMO] a,[" + sql.CYDB + "].[dbo].[T_ICITem] b,[" + sql.CYDB + "].[dbo].[vICBOM] c,[" + sql.CYDB + "].[dbo].[T_ICITem] d " +
                    "where a.FitemID = b.FItemID and a.Fstatus = '1' and b.Fnumber like '14%' and b.FNumber = c.Fnumber and c.FchildNumber = d.FNUmber " +
                    "and c.Fchildmodel like '%外箱%' and c.FuseStatus = '使用'), " +
                    "a2 as (select a.FBillNo,b.FItemID as 產品內碼,b.Fnumber,b.Fmodel,a.FAuxQty,d.FitemID as 外箱編碼,c.FChildNumber,c.FChildModel from " +
                    "[" + sql.CKDB + "].[dbo].[ICMO] a,[" + sql.CKDB + "].[dbo].[T_ICITem] b,[" + sql.CKDB + "].[dbo].[vICBOM] c,[" + sql.CKDB + "].[dbo].[T_ICITem] " +
                    "d " +
                    "where a.FitemID = b.FItemID and a.Fstatus = '1' and b.Fnumber like '14%' and b.FNumber = c.Fnumber and c.FchildNumber = d.FNUmber " +
                    "and c.Fchildmodel like '%外箱%' and c.FuseStatus = '使用'), " +
                    "/*2外箱庫存量*/ " +
                    "b as( " +
                    "select b.Fmodel,b.FitemID,b.FNumber,sum(a.FQty) as 庫存 from " +
                    "[" + sql.CYDB + "].[dbo].[ICinventory] a,[" + sql.CYDB + "].[dbo].[T_ICITem] " +
                    "b " +
                    "where a.FStockID ='809' and a.FitemID = b.FitemID and b.Fmodel like '%外箱%' group by a.FitemID, b.Fmodel, b.FNumber, b.FitemID ), " +
                    "b2 as( " +
                    "select b.Fmodel,b.FitemID,b.FNumber,sum(a.FQty) as 庫存 from " +
                    "[" + sql.CKDB + "].[dbo].[ICinventory] a,[" + sql.CKDB + "].[dbo].[T_ICITem] " +
                    "b " +
                    "where a.FStockID ='19761' and a.FitemID = b.FitemID and b.Fmodel like '%外箱%' group by a.FitemID, b.Fmodel, b.FNumber, b.FitemID ), " +
                    "/*3領料數與需求數*/ " +
                    "c as(select a.Fbillno,a.Fnumber,a.Fmodel,a.FitemID,a.外箱編碼,a.FchildModel,a.FAuxQty as 總需求數,SUM(isnull(b.FBaseQty,0)) as 現階段領料數 from " +
                    "(select a.Fbillno, b.Fnumber, b.Fmodel, b.FitemID, d.FitemID as 外箱編碼, a.FAuxQty, c.FchildModel from " +
                    "[" + sql.CYDB + "].[dbo].[ICMO] a, " +
                    "[" + sql.CYDB + "].[dbo].[T_ICITem] b, " +
                    "[" + sql.CYDB + "].[dbo].[vICBOM] c, " +
                    "[" + sql.CYDB + "].[dbo].[T_ICITem] d " +
                    "where a.FitemID = b.FitemID and a.Fstatus = '1' and b.Fnumber like '14%' and c.Fchildmodel like '%外箱%' and c.FuseStatus = '使用' and b.FNumber = c.Fnumber " +
                    "and c.FchildNumber = d.FNUmber) a left join " +
                    "(select d.Fuse, e.FitemID, d.FBaseQty from [" + sql.CYDB + "].[dbo].[ICStockbillentry] c, " +
                    "[" + sql.CYDB + "].[dbo].[vwICBill_11] d, " +
                    "[" + sql.CYDB + "].[dbo].[T_ICItem] e " +
                    "where c.FinterID = d.FinterID and c.FentryID = d.FEntryID  and c.FitemID = e.FitemID and e.Fmodel like '%外箱%') b on b.Fuse = a.Fbillno and  b.FitemID = a.外箱編碼 " +
                    "group by a.FAuxQty, a.Fnumber, a.Fmodel, a.FitemID, a.外箱編碼, a.FchildModel, a.Fbillno), " +
                    "c2 as(select a.Fbillno,a.Fnumber,a.Fmodel,a.FitemID,a.外箱編碼,a.FchildModel,a.FAuxQty as 總需求數,SUM(isnull(b.FBaseQty,0)) as 現階段領料數 from " +
                    "(select a.Fbillno, b.Fnumber, b.Fmodel, b.FitemID, d.FitemID as 外箱編碼, a.FAuxQty, c.FchildModel from " +
                    "[" + sql.CKDB + "].[dbo].[ICMO] a, " +
                    "[" + sql.CKDB + "].[dbo].[T_ICITem] b, " +
                    "[" + sql.CKDB + "].[dbo].[vICBOM] c, " +
                    "[" + sql.CKDB + "].[dbo].[T_ICITem] d " +
                    "where a.FitemID = b.FitemID and a.Fstatus = '1' and b.Fnumber like '14%' and c.Fchildmodel like '%外箱%' and c.FuseStatus = '使用' and b.FNumber = c.Fnumber " +
                    "and c.FchildNumber = d.FNUmber) a left join " +
                    "(select d.Fuse, e.FitemID, d.FBaseQty from [" + sql.CKDB + "].[dbo].[ICStockbillentry] c, " +
                    "[" + sql.CKDB + "].[dbo].[vwICBill_11] d, " +
                    "[" + sql.CKDB + "].[dbo].[T_ICItem] e " +
                    "where c.FinterID = d.FinterID and c.FentryID = d.FEntryID  and c.FitemID = e.FitemID and e.Fmodel like '%外箱%') b on b.Fuse = a.Fbillno and  b.FitemID = a.外箱編碼 " +
                    "group by a.FAuxQty, a.Fnumber, a.Fmodel, a.FitemID, a.外箱編碼, a.FchildModel, a.Fbillno), " +
                    "/*4入庫數量*/ " +
                    "d as ( " +
                    "select a.FbatchNo,d.FItemID,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName " +
                    "from[" + sql.CYDB + "].[dbo].[ICStockbillentry] a,[" + sql.CYDB + "].[dbo].[vwICBill_2] b,[" + sql.CYDB + "].[dbo].[t_ICItem] " +
                    "d " +
                    "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo<> ''  " +
                    "group by a.FbatchNo,d.FItemID,b.FCUUnitName " +
                    "), " +
                    "d2 as ( " +
                    "select a.FbatchNo,d.FItemID,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName " +
                    "from[" + sql.CKDB + "].[dbo].[ICStockbillentry] a,[" + sql.CKDB + "].[dbo].[vwICBill_2] b,[" + sql.CKDB + "].[dbo].[t_ICItem] " +
                    "d " +
                    "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo<> ''  " +
                    "group by a.FbatchNo,d.FItemID,b.FCUUnitName " +
                    ") " +
                    ",e as( " +
                    "select c.外箱編碼 as FitemID,isnull(b.庫存,0)-isnull(c.欠料,0) as 余料 from( " +
                    "select c.外箱編碼,case when SUM(isnull(c.總需求數,0))>SUM(isnull(c.現階段領料數,0)) then SUM(isnull(c.總需求數,0))-SUM(isnull(c.現階段領料數,0)) else '0' end as 欠料 " +
                    "from c group by c.外箱編碼 )c left join b on c.外箱編碼 = b.FitemID ) " +
                    ",e2 as( " +
                    "select c2.外箱編碼 as FitemID,isnull(b2.庫存,0)-isnull(c2.欠料,0) as 余料 from( " +
                    "select c2.外箱編碼,case when SUM(isnull(c2.總需求數,0))>SUM(isnull(c2.現階段領料數,0)) then SUM(isnull(c2.總需求數,0))-SUM(isnull(c2.現階段領料數,0)) else '0' end as 欠料 " +
                    "from c2 group by c2.外箱編碼 )c2 left join b2 on c2.外箱編碼 = b2.FitemID ) " +
                    "select a.*,isnull(在途, 0) as 在途 from " +
                    "((select c.Fbillno, c.Fmodel, c.FchildModel, c.總需求數, isnull(d.Qty, 0) as 已入庫量, " +
                    "case when isnull(c.總需求數, 0) > isnull(d.Qty, 0) then " +
                    "isnull(c.總需求數, 0) - isnull(d.Qty, 0) else '0' end as 未入庫量, c.現階段領料數, " +
                    "isnull(c.現階段領料數, 0) - isnull(d.Qty, 0) as 現場紙箱數量, isnull(b.庫存, 0) as 庫存, isnull(e.余料, 0) as 是否足夠 " +
                    "from c left join b on c.外箱編碼 = b.FitemID left " +
                    "join d on c.Fbillno = d.Fbatchno left " +
                    "join e on c.外箱編碼 = e.FitemID " +
                    "where c.Fnumber like '%"+ txtBillCode.Text +"%' and c.Fmodel like '%"+ txtPName.Text +"%')union " +
                    "(select c2.Fbillno, c2.Fmodel, c2.FchildModel, c2.總需求數, isnull(d2.Qty, 0) as 已入庫量, " +
                    "case when isnull(c2.總需求數, 0) > isnull(d2.Qty, 0) then " +
                    "isnull(c2.總需求數, 0) - isnull(d2.Qty, 0) else '0' end as 未入庫量, c2.現階段領料數, " +
                    "isnull(c2.現階段領料數, 0) - isnull(d2.Qty, 0) as 現場紙箱數量, isnull(b2.庫存, 0) as 庫存, isnull(e2.余料, 0) as 是否足夠 " +
                    "from c2 left join b2 on c2.外箱編碼 = b2.FitemID left " +
                    "join d2 on c2.Fbillno = d2.Fbatchno left " +
                    "join e2 on c2.外箱編碼 = e2.FitemID " +
                    "where c2.Fnumber like '%%' and c2.Fmodel like '%%')) a left join " +
                    "((select b.FNumber, b.Fmodel, SUM(a.FcommitQty) as 在途, SUM(a.Fqty) as 採購量 from " +
                    "[" + sql.CYDB + "].[dbo].[POOrderEntry] a, " +
                    "[" + sql.CYDB + "].[dbo].[T_ICItem] b " +
                    "where a.FitemID = b.FItemID and a.FcommitQty > 0 " +
                    "and b.Fmodel like '%箱%' group by b.FNumber, b.Fmodel) union " +
                    "(select b.FNumber, b.Fmodel, SUM(a.FcommitQty) as 在途, SUM(a.Fqty) as 採購量 from " +
                    "[" + sql.CKDB + "].[dbo].[POOrderEntry] a, " +
                    "[" + sql.CKDB + "].[dbo].[T_ICItem] b " +
                    "where a.FitemID = b.FItemID and a.FcommitQty > 0 " +
                    "and b.Fmodel like '%箱%' group by b.FNumber, b.Fmodel)) b on a.FchildModel = b.Fmodel " + txtValue + " and a.Fbillno like '%"+ txtOID.Text +"%' order by a.Fmodel";

            Load_Data(query);

            Cursor = Cursors.Default;

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("查无信息");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            Export_Data();
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
                string exportPath = path + @"\纸箱库存检核表导出";
                string filePath = inputPath + @"\纸箱库存检核表";

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

                wSheet.Name = "纸箱库存检核表";

                // storing Each row and column value to excel sheet
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        wSheet.Cells[i + 4, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
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

        private void Load_Data(string query)
        {
            DataTable dt = new DataTable();
            dt = sql.getQuery(query);

            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = item["FbillNo"].ToString();
                dataGridView1.Rows[n].Cells[1].Value = item["Fmodel"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = Convert.ToDecimal(item["總需求數"]).ToString("N0");
                dataGridView1.Rows[n].Cells[3].Value = Convert.ToDecimal(item["已入庫量"]).ToString("N0");
                dataGridView1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["未入庫量"]).ToString("N0");
                dataGridView1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["現階段領料數"]).ToString("N0");
                dataGridView1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["現場紙箱數量"]).ToString("N0");
                dataGridView1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["庫存"]).ToString("N0");
                dataGridView1.Rows[n].Cells[8].Value = Convert.ToDecimal(item["是否足夠"]).ToString("N0");
                dataGridView1.Rows[n].Cells[9].Value = Convert.ToDecimal(item["在途"]).ToString("N0");
            }
        }
    }
}