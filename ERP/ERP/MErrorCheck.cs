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
    public partial class MErrorCheck : Form
    {
        Sql sql = new Sql();
        string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string inputPath = System.Environment.CurrentDirectory;
        string exportPath;
        string filePath;
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;

        public MErrorCheck()
        {
            InitializeComponent();
            comboBox1.SelectedIndex = 0;
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            
            dataGridView1.Rows.Clear();
            Load_Data();
            Cursor = Cursors.Default;

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("查无信息");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if(dataGridView1.Rows.Count != 0)
            {
                Export_Data();
            }
            else
            {
                MessageBox.Show("无信息导出");
            }
        }

        private void Export_Data()
        {
            exportPath = path + @"\串料检核表导出";
            filePath = inputPath + @"\串料检核表";

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

            wSheet.Name = "串料检核表";

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    wSheet.Cells[i + 4, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                }
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 2; j < dataGridView1.ColumnCount - 1; j++)
                {
                    if (Convert.ToDouble(dataGridView1.Rows[i].Cells[j].Value) < 0)
                    {
                        wRange = wSheet.Range[wSheet.Cells[i + 4, j + 1], wSheet.Cells[i + 4, j + 1]];
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

            //wSheet = (Excel.Worksheet)wBook.Worksheets.get_Item(2);
            //wSheet.Select();

            //Save Excel
            wBook.SaveAs(exportPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            wBook = null;
            wSheet = null;
            excelApp = null;

            GC.Collect();
            MessageBox.Show("导出成功");
        }

        private void Load_Data()
        {
            string value;
            if(comboBox1.SelectedIndex == 0)
            {
                value = "( Fstatus =1 or Fstatus=3)";
            }
            else if(comboBox1.SelectedIndex == 1)
            {
                value = "(Fstatus=3)";
            }
            else
            {
                value = "( Fstatus =1)";
            }
            string query = @"with
                            item as 
                            (
                            select f_122,F_123,a.FitemID,Fnumber,b.fbillno,fcheckdate,fmodel, Fstatus from [" + sql.CYDB +"].[dbo].[T_ICItem]a , ["+ sql.CYDB +"].[dbo].[ICMO]b " +
                            "where b.FITEMID =a.FitemID and "+ value +" " +
                            "union all " +
                            "select f_122,F_123,a.FitemID,Fnumber,b.fbillno,fcheckdate,fmodel, Fstatus from [" + sql.CKDB +"].[dbo].[T_ICItem]a, ["+ sql.CKDB +"].[dbo].[ICMO] b " +
                            "where b.FITEMID =a.FitemID and " + value + "  " +
                            "), " +
                            "ph as  " +
                            "( " +
                            "select OWflow,mcode,Mname,convert(varchar(10),odate,120) as pdate,a.oid,a.Opname,popcs,pppcs, " +
                            "case when e.Fnumber like '12.C.02%' then b.popcs*e.F_123 else b.popcs*(e.F_122+e.F_123) end as pokg, " +
                            "case when e.Fnumber like '12.C.02%' then b.pppcs*e.F_122 else b.pppcs*(e.F_122+e.F_123) end as ppkg, " +
                            "case when e.Fnumber like '12.C.02%' then (b.popcs-b.pppcs)*e.F_122/1000 else (b.popcs-b.pppcs)*(e.F_122+e.F_123)/1000 end as pwweight " +
                            "from [ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[Machine] c,item e " +
                            "where a.id = b.poid and c.mcode = a.omachinecode  and e.fbillno = a.oid and phour<>0.01  " +
                            "), " +
                            "main as  " +
                            "( " +
                            "select OWflow,oid, " +
                            "isnull(mcode,'')  as mcode, " +
                            "isnull(mname,'') as Mname , " +
                            "sum(popcs) as popcs,convert(decimal(18,2),sum(pokg)/1000) as pokg, " +
                            "sum(pppcs)as  pppcs,convert(decimal(18,2),sum(ppkg)/1000) as ppkg, " +
                            "convert(decimal(18,2),sum(pwweight)) as pwweight " +
                            "from ph " +
                            "where  OWflow is not null " +
                            "group by OWflow,mcode,oid,Mname " +
                            "), " +
                            "total as ( " +
                            "select oid,owflow, " +
                            "case  " +
                            "when grouping(OWflow)=0 and grouping(mcode)=1 then '小计' else isnull(mcode,'')end as mcode,isnull(mname,'') as mname, " +
                            "sum(popcs) as popcs,sum(pokg)as  pokg, " +
                            "sum(pppcs)as  pppcs,sum(ppkg)as  ppkg, " +
                            "sum(pwweight) as pwweight " +
                            "from main   " +
                            "group by oid,OWflow,mcode,mname with rollup), " +
                            "cleanNull as ( " +
                            "select oid,owflow,mcode,pokg,ppkg,pwweight from total where mcode='小计'), " +
                            "ppfront as(  " +
                            "select oid,owflow,pokg from total where mcode = '小计'), " +
                            "pohind as(  " +
                            "select oid,owflow,ppkg from total where mcode = '小计'), " +
                            "stockOut as ( " +
                            "select * from item as b inner join " +
                            "(select  b.fuse,SUM(b.FBaseQty) as FBaseQty,b.FBaseUnitID  " +
                            "from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_11] b,["+ sql.CYDB +"].[dbo].[T_ICItem] d  " +
                            "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and d.FitemID = a.FitemID and (d.Fnumber like '11%' or d.Fnumber like '12.C.02%')   " +
                            "group by b.fuse,b.FBaseUnitID) as a " +
                            "on b.fbillno = a.fuse " +
                            "union all " +
                            "select * from item as b inner join " +
                            "(select  b.fuse,SUM(b.FBaseQty) as FBaseQty,b.FBaseUnitID  " +
                            "from ["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_11] b,["+ sql.CKDB +"].[dbo].[T_ICItem] d  " +
                            "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and d.FitemID = a.FitemID and (d.Fnumber like '11%' or d.Fnumber like '12.C.02%')   " +
                            "group by b.fuse,b.FBaseUnitID) as a " +
                            "on b.fbillno = a.fuse " +
                            "), " +
                            "firstWorkFlow as ( " +
                            "select oid,owflow,fbaseqty-pokg as minus  from  " +
                            "(select  RN = ROW_NUMBER() OVER ( PARTITION BY oid ORDER BY owflow),oid,owflow,mcode,pokg from total where mcode='小计' ) a,stockout as b  where   " +
                            "b.fuse = a.oid  " +
                            "and  " +
                            "RN = 1  " +
                            "), " +
                            "final as  " +
                            "( select oid,wname,minus from ( " +
                            "select oid,owflow,minus from ( " +
                            "select RN = ROW_NUMBER() OVER ( PARTITION BY v1.oid,v1.OWFlow ORDER BY v1.owflow,v2.owflow),v1.oid,v2.OWFlow,v1.ppkg,v2.pokg,v1.ppkg-v2.pokg as minus  " +
                            "from pohind as v1,ppfront as v2  " +
                            "where v2.OWflow > v1.OWFlow and v1.oid = v2.oid )a where rn=1  " +
                            "union all " +
                            "select* from firstWorkFlow)a,[ChengyiYuntech].[dbo].[workflow] b where  owflow = b.wid ), " +
                            "pivotNodate as (SELECT * FROM final /*數據源*/ " +
                            "AS P " +
                            "PIVOT  " +
                            "( " +
                            "SUM(minus) FOR  " +
                            "p.wname/*需要行轉列的列*/ IN ([淋模],[印刷],[切纸],[覆瓦],[上光],[断张],[模切],[成型],[分条],[制袋],[包装],[离线裁切]/*列的值*/) " +
                            ") AS T " +
                            ") " +
                            "select fmodel, oid,isnull(a.淋模,0) as 淋模,isnull(a.印刷,0) as 印刷,isnull(a.切纸,0) as 切纸,isnull(a.覆瓦,0) as 覆瓦,isnull(a.上光,0) as 上光,isnull(a.断张,0) as 断张,isnull(a.模切,0) as 模切,isnull(a.成型,0) as 成型,isnull(a.分条,0) as 分条,isnull(a.制袋,0) as 制袋,isnull(a.包装,0) as 包装,isnull(a.离线裁切,0) as 离线裁切 " +
                            " ,case when Fstatus = 3 then '完成' else '下达中' end as '状态' from  " +
                            "pivotNodate as a, " +
                            "item as b " +
                            "where a.oid = b.fbillno and fcheckdate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and fmodel like '%" + txtFModel.Text +"%' order by fcheckdate,oid";

            DataTable dt = new DataTable();
            dt = sql.getQuery(query);

            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = item["fmodel"].ToString();
                dataGridView1.Rows[n].Cells[1].Value = item["oid"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = Convert.ToDecimal(item["淋模"]).ToString("N2");
                dataGridView1.Rows[n].Cells[3].Value = Convert.ToDecimal(item["印刷"]).ToString("N2");
                dataGridView1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["切纸"]).ToString("N2");
                dataGridView1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["覆瓦"]).ToString("N2");
                dataGridView1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["上光"]).ToString("N2");
                dataGridView1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["断张"]).ToString("N2");
                dataGridView1.Rows[n].Cells[8].Value = Convert.ToDecimal(item["模切"]).ToString("N2");
                dataGridView1.Rows[n].Cells[9].Value = Convert.ToDecimal(item["成型"]).ToString("N2");
                dataGridView1.Rows[n].Cells[10].Value = Convert.ToDecimal(item["分条"]).ToString("N2");
                dataGridView1.Rows[n].Cells[11].Value = Convert.ToDecimal(item["制袋"]).ToString("N2");
                dataGridView1.Rows[n].Cells[12].Value = Convert.ToDecimal(item["包装"]).ToString("N2");
                dataGridView1.Rows[n].Cells[13].Value = Convert.ToDecimal(item["离线裁切"]).ToString("N2");
                dataGridView1.Rows[n].Cells[14].Value = item["状态"].ToString();
            }
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                for (int i = 2; i < dataGridView1.ColumnCount - 1; i++)
                {
                    string execute = Convert.ToString(this.dataGridView1.Rows[e.RowIndex].Cells[i].Value);

                    if (Convert.ToDecimal(execute) < 0)
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[i].Style.BackColor = Color.Yellow;
                    }
                }

                string execute2 = Convert.ToString(this.dataGridView1.Rows[e.RowIndex].Cells[14].Value);

                if(execute2 == "下达中")
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightSkyBlue;
                }
            }
        }
    }
}
