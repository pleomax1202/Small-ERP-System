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

namespace Combination
{
    public partial class 发货通知检核表 : Form
    {
        public 发货通知检核表()
        {
            InitializeComponent();
        }

        Sql sql = new Sql();

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            dataGridView1.Rows.Clear();

            string query = @"select a.FFetchDate as 交貨日期,a.FSourceBillNo as 來源單號,a.FModel as 品名,a.FNumber as 物料長代碼,a.FAuxQty as 需求量,isnull(b.Qty,0) as 庫存量,case when a.FAuxQty > b.Qty then a.FAuxQty - b.Qty  else '0' end as 差額,a.FCUUnitName 單位 from
                            (select isnull(a.FFetchDate,'') as FFetchDate,isnull(a.FSourceBillNo,'') as FSourceBillNo,isnull(a.FModel,'') as FModel,isnull(a.FNumber,'') as FNumber,
                            isnull(a.FCUUnitName,'小计') as FCUUnitName,isnull(a.FitemID,'') as FitemID, isnull(a.FDefaultLoc,'') as FDefaultLoc,a.FAuxQty from (
                            select a.FFetchDate,c.FModel,c.FNumber,b.FCUUnitName,c.FitemID,c.FDefaultLoc,b.FSourceBillNo,SUM(a.FAuxQty) as FAuxQty from 
                            [" + sql.CYDB + "].[dbo].[SEOutStockEntry] a, " +
                            "["+ sql.CYDB +"].[dbo].[vwICBill_34] b, " +
                            "["+ sql.CYDB +"].[dbo].[t_ICItem] " +
                            "c " +
                            "where a.FInterID = b.FinterID and a.FEntryID = b.FEntryID and a.FitemID = c.FitemID and a.FFetchDate = '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "'  " +
                            "group by a.FFetchDate,c.FModel,c.FNumber,b.FCUUnitName,c.FitemID,c.FDefaultLoc,b.FSourceBillNo with rollup) a " +
                            "where a.FDefaultLoc is not null or(a.FNumber is not null and a.FCUUnitName is null)) a left join " +
                            "(select b.FItemID, b.Fmodel, b.Fnumber, d.Fname as unit, a.Qty/d.Fcoefficient as Qty, isnull(h.ww,0) as Uprice,a.FStockID from " +
                            "(select FitemID, FstockID, sum(FQty) as Qty from ["+ sql.CYDB +"].[dbo].[ICinventory] where (FstockID = '810' or FstockID = '20421') group by FitemID,FstockID) a left join " +
                            "(select c.FitemID, c.FNumber, SUM(a.FPrice* b.Fexchangerate)/ count(c.FNumber) as ww from " +
                            "["+ sql.CYDB +"].[dbo].[ICPrcPlyEntry] a, " +
                            "["+ sql.CYDB +"].[dbo].[t_Currency] b,  " +
                            "["+ sql.CYDB +"].[dbo].[t_icitem] c, " +
                            "["+ sql.CYDB +"].[dbo].[t_organization] " +
                            "d " +
                            "where a.FitemID = c.FitemID and a.FCuryID = b.FCurrencyID " +
                            "and d.Fdeleted = '0' and a.Fchecked = '1' and c.Fdeleted = '0' " +
                            "and c.FNumber like '14%' group by c.FitemID, c.FNumber) h on a.FitemID = h.FitemID,  " +
                            "["+ sql.CYDB +"].[dbo].[t_icitem] b, " +
                            "["+ sql.CYDB +"].[dbo].[t_stock] c, " +
                            "["+ sql.CYDB +"].[dbo].[t_measureunit] " +
                            "d " +
                            "where a.FitemID = b.FitemID and a.FstockID = c.FitemID " +
                            "and(b.FstoreUnitID = d.FMeasureUnitID and b.FUnitGroupID = d.FunitGroupID) " +
                            "and b.Fdeleted = '0'  and b.Fnumber like '14.%'  " +
                            "and (c.Fnumber = '03' or c.Fnumber = '04')) b on a.FitemID = b.FiteMID and a.FDefaultLoc = b.FStockID";

            DataTable dt = new DataTable();
            dt = sql.getQuery(query);

            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = Convert.ToDateTime(item["交貨日期"]).ToString("yyyy/MM/dd");
                dataGridView1.Rows[n].Cells[1].Value = item["來源單號"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = item["品名"].ToString();
                dataGridView1.Rows[n].Cells[3].Value = item["物料長代碼"].ToString();
                dataGridView1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["需求量"]).ToString("N0");
                if (item["庫存量"] != System.DBNull.Value)
                {
                    dataGridView1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["庫存量"]).ToString("N0");
                }
                dataGridView1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["差額"]).ToString("N0");
                dataGridView1.Rows[n].Cells[7].Value = item["單位"].ToString();
            }

            Cursor = Cursors.Default;

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("查无信息");
            }
        }

        private void dataGridView1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                string execute = Convert.ToString(this.dataGridView1.Rows[e.RowIndex].Cells["Column8"].Value);
                if (Convert.ToDecimal(execute) > 0)
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
        }
    }
}
