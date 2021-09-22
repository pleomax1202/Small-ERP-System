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

namespace Combination
{
    public partial class CheckExternalBoxAndLabel : Form
    {

        string querySum;
        string queryDetail, queryHeader;
        string path, inputPath, exportPath, filePath;
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;
        Sql sql = new Sql();


        public CheckExternalBoxAndLabel()
        {
            InitializeComponent();

        }

        private void ckbDetail_OnChange(object sender, EventArgs e)
        {
            if (ckbDetail.Checked == true)
            {
                dateTimePicker1.Enabled = false;
                dateTimePicker2.Enabled = false;
                label2.Visible = true;
                txtOrderCode.Visible = true;
                txtModel.Visible = true;
                txtOrderQty.Visible = true;
                txtStorageQty.Visible = true;
                txtBoxQty.Visible = true;
                txtLabelQty.Visible = true;
                label5.Visible = true;
                label6.Visible = true;
                label7.Visible = true;
                label8.Visible = true;
                label9.Visible = true;
                dgvTot.Visible = false;
                dgv2.Visible = true;
            }
            else
            {
                dateTimePicker1.Enabled = true;
                dateTimePicker2.Enabled = true;
                label2.Visible = false;
                txtOrderCode.Visible = false;
                txtModel.Visible = false;
                txtOrderQty.Visible = false;
                txtStorageQty.Visible = false;
                txtBoxQty.Visible = false;
                txtLabelQty.Visible = false;
                label5.Visible = false;
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                label9.Visible = false;
                dgvTot.Visible = true;
                dgv2.Visible = false;
            }
        }

        private void Load_HeaderData(string query)
        {
            DataTable dt = new DataTable();
            dt = sql.getQuery(query);

            foreach (DataRow item in dt.Rows)
            {
                txtModel.Text = item["產品名稱"].ToString();
                txtOrderQty.Text = Convert.ToDecimal(item["下單數"]).ToString("N0");
                txtStorageQty.Text = Convert.ToDecimal(item["入庫量"]).ToString("N0");
                txtBoxQty.Text = Convert.ToDecimal(item["外箱領料"]).ToString("N0");
                txtLabelQty.Text = Convert.ToDecimal(item["標籤領料"]).ToString("N0");
            }
        }

        public void Load_Data(string query)
        {
            Sql sql = new Sql();
            DataTable dt = new DataTable();
            dt = sql.getQuery(query);

            if(query == querySum)
            {
                foreach (DataRow item in dt.Rows)
                {
                    int n = this.dgvTot.Rows.Add();
                    dgvTot.Rows[n].Cells[0].Value = item["OID"].ToString();
                    dgvTot.Rows[n].Cells[1].Value = item["Fmodel"].ToString();
                    dgvTot.Rows[n].Cells[2].Value = Convert.ToDecimal(item["需求量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[3].Value = Convert.ToDecimal(item["已入庫量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[4].Value = Convert.ToDecimal(item["排程預計產量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[5].Value = Convert.ToDecimal(item["外箱已領數量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[6].Value = Convert.ToDecimal(item["標籤已領數量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[7].Value = Convert.ToDecimal(item["現場外箱余量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[8].Value = Convert.ToDecimal(item["現場標籤余量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[9].Value = Convert.ToDecimal(item["外箱應發料數量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[10].Value = Convert.ToDecimal(item["標籤應發料數量"]).ToString("N0");
                    dgvTot.Rows[n].Cells[11].Value = Convert.ToDecimal(item["外箱庫存"]).ToString("N0");
                    dgvTot.Rows[n].Cells[12].Value = Convert.ToDecimal(item["標籤庫存"]).ToString("N0");
                    dgvTot.Rows[n].Cells[13].Value = item["FbaseUnitID"].ToString();
                }
            }
            else if(query == queryDetail)
            {
                foreach (DataRow item in dt.Rows)
                {
                    int n = this.dgv2.Rows.Add();
                    dgv2.Rows[n].Cells[0].Value = item["日期"].ToString();
                    dgv2.Rows[n].Cells[1].Value = Convert.ToDecimal(item["Qty"]).ToString("N0");
                    dgv2.Rows[n].Cells[2].Value = item["FCUUnitName"].ToString();
                    dgv2.Rows[n].Cells[3].Value = Convert.ToDecimal(item["外箱當天領料"]).ToString("N0");
                    dgv2.Rows[n].Cells[4].Value = Convert.ToDecimal(item["標籤當天領料"]).ToString("N0");
                    dgv2.Rows[n].Cells[5].Value = Convert.ToDecimal(item["外箱累計余量"]).ToString("N0");
                    dgv2.Rows[n].Cells[6].Value = Convert.ToDecimal(item["標籤累計余量"]).ToString("N0");
                    dgv2.Rows[n].Cells[7].Value = item["FBaseUnitID"].ToString();
                }
            }
        }   

        private void btnSearch_Click(object sender, EventArgs e)
        {
            if (ckbDetail.Checked == true)
            {
                dgv2.Rows.Clear();
                queryHeader = "select c.FBillno as 製造單號,c.Fmodel as 產品名稱,c.FAuxQty as 下單數,isnull(a.Qty,0) as 入庫量,isnull(b.FBaseQty,0) as 外箱領料,isnull(d.FBaseQty,0) as 標籤領料 from  " +
                              "(select b.Fbillno, b.FAuxQty, b.Fmodel from((select a.Fbillno, a.FAuxQty, b.FMODEL from[" + sql.CYDB + "].[dbo].[ICMO] a,[" + sql.CYDB + "].[dbo].[T_ICITem] b where a.FitemID = b.FitemID) union all(select a.Fbillno, a.FAuxQty, b.FMODEL from[" + sql.CKDB + "].[dbo].[ICMO] a,[" + sql.CYDB + "].[dbo].[T_ICITem] b where a.FitemID = b.FitemID)) b) c left join " +
                              "(select a.FbatchNo, a.Fmodel, SUM(a.Qty) as Qty, a.FCUUnitName from(select * from((select a.FKFDate, a.FbatchNo, a.Fmodel, a.Qty, a.FCUUnitName, isnull(b.FChildQty, 0) as FChildQty   " +
                              "from(select d.Fnumber, a.FKFDate, a.FbatchNo, d.Fmodel, Sum(isnull(a.FAuxQty, 0)) as Qty, b.FCUUnitName, Sum(isnull(a.FAuxQty, 0)) * b.FEntrySelfA0245 as PCS " +
                              "from[" + sql.CYDB + "].[dbo].[ICStockbillentry] a,[" + sql.CYDB + "].[dbo].[vwICBill_2] b,[" + sql.CYDB + "].[dbo].[t_ICItem] d " +
                              "where a.FinterID = b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                              "group by a.FKFDate, a.FbatchNo, d.Fmodel, b.FEntrySelfA0245, b.FCUUnitName, d.Fnumber) a left join " +
                              "(select FNumber, FChildQty from[" + sql.CYDB + "].[dbo].[vICBOM] where Fchildmodel like '%贴纸%qs%' and FuseStatus = '使用')  b on a.FNumber = b.Fnumber) union all " +
                              "(select a.FKFDate, a.FbatchNo, a.Fmodel, a.Qty, a.FCUUnitName, b.FChildQty from " +
                              "(select d.Fnumber, a.FKFDate, a.FbatchNo, d.Fmodel, Sum(isnull(a.FAuxQty, 0)) as Qty, b.FCUUnitName, Sum(isnull(a.FAuxQty, 0)) * b.FEntrySelfA0245 as PCS " +
                              "from[" + sql.CKDB + "].[dbo].[ICStockbillentry] a,[" + sql.CKDB + "].[dbo].[vwICBill_2] b,[" + sql.CKDB + "].[dbo].[t_ICItem] d " +
                              "where a.FinterID = b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                              "group by a.FKFDate, a.FbatchNo, d.Fmodel, b.FEntrySelfA0245, b.FCUUnitName, d.Fnumber) a left join " +
                              "(select FNumber, FChildQty from[" + sql.CKDB + "].[dbo].[vICBOM] where Fchildmodel like '%贴纸%qs%' and FuseStatus = '使用')  b on a.FNumber = b.Fnumber)) a) a " +
                              "group by a.FbatchNo,a.Fmodel,a.FCUUnitName) a on a.Fbatchno = c.Fbillno left join " +
                              "(select* from ((select b.Fuse, d.Fmodel, SUM(b.FBaseQty) as FBaseQty,b.FBaseUnitID from[" + sql.CYDB + "].[dbo].[ICStockbillentry] a,[" + sql.CYDB + "].[dbo].[vwICBill_11] b,[" + sql.CYDB + "].[dbo].[ICStockbill] c,[" + sql.CYDB + "].[dbo].[T_ICItem] " +
                              "d " +
                              "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%外箱%'  group by b.Fuse,d.Fmodel,b.FBaseUnitID) Union all " +
                              "(select b.Fuse, d.Fmodel, SUM(b.FBaseQty) as FBaseQty, b.FBaseUnitID from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[ICStockbill] c, [" + sql.CKDB + "].[dbo].[T_ICItem] d " +
                              "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%外箱%'  group by b.Fuse, d.Fmodel, b.FBaseUnitID))b) b on c.FBillNo = b.FUse left join " +
                              "(select* from ((select b.Fuse, SUM(b.FBaseQty) as FBaseQty, b.FBaseUnitID from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a, [" + sql.CYDB + "].[dbo].[vwICBill_11] b, [" + sql.CYDB + "].[dbo].[ICStockbill] c, [" + sql.CYDB + "].[dbo].[T_ICItem] d " +
                              "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%贴纸%qs%'  group by b.Fuse, b.FBaseUnitID) Union all " +
                              "(select b.Fuse, SUM(b.FBaseQty) as FBaseQty, b.FBaseUnitID from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[ICStockbill] c, [" + sql.CKDB + "].[dbo].[T_ICItem] d " +
                              "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%贴纸%qs%'  group by b.Fuse, b.FBaseUnitID))b) d on d.Fuse = c.Fbillno " +
                              "where c.FbillNo = '" + txtOrderCode.Text + "'";

                queryDetail = @"select convert(char,V1.FKFdate,23) as 日期,V1.Qty,isnull(V1.FCUUnitName,'箱') as FCUUnitName,V1.FBaseQty as 外箱當天領料,V1.FBaseQty2 as 標籤當天領料,SUM(isnull(V2.余量,0)) as 外箱累計余量,SUM(isnull(V2.標籤余量,0)) as 標籤累計余量,isnull(V1.FBaseUnitID,'PCS') as FBaseUnitID from 
                                (select isnull(isnull(b.Fdate,a.FKFDate),c.FDate) as FKFDate,a.FBatchNo,isnull(a.Qty,0) as Qty,a.FCUUnitName,isNULL(b.FBaseQty,0) as FBaseQty,b.FBaseUnitID,isNULL(b.FBaseQty,0)-isnull(a.Qty,0) as 余量,isNULL(c.FBaseQty,0) as FBaseQty2,isNULL(c.FBaseQty,0)-isnull(a.Qty,0)*isnull(c.FChildQty,a.FChildQty) as 標籤余量 from 
                                (select * from((select a.FKFDate,a.FbatchNo,a.Fmodel,a.Qty,a.FCUUnitName,isnull(b.FChildQty,0) as FChildQty
                                 from (select d.Fnumber,a.FKFDate,a.FbatchNo,d.Fmodel,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS
                                from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a,[" + sql.CYDB + "].[dbo].[vwICBill_2] b,[" + sql.CYDB + "].[dbo].[t_ICItem] d " +
                                "where a.FinterID = b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo<> '' " +
                                "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName,d.Fnumber) a left join " +
                                "(select FNumber, FChildQty from[" + sql.CYDB + "].[dbo].[vICBOM] where Fchildmodel like '%贴纸%qs%' and FuseStatus = '使用')  b on a.FNumber = b.Fnumber) union all " +
                                "(select a.FKFDate, a.FbatchNo, a.Fmodel, a.Qty, a.FCUUnitName, b.FChildQty from " +
                                "(select d.Fnumber, a.FKFDate, a.FbatchNo, d.Fmodel, Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty, 0)) * b.FEntrySelfA0245 as PCS " +
                                "from[" + sql.CKDB + "].[dbo].[ICStockbillentry] a,[" + sql.CKDB + "].[dbo].[vwICBill_2] b,[" + sql.CKDB + "].[dbo].[t_ICItem] " +
                                "d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo<> '' " +
                                "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName,d.Fnumber) a left join " +
                                "(select FNumber, FChildQty from [" + sql.CKDB + "].[dbo].[vICBOM] where Fchildmodel like '%贴纸%qs%' and FuseStatus = '使用')  b on a.FNumber = b.Fnumber)) a " +
                                "where a.FbatchNo = '" + txtOrderCode.Text + "') a full outer join " +
                                "(select b.FDate, b.Fuse, b.Fmodel, SUM(b.FBaseQty) as FbaseQty, b.FBaseUnitID from ((select c.Fdate, b.Fuse, d.Fmodel, b.FBaseQty, b.FBaseUnitID from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a, [" + sql.CYDB + "].[dbo].[vwICBill_11] b, [" + sql.CYDB + "].[dbo].[ICStockbill] c, [" + sql.CYDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%外箱%' and b.Fuse = '" + txtOrderCode.Text + "') union all " +
                                "(select c.Fdate, b.Fuse, d.Fmodel, b.FBaseQty, b.FBaseUnitID from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[ICStockbill] c, [" + sql.CKDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%外箱%' and b.Fuse = '" + txtOrderCode.Text + "'))b group by b.FDate, b.Fuse, b.Fmodel, b.FBaseUnitID)b on a.FKFdate = b.Fdate full outer join " +
                                "(select a.Fbillno, a.FchildQty, b.Fdate, SUM(b.FbaseQty) as FbaseQty, b.FBaseUnitID from " +
                                "((select Fbillno, c.FchildQty from [" + sql.CYDB + "].[dbo].[ICMO] a, [" + sql.CYDB + "].[dbo].[t_ICItem] b, [" + sql.CYDB + "].[dbo].[vICBOM] c where a.FitemID = b.FitemID and b.FNumber = c.Fnumber and c.Fchildmodel like '%贴纸%qs%' and c.FuseStatus = '使用') union all " +
                                "(select Fbillno, c.FchildQty from [" + sql.CKDB + "].[dbo].[ICMO] a, [" + sql.CKDB + "].[dbo].[t_ICItem] b, [" + sql.CKDB + "].[dbo].[vICBOM] c where a.FitemID = b.FitemID and b.FNumber = c.Fnumber and c.Fchildmodel like '%贴纸%qs%' and c.FuseStatus = '使用'))a left join " +
                                "((select c.Fdate, b.Fuse, d.Fmodel, b.FBaseQty, b.FBaseUnitID from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a, [" + sql.CYDB + "].[dbo].[vwICBill_11] b, [" + sql.CYDB + "].[dbo].[ICStockbill] c, [" + sql.CYDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%贴纸%qs%') union all " +
                                "(select c.Fdate, b.Fuse, d.Fmodel, b.FBaseQty, b.FBaseUnitID from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[ICStockbill] c, [" + sql.CKDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%贴纸%qs%'))b on a.Fbillno = b.Fuse where a.Fbillno = '" + txtOrderCode.Text + "' " +
                                "group by a.Fbillno,a.FchildQty,b.Fdate,b.FBaseUnitID) c on isnull(a.FKFDate, b.Fdate) = c.Fdate) v1, " +
                                "(select isnull(isnull(a.FKFdate, b.FDate),c.Fdate) as FKFDate,a.FBatchNo,isnull(a.Qty,0) as Qty,a.FCUUnitName,isnull(isnull(b.Fdate, a.FKFDate), c.FDate) as FDate,isNULL(b.FBaseQty,0) as FBaseQty,isnull(b.FBaseUnitID,'PCS') as FBaseUnitID,isNULL(b.FBaseQty,0)-isnull(a.Qty,0) as 余量,isNULL(c.FBaseQty,0) as FBaseQty2,isNULL(c.FBaseQty,0)-isnull(a.Qty,0)*isnull(c.FChildQty, a.FChildQty) as 標籤余量 from " +
                                "(select* from((select a.FKFDate, a.FbatchNo, a.Fmodel, a.Qty, a.FCUUnitName, isnull(b.FChildQty,0) as FChildQty " +
                                "from(select d.Fnumber, a.FKFDate, a.FbatchNo, d.Fmodel, Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS " +
                                "from[" + sql.CYDB + "].[dbo].[ICStockbillentry] a,[" + sql.CYDB + "].[dbo].[vwICBill_2] b,[" + sql.CYDB + "].[dbo].[t_ICItem] " +
                                "d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo<> '' " +
                                "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName,d.Fnumber) a left join " +
                                "(select FNumber, FChildQty from [" + sql.CYDB + "].[dbo].[vICBOM] where Fchildmodel like '%贴纸%qs%' and FuseStatus = '使用')  b on a.FNumber = b.Fnumber) union all " +
                                "(select a.FKFDate, a.FbatchNo, a.Fmodel, a.Qty, a.FCUUnitName, b.FChildQty from " +
                                "(select d.Fnumber, a.FKFDate, a.FbatchNo, d.Fmodel, Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS " +
                                "from[" + sql.CKDB + "].[dbo].[ICStockbillentry] a,[" + sql.CKDB + "].[dbo].[vwICBill_2] b,[" + sql.CKDB + "].[dbo].[t_ICItem] " +
                                "d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo<> '' " +
                                "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName,d.Fnumber) a left join " +
                                "(select FNumber, FChildQty from [" + sql.CKDB + "].[dbo].[vICBOM] where Fchildmodel like '%贴纸%qs%' and FuseStatus = '使用')  b on a.FNumber = b.Fnumber)) a " +
                                "where a.FbatchNo = '" + txtOrderCode.Text + "') a full outer join " +
                                "(select b.FDate, b.Fuse, b.Fmodel, SUM(b.FBaseQty) as FbaseQty, b.FBaseUnitID from ((select c.Fdate, b.Fuse, d.Fmodel, b.FBaseQty, b.FBaseUnitID from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a, [" + sql.CYDB + "].[dbo].[vwICBill_11] b, [" + sql.CYDB + "].[dbo].[ICStockbill] c, [" + sql.CYDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%外箱%' and b.Fuse = '" + txtOrderCode.Text + "') union all " +
                                "(select c.Fdate, b.Fuse, d.Fmodel, b.FBaseQty, b.FBaseUnitID from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[ICStockbill] c, [" + sql.CKDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%外箱%' and b.Fuse = '" + txtOrderCode.Text + "'))b group by b.FDate, b.Fuse, b.Fmodel, b.FBaseUnitID)b on a.FKFdate = b.Fdate full outer join " +
                                "(select a.Fbillno, a.FchildQty, b.Fdate, SUM(b.FbaseQty) as FbaseQty, b.FBaseUnitID from " +
                                "((select Fbillno, c.FchildQty from [" + sql.CYDB + "].[dbo].[ICMO] a, [" + sql.CYDB + "].[dbo].[t_ICItem] b, [" + sql.CYDB + "].[dbo].[vICBOM] c where a.FitemID = b.FitemID and b.FNumber = c.Fnumber and c.Fchildmodel like '%贴纸%qs%' and c.FuseStatus = '使用') union all " +
                                "(select Fbillno, c.FchildQty from [" + sql.CKDB + "].[dbo].[ICMO] a, [" + sql.CKDB + "].[dbo].[t_ICItem] b, [" + sql.CKDB + "].[dbo].[vICBOM] c where a.FitemID = b.FitemID and b.FNumber = c.Fnumber and c.Fchildmodel like '%贴纸%qs%' and c.FuseStatus = '使用'))a left join " +
                                "((select c.Fdate, b.Fuse, d.Fmodel, b.FBaseQty, b.FBaseUnitID from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a, [" + sql.CYDB + "].[dbo].[vwICBill_11] b, [" + sql.CYDB + "].[dbo].[ICStockbill] c, [" + sql.CYDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%贴纸%qs%') union all " +
                                "(select c.Fdate, b.Fuse, d.Fmodel, b.FBaseQty, b.FBaseUnitID from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[ICStockbill] c, [" + sql.CKDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%贴纸%qs%'))b on a.Fbillno = b.Fuse where a.Fbillno = '" + txtOrderCode.Text + "' " +
                                "group by a.Fbillno,a.FchildQty,b.Fdate,b.FBaseUnitID) c on isnull(a.FKFDate, b.Fdate) = c.Fdate) v2 " +
                                "where v1.FKFDate > v2.FKFDate or(v1.FKFDate = v2.FKFDate and v1.Qty = V2.Qty and v1.FBaseQty = V2.FBaseQty) " +
                                "Group by V1.FKFdate,V1.Qty,V1.FCUUnitName,V1.FBaseQty,v1.FBaseUnitID,v1.FBaseQty2";

                Load_HeaderData(queryHeader);
                Load_Data(queryDetail);
            }
            else
            {
                dgvTot.Rows.Clear();
                querySum = @"select v1.OID,v1.Fmodel,v2.FAuxQty as 需求量,isnull(v2.Qty,0) as 已入庫量,v1.應發料數量 as 排程預計產量,isnull(v2.FCUUnitName,'箱') as FCUUnitName,isnull(v2.FbaseQty,0) as 外箱已領數量,isnull(v2.FbaseQty2,0) as 標籤已領數量,
                                isnull(v2.FbaseQty,0)-isnull(v2.Qty,0) as 現場外箱余量,isnull(v2.FbaseQty2,0)-isnull(v2.Qty,0)*isnull(v2.FChildQty,0) as 現場標籤余量,
                                case when (v1.應發料數量-(isnull(v2.FbaseQty,0)-isnull(v2.Qty,0))) <0 then 0 else (v1.應發料數量-(isnull(v2.FbaseQty,0)-isnull(v2.Qty,0))) end as 外箱應發料數量,
                                case when (v1.應發料數量*isnull(v2.FChildQty,0)-(isnull(v2.FbaseQty2,0)-isnull(v2.Qty,0)*isnull(v2.FChildQty,0))) <0 then 0 else (v1.應發料數量*isnull(v2.FChildQty,0)-(isnull(v2.FbaseQty2,0)-isnull(v2.Qty,0)*isnull(v2.FChildQty,0))) end as 標籤應發料數量,isnull(v2.Qty1,0) as 外箱庫存,isnull(v2.Qty2,0) as 標籤庫存,
                                isnull(v2.FbaseUnitID,'PCS') as FbaseUnitID from 
                                (Select b.OID,b.Fmodel,SUM(b.OHour) as 累積排程工時,SUM(b.應發料數量) as 應發料數量 from ((select OID,d.Fmodel,a.Ohour as OHour,left(a.OMachineCode,2) as 機台名稱,
                                Round((case when b.Munit = 'KG' then (case when d.FNumber like '12.C%' then a.Ohour*b.MSpeed*60/(F_122) else  a.Ohour*b.MSpeed*60/(F_122+F_123) end)
                                when b.Munit = '米' then (a.Ohour*b.MSpeed*60*d.F_110*1000)/(d.F_108*d.F_102) when b.Munit = '张' then a.Ohour*b.MSpeed*60*d.F_110/d.F_102 when b.Munit = '箱' then a.Ohour*b.MSpeed*60 else a.Ohour*b.MSpeed*60/d.F_102 end),0)  as 應發料數量
                                from [chengyiYuntech].[dbo].[ProduceOrder] a,[chengyiYuntech].[dbo].[Machine] b,[" + sql.CYDB + "].[dbo].[ICMO] c,[" + sql.CYDB + "].[dbo].[T_ICItem] d " +
                                "where a.OSample = '0' and a.OMachineCode = b.Mcode and b.MoutUnit = '箱' and a.OID = c.FbillNo and ODate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and c.FitemID= d.FitemID and (d.F_102 <> '0' or d.F_108<> '0' or d.F_110<> '0')) union all " +
                                "(select OID, d.Fmodel, a.Ohour as OHour, left(a.OMachineCode, 2) as 機台名稱, " +
                                "Round((case when b.Munit = 'KG' then(case when d.FNumber like '12.C%' then a.Ohour * b.MSpeed * 60 / (F_122) else  a.Ohour * b.MSpeed * 60 / (F_122 + F_123) end) " +
                                "when b.Munit = '米' then(a.Ohour * b.MSpeed * 60 * d.F_110 * 1000) / (d.F_108 * d.F_102) when b.Munit = '张' then a.Ohour * b.MSpeed * 60 * d.F_110 / d.F_102 when b.Munit = '箱' then a.Ohour * b.MSpeed * 60 else a.Ohour * b.MSpeed * 60 / d.F_102 end), 0) as 應發料數量 " +
                                "from[chengyiYuntech].[dbo].[ProduceOrder] a,[chengyiYuntech].[dbo].[Machine] b,[" + sql.CKDB + "].[dbo].[ICMO] c,[" + sql.CKDB + "].[dbo].[T_ICItem] " +
                                "d " +
                                "where a.OSample = '0' and a.OMachineCode = b.Mcode and b.MoutUnit = '箱' and a.OID = c.FbillNo and ODate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and c.FitemID= d.FitemID and (d.F_102<> '0' or d.F_108<> '0' or d.F_110<> '0'))) b group by b.OID, b.Fmodel) v1 left join " +
                                "(select* from ((select a.Fbillno, a.Fmodel, a.FauxQty, b.Qty, b.FCUUnitName, c.FbaseQty, c.FbaseUnitID, d.FbaseQty as FbaseQty2, d.FbaseUnitID as FbaseUnitID2, c.Qty as Qty1, d.Qty as Qty2, e.FChildQty from " +
                                "(select a.Fbillno, a.FAuxQty, b.FMODEL, b.FNumber from [" + sql.CYDB + "].[dbo].[ICMO] a, [" + sql.CYDB + "].[dbo].[T_ICITem] b where a.FitemID = b.FitemID and b.Fnumber Like '14%') a left join " +
                                "(select a.FbatchNo, a.Fmodel, SUM(a.Qty) as Qty, a.FCUUnitName " +
                                "from(select a.FKFDate, a.FbatchNo, d.Fmodel, Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS " +
                                "from[" + sql.CYDB + "].[dbo].[ICStockbillentry] a,[" + sql.CYDB + "].[dbo].[vwICBill_2] b,[" + sql.CYDB + "].[dbo].[t_ICItem] " +
                                "d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo<> ''  " +
                                "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName) a group by a.FbatchNo, a.Fmodel, a.FCUUnitName) b on a.Fbillno = b.FbatchNo left join " +
                                "(select b.Fuse, SUM(b.FBaseQty) as FBaseQty, e.Qty, b.FBaseUnitID " +
                                "from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a, [" + sql.CYDB + "].[dbo].[vwICBill_11] b, [" + sql.CYDB + "].[dbo].[ICStockbill] c, [" + sql.CYDB + "].[dbo].[T_ICItem] d, (select FitemID, sum(FQty) as Qty from [" + sql.CYDB + "].[dbo].[ICinventory] where FStockID = '809' group by FitemID) e " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%外箱%' and e.FitemID = a.FitemID " +
                                "group by b.Fuse, b.FBaseUnitID, e.Qty) c on a.Fbillno = c.Fuse left join " +
                                "(select b.Fuse, SUM(b.FBaseQty) as FBaseQty, e.Qty, b.FBaseUnitID " +
                                "from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a, [" + sql.CYDB + "].[dbo].[vwICBill_11] b, [" + sql.CYDB + "].[dbo].[ICStockbill] c, [" + sql.CYDB + "].[dbo].[T_ICItem] d, (select FitemID, sum(FQty) as Qty from [" + sql.CYDB + "].[dbo].[ICinventory] where FStockID = '809' group by FitemID) e " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '贴纸%QS%' and e.FitemID = a.FitemID " +
                                "group by b.Fuse, b.FBaseUnitID, e.Qty) d on a.Fbillno = d.Fuse left join " +
                                "(select* from [" + sql.CYDB + "].[dbo].[vICBOM] where Fchildmodel like '%贴纸%qs%' and FuseStatus = '使用') e on a.Fnumber = e.Fnumber) union all " +
                                "(select a.Fbillno, a.Fmodel, a.FauxQty, b.Qty, b.FCUUnitName, c.FbaseQty, c.FbaseUnitID, d.FbaseQty  as FbaseQty2, d.FbaseUnitID as FbaseUnitID2, c.Qty as Qty1, d.Qty as Qty2, e.FChildQty from " +
                                "(select a.Fbillno, a.FAuxQty, b.FMODEL, b.FNumber from [" + sql.CKDB + "].[dbo].[ICMO] a, [" + sql.CKDB + "].[dbo].[T_ICITem] b where a.FitemID = b.FitemID) a left join " +
                                "(select a.FbatchNo, a.Fmodel, SUM(a.Qty) as Qty, a.FCUUnitName " +
                                "from(select a.FKFDate, a.FbatchNo, d.Fmodel, Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS " +
                                "from[" + sql.CKDB + "].[dbo].[ICStockbillentry] a,[" + sql.CKDB + "].[dbo].[vwICBill_2] b,[" + sql.CKDB + "].[dbo].[t_ICItem] " +
                                "d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo<> ''  " +
                                "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName) a group by a.FbatchNo, a.Fmodel, a.FCUUnitName) b on a.Fbillno = b.FbatchNo left join " +
                                "(select b.Fuse, SUM(b.FBaseQty) as FBaseQty, e.Qty, b.FBaseUnitID " +
                                "from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[ICStockbill] c, [" + sql.CKDB + "].[dbo].[T_ICItem] d, (select FitemID, sum(FQty) as Qty from [" + sql.CKDB + "].[dbo].[ICinventory] where FStockID = '19761' group by FitemID) e " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '%外箱%' and e.FitemID = a.FitemID " +
                                "group by b.Fuse, b.FBaseUnitID, e.Qty) c on a.Fbillno = c.Fuse left join " +
                                "(select b.Fuse, SUM(b.FBaseQty) as FBaseQty, e.Qty, b.FBaseUnitID " +
                                "from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[ICStockbill] c, [" + sql.CKDB + "].[dbo].[T_ICItem] d, (select FitemID, sum(FQty) as Qty from [" + sql.CKDB + "].[dbo].[ICinventory] where FStockID = '19761' group by FitemID) e " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and d.Fmodel like '贴纸%QS%' and e.FitemID = a.FitemID " +
                                "group by b.Fuse, b.FBaseUnitID, e.Qty) d on a.Fbillno = d.Fuse left join " +
                                "(select* from [" + sql.CKDB + "].[dbo].[vICBOM] where Fchildmodel like '%贴纸%qs%' and FuseStatus = '使用') e on a.Fnumber = e.Fnumber))a)  " +
                                "v2 on v1.OID = V2.FbillNo";

                Load_Data(querySum);
            }

            if (dgvTot.Rows.Count == 0 && ckbDetail.Checked == false)
            {
                MessageBox.Show("查无信息");
            }
            else if (dgv2.Rows.Count == 0 && ckbDetail.Checked == true)
            {
                MessageBox.Show("查无信息");
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
                path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                inputPath = System.Environment.CurrentDirectory;

                if(ckbDetail.Checked == true)
                {
                    exportPath = path + @"\标签外箱发料入庫明细表导出";
                    filePath = inputPath + @"\标签外箱发料入庫明细表";
                }
                else
                {
                    exportPath = path + @"\标签外箱发料入庫汇总表导出";
                    filePath = inputPath + @"\标签外箱发料入庫汇总表";
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

                if (ckbDetail.Checked == true)
                {
                    wSheet.Name = "标签外箱发料入庫明细表";
                    wSheet.Cells[3, 3] = txtOrderCode.Text;
                    wSheet.Cells[3, 6] = txtModel.Text;
                    wSheet.Cells[4, 2] = txtOrderQty.Text;
                    wSheet.Cells[4, 4] = txtStorageQty.Text;
                    wSheet.Cells[4, 6] = txtBoxQty.Text;
                    wSheet.Cells[4, 8] = txtLabelQty.Text;

                    for (int i = 0; i < dgv2.Rows.Count; i++)
                    {
                        for (int j = 0; j < dgv2.ColumnCount; j++)
                        {
                            wSheet.Cells[i + 6, j + 1] = Convert.ToString(dgv2.Rows[i].Cells[j].Value);
                        }
                    }
                }
                else
                {
                    wSheet.Name = "标签外箱发料入庫汇总表";

                    for (int i = 0; i < dgvTot.Rows.Count; i++)
                    {
                        for (int j = 0; j < dgvTot.ColumnCount; j++)
                        {
                            wSheet.Cells[i + 4, j + 1] = Convert.ToString(dgvTot.Rows[i].Cells[j].Value);
                        }
                    }
                }

                wRange = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                // storing Each row and column value to excel sheet



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
    }
}
