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
    public partial class 入庫檢核表 : Form
    {
        string query1, query2, query3, query4, query5;
        string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string inputPath = System.Environment.CurrentDirectory;
        string exportPath;
        string filePath;
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;
        Sql sql = new Sql();

        public 入庫檢核表()
        {
            InitializeComponent();
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            if (TotCheckedBox.Checked == true)
            {
                this.dataGridView3.Rows.Clear();

                query4 = @"/*▼金蝶入库汇总▼*/
		                        select fbatchno as 制造单号,fmodel as 产品名称,数量 AS 生产数量,FNAME as 单位,入库数量,单位,convert (decimal(18,0),v3.fauxqty) as 订单数量 from 
			                        (select FbatchNO,convert (decimal(18,0),sum(qty)) as 入库数量, FCUUnitName as 单位 from
				                        ((select a.FKFDate,a.FbatchNo,d.Fmodel,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS 
				                        from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_2] b,["+ sql.CYDB +"].[dbo].[t_ICItem] d " +
                                        "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                                        "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName)  " +
                                        "union " +
                                        "(select a.FKFDate,a.FbatchNo,d.Fmodel,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS  " +
                                        "from ["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_2] b,["+ sql.CKDB +"].[dbo].[t_ICItem] d " +
                                        "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                                        "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName) ) v where (FCUUnitName='箱' or FCUUnitName='板' ) " +
                                        "group by fbatchno,fmodel,fcuunitname)v1,  " +
                                        "/*程奎程益生产*/	  " +
                                        "(select fbillno,fmodel,sum(ppqty) as 数量,fname from " +
                                        "(SELECT c.FBILLNO,a.[FModel],b.FNAME FROM ["+ sql.CYDB +"].[dbo].[t_ICItem]a, ["+ sql.CYDB +"].[dbo].[t_measureUnit]b,["+ sql.CYDB +"].[dbo].[ICMO]c " +
                                        "where b.FmeasureUnitID=a.FStoreUnitID " +
                                        "and a.fitemid=c.FitemID)a,/*入庫單位*/ " +
                                        "((select convert(varchar(10),fcheckdate,120)as pdate,oid,case when moutunit ='张' then ppqty*f_110/f_102 else ppqty  end as ppqty, " +
                                        "case when moutunit ='张' then '板'else moutunit end  as unit ,f_110,F_122,F_123,F_102 " +
                                        "from [ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[scanrecord]b,[ChengyiYuntech].[dbo].[Machine] c, " +
                                        "["+ sql.CYDB +"].[dbo].[ICMO] d,["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                        "where a.oid = d.fbillno  " +
                                        "and d.FitemID = e.FitemID  " +
                                        "and a.omachinecode=c.mcode " +
                                        "and a.id=b.poid " +
                                        "and e.f_102 <> 0)/*CY*/ " +
                                        "union all " +
                                        "(select convert(varchar(10),fcheckdate,120)as pdate,oid,case when moutunit ='张' then ppqty*f_110/f_102 else ppqty  end as ppqty, " +
                                        "case when moutunit ='张' then '板'else moutunit end  as unit ,f_110,F_122,F_123,F_102 " +
                                        "from [ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[scanrecord]b,[ChengyiYuntech].[dbo].[Machine] c, " +
                                        "["+ sql.CKDB +"].[dbo].[ICMO] d,["+ sql.CKDB +"].[dbo].[T_ICItem] e " +
                                        "where a.oid = d.fbillno  " +
                                        "and d.FitemID = e.FitemID  " +
                                        "and a.omachinecode=c.mcode " +
                                        "and a.id=b.poid " +
                                        "and e.f_102 <> 0) /*CK*/ )b " +
                                        "where b.oid=a.fbillno " +
                                        "and unit =fname " +
                                        "and pdate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' " +
                                        "and fmodel like '%" + textBox5.Text + "%' " +
                                        "group by fbillno,fmodel,fname)v2,[" + sql.CYDB + "].[dbo].[ICMO] " +
                                        "v3 " +
                                        "where v2.fbillno=v1.FbatchNO " +
                                        "and v3.fbillno = v2.fbillno";


                Load_Data(query4);

                if (dataGridView3.Rows.Count == 0)
                {
                    MessageBox.Show("查无信息");
                }
            }
            else if(ckbDate.Checked == true)
            {
                query5 = @"with v0 as 
                        (select  b.FKFDATE,a.FbatchNo,d.Fmodel,isnull(a.FAuxQty,0) as Qty,b.FCUUnitName,isnull(a.FAuxQty,0)*b.FEntrySelfA0245 as PCS  from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_2] b,["+ sql.CYDB +"].[dbo].[t_ICItem] d " +
                        "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                        "union all " +
                        "select  b.FKFDATE,a.FbatchNo,d.Fmodel,isnull(a.FAuxQty,0) as Qty,b.FCUUnitName,isnull(a.FAuxQty,0)*b.FEntrySelfA0245 as PCS  from ["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_2] b,["+ sql.CKDB +"].[dbo].[t_ICItem] d " +
                        "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' ), " +
                        "storeNum as  " +
                        "(select convert(varchar(10),FKFDATE,120) as idate,FbatchNo,Fmodel,convert(decimal(18,0),Sum(Qty)) as Qty,FCUUnitName,convert(decimal(18,0),Sum(PCS)) as PCS from v0  " +
                        "where (fcuunitname ='箱' or fcuunitname = '板') and (FKFdate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "') " +
                        "group by FKFDATE,FbatchNo,Fmodel,FCUUnitName ), " +
                        "inUnit as " +
                        "(SELECT c.fbillno,a.[FModel],b.FNAME " +
                        "FROM["+ sql.CYDB +"].[dbo].[t_ICItem] a, ["+ sql.CYDB +"].[dbo].[t_measureUnit] b,["+ sql.CYDB +"].[dbo].[ICMO] " +
                        "c " +
                        "where b.FmeasureUnitID=a.FStoreUnitID " +
                        "and a.fitemid=c.FitemID " +
                        "union all " +
                        "SELECT c.fbillno,a.[FModel],b.FNAME " +
                        "FROM["+ sql.CKDB +"].[dbo].[t_ICItem] a, ["+ sql.CKDB +"].[dbo].[t_measureUnit] b,["+ sql.CKDB +"].[dbo].[ICMO] " +
                        "c " +
                        "where b.FmeasureUnitID=a.FStoreUnitID " +
                        "and a.fitemid=c.FitemID), " +
                        "productInfo as ( " +
                        "select convert(varchar(10),odate,120)as pdate,oid,c.moutunit,case when moutunit = '张' then ppqty*f_110/f_102 else ppqty end as ppqty, " +
                        "case when moutunit = '张' then '板'else moutunit end  as unit ,f_110,F_122,F_123,F_102 " +
                        "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[scanrecord] b,[ChengyiYuntech].[dbo].[Machine] c, " +
                        "["+ sql.CYDB +"].[dbo].[ICMO] d,["+ sql.CYDB +"].[dbo].[T_ICItem] " +
                        "e " +
                        "where  a.oid = d.fbillno " +
                        "and d.FitemID = e.FitemID " +
                        "and a.omachinecode=c.mcode " +
                        "and a.id=b.poid " +
                        "and e.f_102<> 0 " +
                        "and phour<>'0.01' " +
                        "union all " +
                        "select convert(varchar(10),odate,120)as pdate,oid,c.moutunit,case when moutunit = '张' then ppqty*f_110/f_102 else ppqty end as ppqty, " +
                        "case when moutunit = '张' then '板'else moutunit end  as unit ,f_110,F_122,F_123,F_102 " +
                        "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[scanrecord] b,[ChengyiYuntech].[dbo].[Machine] c, " +
                        "["+ sql.CKDB +"].[dbo].[ICMO] d,["+ sql.CKDB +"].[dbo].[T_ICItem] " +
                        "e " +
                        "where  a.oid = d.fbillno " +
                        "and d.FitemID = e.FitemID " +
                        "and a.omachinecode=c.mcode " +
                        "and a.id=b.poid " +
                        "and e.f_102<> 0 " +
                        "and phour<>'0.01' ), " +
                        "productInfo2 as ( " +
                        "select pdate, oid, sum(ppqty) as ppqty,unit from productInfo a, inUnit b where(pdate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "') and(unit = '箱' or unit = '板')   and b.Fbillno = a.oid " +
                        "and a.unit= b.fname " +
                        "group by pdate, oid, unit) " +
                        "select isnull(convert(varchar(10),pdate,120),idate) as 日期,isnull(oid,'') as 生产单号,isnull(unit,'') as 单位, " +
                        "isnull(convert(varchar, convert(float, ceiling(ppqty*100)/100)),'') as 生产数量, " +
                        "isnull(convert(varchar, qty),'') as 入库数量,isnull(fcuunitname,'') as 单位,isnull(FbatchNo,'') as 入库单号 " +
                        "from productInfo2 " +
                        "full outer join storeNum " +
                        "on oid = fbatchno and pdate = idate and unit = fcuunitname " +
                        "order by pdate desc,oid";

                Load_Data(query5);
            }
            else
            {
                this.dataGridView1.Rows.Clear();
                this.dataGridView2.Rows.Clear();

                query1 = @"select isnull(pdate,'合计') as 日期,convert(decimal(18,0),ppqty)as 生产数量,isnull(unit,'') as 生产单位,
			                            convert(decimal(18,2),storenum)as 入库数量,case when storeunit='张' then '板' else isnull(storeunit,'') end as 入库单位,convert(decimal(18,2),ppcs)as PCS 
			                            from 
				                            (select  pdate ,sum(ppqty) as ppqty,unit,sum(storenum)as storenum,storeunit,sum(ppcs) as ppcs from
					                            (select Pdate,ppqty,unit,case when unit='张' then ppqty*f_110/f_102 else ppqty end as storenum,
					                            case when unit ='张' then '板' else unit end as storeunit,
					                            case when unit ='张' then ppqty*f_110 else ppqty*f_102 end as ppcs   from
						                            ((select a.FbatchNo,case when b.FCUUnitName='板' then '张' else b.FCUUnitName end as FCUUnitName
						                            from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_2] b,["+ sql.CYDB +"].[dbo].[t_ICItem] d " +
                                                    "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                                                    "and FbatchNo  like  '%" + textBox1.Text + "') " +
                                                    "union " +
                                                    "(select a.FbatchNo,case when b.FCUUnitName = '板' then '张' else b.FCUUnitName end as FCUUnitName " +
                                                    "from["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_2] b,["+ sql.CKDB +"].[dbo].[t_ICItem] d " +
                                                    "where a.FinterID = b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                                                    "and FbatchNo like  '%" + textBox1.Text + "')) a ,/*K3程益程奎入库*/ " +
                                                    "((select oid, convert(char, odate, 23) as Pdate, b.ppqty, c.moutunit as unit, f_110, F_122, F_123, F_102 " +
                                                    "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[scanrecord] b,[ChengyiYuntech].[dbo].[Machine] c, " +
                                                    "["+ sql.CYDB +"].[dbo].[ICMO] d,["+ sql.CYDB +"].[dbo].[T_ICItem] " +
                                                    "e " +
                                                    "where a.oid = d.fbillno " +
                                                    "and d.FitemID = e.FitemID " +
                                                    "and a.omachinecode=c.mcode " +
                                                    "and a.id=b.poid " +
                                                    "and oid like  '%" + textBox1.Text + "' )/*CY*/ " +
                                                    "union all " +
                                                    "(select oid, convert(char, odate,23) as Pdate,b.ppqty,c.moutunit as unit ,f_110,F_122,F_123,F_102 " +
                                                    "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[scanrecord] b,[ChengyiYuntech].[dbo].[Machine] c, " +
                                                    "["+ sql.CKDB +"].[dbo].[ICMO] d,["+ sql.CKDB +"].[dbo].[T_ICItem] " +
                                                    "e " +
                                                    "where a.oid = d.fbillno " +
                                                    "and d.FitemID = e.FitemID " +
                                                    "and a.omachinecode=c.mcode " +
                                                    "and a.id=b.poid " +
                                                    "and oid like  '%" + textBox1.Text + "') /*CK*/ )b     /*程奎程益生产*/ " +
                                                    "where b.unit=a.FCUUnitName)c /*b生产资料换算*/ " +
                                                    "group by pdate, unit, storeunit with rollup)d " +
                                                    "where storeunit is not null or pdate is null";

                query2 = @"select  isnull(日期,'合计')as 日期,入库数量,isnull(单位,'') as 单位 ,PCS from
		                (select 日期,sum(入库数量) as 入库数量,单位,sum(PCS)  as PCS from
			                (select a.FbatchNo,convert(char,FKFDate,23) as 日期,convert (decimal(18,0),qty) as 入库数量, FCUUnitName as 单位,convert (decimal(18,2),PCS) as PCS  from
				                ((select a.FKFDate,a.FbatchNo,d.Fmodel,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS 
				                from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_2] b,["+ sql.CYDB +"].[dbo].[t_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                                "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName)  " +
                                "union " +
                                "(select a.FKFDate,a.FbatchNo,d.Fmodel,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS  " +
                                "from ["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_2] b,["+ sql.CKDB +"].[dbo].[t_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID and FbatchNo <> '' " +
                                "group by a.FKFDate,a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName)) a " +
                                "where a.FbatchNo like '%" + textBox1.Text + "')b " +
                   "group by 日期,单位 with rollup)c " +
                   "where 单位 is not null or 日期 is null";

                query3 = @"select a.Fmodel as 产品名称,convert(decimal(18,0),a.FAuxQty) as 订单量,a.FCUUnitName as 订单单位,isnull(convert(decimal(18,0),c.Qty),0) as 金蝶领料,isnull(c.FbaseUnitID,'kg') as 领料单位 from 
                    ((select v1.Fbillno,v1.Fmodel,v1.F_149,v1.FAuxQty,v1.需求pcs,isnull(v2.Qty,0) as Qty,V1.FName as FCUUnitName,isnull(v2.PCS,0) as PCS from (select a.Fbillno,b.Fmodel,a.FAuxQty,a.FAuxQty*b.F_102 as 需求pcs,a.FStatus,c.FName,b.F_149 from ["+ sql.CYDB +"].[dbo].[ICMO] a,["+ sql.CYDB +"].[dbo].[t_ICItem] b,["+ sql.CYDB +"].[dbo].[t_measureUnit] c where a. FitemId = b.FitemID and a.FUnitID = c.FMeasureUnitID) v1 left join " +
                    "(select a.FbatchNo,d.Fmodel,Sum(a.FAuxQty) as Qty,b.FCUUnitName,Sum(a.FAuxQty)*b.FEntrySelfA0245 as PCS from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_2] b,["+ sql.CYDB +"].[dbo].[t_ICItem] d where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID group by a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName) v2 on v1.Fbillno = v2.FbatchNo)  union " +
                    "(select v1.Fbillno,v1.Fmodel,v1.F_149,v1.FAuxQty,v1.需求pcs,isnull(v2.Qty,0) as Qty,V1.FName as FCUUnitName,isnull(v2.PCS,0) as PCS from (select a.Fbillno,b.Fmodel,a.FAuxQty,a.FAuxQty*b.F_102 as 需求pcs,a.FStatus,c.FName,b.F_149 from ["+ sql.CKDB +"].[dbo].[ICMO] a,["+ sql.CKDB +"].[dbo].[t_ICItem] b,["+ sql.CKDB +"].[dbo].[t_measureUnit] c where a. FitemId = b.FitemID and a.FUnitID = c.FMeasureUnitID) v1 left join " +
                    "(select a.FbatchNo,d.Fmodel,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS from ["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_2] b,["+ sql.CKDB +"].[dbo].[t_ICItem] d where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID group by a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName) v2 on v1.Fbillno = v2.Fbatchno)) a left join " +
                    "(select b.Fuse,SUM(b.FBaseQty) as Qty,b.FbaseUnitID  from ((select b.Fuse,d.Fmodel,SUM(b.FBaseQty) as FBaseQty,b.FBaseUnitID from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_11] b,["+ sql.CYDB +"].[dbo].[ICStockbill] c,["+ sql.CYDB +"].[dbo].[T_ICItem] d where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and (d.Fnumber like '11%' or d.Fnumber like '12.C%') group by b.Fuse,d.Fmodel,b.FBaseUnitID) union " +
                    "(select b.Fuse,d.Fmodel,SUM(b.FBaseQty) as FBaseQty,b.FBaseUnitID from ["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_11] b,["+ sql.CKDB +"].[dbo].[ICStockbill] c,["+ sql.CKDB +"].[dbo].[T_ICItem] d where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and (d.Fnumber like '11%' or d.Fnumber like '12.C%') group by b.Fuse,d.Fmodel,b.FBaseUnitID)) b group by b.Fuse,b.FbaseUnitID) c on a.Fbillno = c.Fuse " +
                    "where a.Fbillno = '" + textBox1.Text + "' ";

                Load_Data(query3);
                Load_Data(query1);
                Load_Data(query2);
                Cursor = Cursors.Default;
                Check_DgvData();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

        }

        private void btnExport_Click_1(object sender, EventArgs e)
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

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = btnSeek;
            }
        }

        private void Export_Data()
        {

            if (TotCheckedBox.Checked == true)
            {
                exportPath = path + @"\生产入库检核汇总表导出";
                filePath = inputPath + @"\生产入库检核汇总表";
            }
            else if (ckbDate.Checked == true)
            {
                exportPath = path + @"\生产日期入库汇总表导出";
                filePath = inputPath + @"\生产日期入库汇总表";
            }
            else
            {
                exportPath = path + @"\生产与入库检核表导出";
                filePath = inputPath + @"\生产与入库检核表";
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

            if (TotCheckedBox.Checked == true)
            {
                wSheet.Name = "生产入库汇总表";
                wSheet.Cells[2, 1] = "生产入库汇总表     (" + dateTimePicker1.Value.ToString("yyyy/MM/dd") + ") - (" + dateTimePicker2.Value.ToString("yyyy/MM/dd") + ")";
                wRange = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView3.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 4, j + 1] = Convert.ToString(dataGridView3.Rows[i].Cells[j].Value);
                    }
                }
            }
            else if (ckbDate.Checked == true)
            {
                wSheet.Name = "生产日期入库汇总表";
                wSheet.Cells[2, 1] = "生产日期入库汇总表";
                wRange = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                for (int i = 0; i < dgvDate.Rows.Count; i++)
                {
                    for (int j = 0; j < dgvDate.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 4, j + 1] = Convert.ToString(dgvDate.Rows[i].Cells[j].Value);
                    }
                }
            }
            else
            {
                wSheet.Name = "生产入库检核表";
                wSheet.Cells[2, 1] = "生产入库检核表";
                wRange = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                wSheet.Cells[3, 2] = textBox1.Text;
                wSheet.Cells[3, 6] = textBox3.Text;
                wSheet.Cells[4, 2] = textBox2.Text;
                wSheet.Cells[4, 4] = UnitOrder.Text;



                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 7, j + 1] = Convert.ToString(dataGridView1.Rows[i].Cells[j].Value);
                    }
                }
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 7, j + 7] = Convert.ToString(dataGridView2.Rows[i].Cells[j].Value);
                    }
                }
            }

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

        private void textBox1_Leave(object sender, EventArgs e)
        {
            query3 = @"select a.Fmodel as 产品名称,convert(decimal(18,0),a.FAuxQty) as 订单量,a.FCUUnitName as 订单单位,isnull(convert(decimal(18,0),c.Qty),0) as 金蝶领料,isnull(c.FbaseUnitID,'kg') as 领料单位 from 
                    ((select v1.Fbillno,v1.Fmodel,v1.F_149,v1.FAuxQty,v1.需求pcs,isnull(v2.Qty,0) as Qty,V1.FName as FCUUnitName,isnull(v2.PCS,0) as PCS from (select a.Fbillno,b.Fmodel,a.FAuxQty,a.FAuxQty*b.F_102 as 需求pcs,a.FStatus,c.FName,b.F_149 from ["+ sql.CYDB +"].[dbo].[ICMO] a,["+ sql.CYDB +"].[dbo].[t_ICItem] b,["+ sql.CYDB +"].[dbo].[t_measureUnit] c where a. FitemId = b.FitemID and a.FUnitID = c.FMeasureUnitID) v1 left join " +
                    "(select a.FbatchNo,d.Fmodel,Sum(a.FAuxQty) as Qty,b.FCUUnitName,Sum(a.FAuxQty)*b.FEntrySelfA0245 as PCS from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_2] b,["+ sql.CYDB +"].[dbo].[t_ICItem] d where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID group by a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName) v2 on v1.Fbillno = v2.FbatchNo)  union " +
                    "(select v1.Fbillno,v1.Fmodel,v1.F_149,v1.FAuxQty,v1.需求pcs,isnull(v2.Qty,0) as Qty,V1.FName as FCUUnitName,isnull(v2.PCS,0) as PCS from (select a.Fbillno,b.Fmodel,a.FAuxQty,a.FAuxQty*b.F_102 as 需求pcs,a.FStatus,c.FName,b.F_149 from ["+ sql.CKDB +"].[dbo].[ICMO] a,["+ sql.CKDB +"].[dbo].[t_ICItem] b,["+ sql.CKDB +"].[dbo].[t_measureUnit] c where a. FitemId = b.FitemID and a.FUnitID = c.FMeasureUnitID) v1 left join " +
                    "(select a.FbatchNo,d.Fmodel,Sum(isnull(a.FAuxQty,0)) as Qty,b.FCUUnitName,Sum(isnull(a.FAuxQty,0))*b.FEntrySelfA0245 as PCS from ["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_2] b,["+ sql.CKDB +"].[dbo].[t_ICItem] d where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID group by a.FbatchNo,d.Fmodel,b.FEntrySelfA0245,b.FCUUnitName) v2 on v1.Fbillno = v2.Fbatchno)) a left join " +
                    "(select b.Fuse,SUM(b.FBaseQty) as Qty,b.FbaseUnitID  from ((select b.Fuse,d.Fmodel,SUM(b.FBaseQty) as FBaseQty,b.FBaseUnitID from ["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_11] b,["+ sql.CYDB +"].[dbo].[ICStockbill] c,["+ sql.CYDB +"].[dbo].[T_ICItem] d where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and (d.Fnumber like '11%' or d.Fnumber like '12.C%') group by b.Fuse,d.Fmodel,b.FBaseUnitID) union " +
                    "(select b.Fuse,d.Fmodel,SUM(b.FBaseQty) as FBaseQty,b.FBaseUnitID from ["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_11] b,["+ sql.CKDB +"].[dbo].[ICStockbill] c,["+ sql.CKDB +"].[dbo].[T_ICItem] d where a.FinterID= b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and (d.Fnumber like '11%' or d.Fnumber like '12.C%') group by b.Fuse,d.Fmodel,b.FBaseUnitID)) b group by b.Fuse,b.FbaseUnitID) c on a.Fbillno = c.Fuse " +
                    "where a.Fbillno = '" + textBox1.Text + "' ";

            Load_Data(query3);
        }

        private void dgvDate_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                string execute = Convert.ToString(this.dgvDate.Rows[e.RowIndex].Cells[3].Value);
                string execute2 = Convert.ToString(this.dgvDate.Rows[e.RowIndex].Cells[4].Value);
                if (execute != execute2)
                {
                    dgvDate.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
            }
        }

        private void ckbDate_OnChange(object sender, EventArgs e)
        {
            if(ckbDate.Checked == true)
            {
                label1.Text = "生产日期";
                dateTimePicker1.Visible = true;
                dateTimePicker2.Visible = true;
                textBox1.Visible = false;
                textBox2.Visible = false;
                textBox3.Visible = false;
                textBox5.Visible = false;
                label2.Visible = false;
                label3.Visible = false;
                label6.Visible = true;
                UnitOrder.Visible = false;
                tableLayoutPanel1.Visible = false;
                TotCheckedBox.Checked = false;
                dgvDate.Visible = true;
            }
            else
            {
                label1.Text = "制造单号";
                textBox1.Visible = true;
                textBox5.Visible = false;
                label3.Visible = true;
                label3.Text = "产品名称";
                textBox3.Visible = true;
                dateTimePicker1.Visible = false;
                label6.Visible = false;
                dateTimePicker2.Visible = false;
                label2.Visible = true;
                textBox2.Visible = true;
                UnitOrder.Visible = true;
                tableLayoutPanel1.Visible = true;
                dgvDate.Visible = false;
                ckbDate.Checked = false;
            }
        }

        private void TotCheckedBox_OnChange(object sender, EventArgs e)
        {
            if(TotCheckedBox.Checked == true)
            {
                label1.Text = "制单日期";
                textBox1.Visible = false;
                textBox5.Visible = true;
                label3.Visible = true;
                label3.Text = "产品名称";
                textBox3.Visible = false;
                dateTimePicker1.Visible = true;
                label6.Visible = true;
                dateTimePicker2.Visible = true;
                label2.Visible = false;
                textBox2.Visible = false;
                UnitOrder.Visible = false;
                tableLayoutPanel1.Visible = false;
                dgvDate.Visible = false;
                ckbDate.Checked = false;
            }
            else
            {
                label1.Text = "制造单号";
                textBox1.Visible = true;
                textBox5.Visible = false;
                label3.Text = "产品名称";
                textBox3.Visible = true;
                dateTimePicker1.Visible = false;
                label6.Visible = false;
                dateTimePicker2.Visible = false;
                label2.Visible = true;
                textBox2.Visible = true;
                UnitOrder.Visible = true;
                tableLayoutPanel1.Visible = true;
                dgvDate.Visible = false;
                ckbDate.Checked = false;
            }
        }

        private void Check_DgvData()
        {
            if(dataGridView1.Rows.Count == 0 || dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("查无信息");
            }
        }

        private void Load_Data(string query)
        {
            if (query == query1)
            {
                DataTable dt = new DataTable();
                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = Convert.ToString(item["日期"]);
                    dataGridView1.Rows[n].Cells[1].Value = Convert.ToDecimal(item["生产数量"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[2].Value = item["生产单位"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = Convert.ToDecimal(item["入库数量"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[4].Value = item["入库单位"].ToString();
                    dataGridView1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["PCS"]).ToString("N0");
                }
            }
            else if (query == query2)
            {
                DataTable dt = new DataTable();
                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    int n = dataGridView2.Rows.Add();
                    dataGridView2.Rows[n].Cells[0].Value = Convert.ToString(item["日期"]);
                    dataGridView2.Rows[n].Cells[1].Value = Convert.ToDecimal(item["入库数量"]).ToString("N0");
                    dataGridView2.Rows[n].Cells[2].Value = item["单位"].ToString();
                    dataGridView2.Rows[n].Cells[3].Value = Convert.ToDecimal(item["PCS"]).ToString("N0");
                }
            }
            else if(query == query3)
            {
                DataTable dt = new DataTable();
                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    textBox3.Text = Convert.ToString(item["产品名称"]);
                    textBox2.Text = Convert.ToString(item["订单量"]);
                    UnitOrder.Text = Convert.ToString(item["订单单位"]);
                }
            }
            else if (query == query4)
            {
                DataTable dt = new DataTable();
                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    int n = dataGridView3.Rows.Add();
                    dataGridView3.Rows[n].Cells[0].Value = Convert.ToString(item["制造单号"]);
                    dataGridView3.Rows[n].Cells[1].Value = Convert.ToString(item["产品名称"]);
                    dataGridView3.Rows[n].Cells[2].Value = Convert.ToDecimal(item["生产数量"]).ToString("N0");
                    dataGridView3.Rows[n].Cells[3].Value = Convert.ToString(item["单位"]);
                    dataGridView3.Rows[n].Cells[4].Value = Convert.ToDecimal(item["入库数量"]).ToString("N0");
                    dataGridView3.Rows[n].Cells[5].Value = Convert.ToString(item["单位"]);
                    dataGridView3.Rows[n].Cells[6].Value = Convert.ToString(item["订单数量"]);
                }
            }
            else if(query == query5)
            {
                DataSet ds = new DataSet();
                ds = sql.SqlCmdDS(query);

                dgvDate.DataSource = ds.Tables[0];
                dgvDate.Columns[5].HeaderText = "单位";
                if(dgvDate.Rows.Count == 0)
                {
                    MessageBox.Show("查无信息");
                }
            }
        }
    }
}
