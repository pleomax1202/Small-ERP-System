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
    public partial class 任务单流程表 : Form
    {
        Sql sql = new Sql();

        public 任务单流程表()
        {
            InitializeComponent();
        }

        DateTime dateTime = DateTime.Now;
        string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string inputPath = System.Environment.CurrentDirectory;
        string exportPath;
        string filePath;
        Excel.Application excelApp;
        Excel._Workbook wBook;
        Excel._Worksheet wSheet;
        Excel.Range wRange;

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            Load_TitleData();
            if (checkbox1.Checked == true)
            {
                dgvHeaderText();
                dgv1.Columns[2].Visible = true;

                string query = @"with
                                item as 
                                (
                                select f_122,F_123,a.FitemID,Fnumber,b.fbillno from [" + sql.CYDB + "].[dbo].[T_ICItem]a , [" + sql.CYDB + "].[dbo].[ICMO]b " +
                                "where b.FITEMID = a.FitemID " +
                                "union all " +
                                "select f_122, F_123, a.FitemID,Fnumber,b.fbillno from["+ sql.CKDB +"].[dbo].[T_ICItem] a, ["+ sql.CKDB +"].[dbo].[ICMO] " +
                                "b " +
                                "where b.FITEMID =a.FitemID " +
                                "), " +
                                "ph as ( " +
                                "select OWflow, mcode, Mname, convert(varchar(10),odate,120) as pdate,a.oid,a.Opname,popcs,pppcs, " +
                                "case when e.Fnumber like '12.C%' then b.popcs* e.F_123 else b.popcs*(e.F_122+e.F_123) end as pokg, " +
                                "case when e.Fnumber like '12.C%' then b.pppcs* e.F_122 else b.pppcs*(e.F_122+e.F_123) end as ppkg, " +
                                "case when e.Fnumber like '12.C%' then (b.popcs-b.pppcs)*e.F_122/1000 else (b.popcs-b.pppcs)*(e.F_122+e.F_123)/1000 end as pwweight,/*克重乘以哪個*/ " +
                                "oorder " +
                                "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[Machine] c,item e " +
                                "where a.id = b.poid and c.mcode = a.omachinecode  and b.phour<>'0.01' and e.fbillno = a.oid " +
                                "),/*生產紀錄-機台單號產量換算*/ " +
                                "main as  " +
                                "( " +
                                "select OWflow, " +
                                "case when grouping(mcode) =1 and grouping(OWflow) =0 then '小計'  else isnull(mcode,'') end as mcode, " +
                                "isnull(mname,'') as Mname ,isnull(pdate,'') as pdate,isnull(oid,'') as oid , " +
                                "sum(popcs) as popcs,convert(decimal(18,2),sum(pokg)/1000) as pokg, " +
                                "sum(pppcs)as  pppcs,convert(decimal(18,2),sum(ppkg)/1000) as ppkg, " +
                                "convert(decimal(18,2),sum(pwweight)) as pwweight, " +
                                "convert(decimal(18,2),sum(pwweight)/(sum(pokg)/1000)*100) as pwpersent " +
                                "from ph " +
                                "where ph.oid like '%" + textBox1.Text + "' and OWflow is not null " +
                                "group by OWflow,mcode,pdate,oid,Mname with rollup " +
                                ")/*查詢條件-製造單號*/ " +
                                "select OWflow, mcode, Mname, pdate, popcs, pokg, pppcs, ppkg, pwweight, pwpersent from main where(pdate= '' and mcode = N'小計') or oid<>'' and Mname<>''  " +
                                "order by OWflow";


                          Load_Data(query);
            }
            if (checkbox2.Checked == true)
            {
                dgvHeaderText();
                dgv1.Columns[2].Visible = false;

                string query = @"with
                                item as 
                                (
                                select f_122,F_123,a.FitemID,Fnumber,b.fbillno from [" + sql.CYDB + "].[dbo].[T_ICItem]a , [" + sql.CYDB + "].[dbo].[ICMO]b " +
                                "where b.FITEMID = a.FitemID " +
                                "union all " +
                                "select f_122, F_123, a.FitemID,Fnumber,b.fbillno from["+ sql.CKDB +"].[dbo].[T_ICItem] a, ["+ sql.CKDB +"].[dbo].[ICMO] " +
                                "b " +
                                "where b.FITEMID =a.FitemID " +
                                "), " +
                                "ph as  " +
                                "( " +
                                "select OWflow, mcode, Mname, convert(varchar(10),odate,120) as pdate,a.oid,a.Opname,popcs,pppcs, " +
                                "case when e.Fnumber like '12.C%' then b.popcs* e.F_123 else b.popcs*(e.F_122+e.F_123) end as pokg, " +
                                "case when e.Fnumber like '12.C%' then b.pppcs* e.F_122 else b.pppcs*(e.F_122+e.F_123) end as ppkg, " +
                                "case when e.Fnumber like '12.C%' then (b.popcs-b.pppcs)*e.F_122/1000 else (b.popcs-b.pppcs)*(e.F_122+e.F_123)/1000 end as pwweight " +
                                "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[Machine] c,item e " +
                                "where a.id = b.poid and c.mcode = a.omachinecode  and e.fbillno = a.oid and phour<>0.01 " +
                                "), " +
                                "main as  " +
                                "( " +
                                "select OWflow, oid, " +
                                "isnull(mcode,'')  as mcode, " +
                                "isnull(mname,'') as Mname , " +
                                "sum(popcs) as popcs,convert(decimal(18,2),sum(pokg)/1000) as pokg, " +
                                "sum(pppcs)as  pppcs,convert(decimal(18,2),sum(ppkg)/1000) as ppkg, " +
                                "convert(decimal(18,2),sum(pwweight)) as pwweight, " +
                                "convert(decimal(18,2),sum(pwweight)/(sum(pokg)/1000)*100) as pwpersent " +
                                "from ph " +
                                "where ph.oid like '%" + textBox1.Text + "' and OWflow is not null " +
                                "group by OWflow,mcode,oid,Mname " +
                                "), " +
                                "total as  " +
                                "( " +
                                "select owflow, " +
                                "case  " +
                                "when grouping(OWflow)=0 and grouping(mcode)=1 then '小计' else isnull(mcode,'')end as mcode,isnull(mname,'') as mname, " +
                                "sum(popcs) as popcs,sum(pokg)as  pokg, " +
                                "sum(pppcs)as  pppcs,sum(ppkg)as  ppkg, " +
                                "sum(pwweight) as pwweight, " +
                                "sum(pwpersent)as pwpersent " +
                                "from main group by OWflow, mcode, mname with rollup " +
                                "), " +
                                "cleanNull as ( " +
                                "select* from total where mname<>'' or mcode = '小计'), " +
                                "/*0521 add the other pw calculation*/ " +
                                "ppfront as( " +
                                "select owflow, pokg from total where mcode = '小计'), " +
                                "pohind as( " +
                                "select owflow, ppkg from total where mcode = '小计'), " +
                                "calculatePw as ( " +
                                "select* from ( " +
                                "select RN = ROW_NUMBER() OVER(PARTITION BY v1.OWFlow ORDER BY v1.owflow, v2.owflow),v2.OWFlow,v1.ppkg,v2.pokg,v1.ppkg-v2.pokg as minus " +
                                "from pohind as v1,ppfront as v2 " +
                                "where v2.OWflow > v1.OWFlow)a where rn=1), " +
                                "stockOutTable as ( " +
                                "select* from item as b inner join " +
                                "(select b.fuse, SUM(b.FBaseQty) as FBaseQty, b.FBaseUnitID " +
                                "from [" + sql.CYDB + "].[dbo].[ICStockbillentry] a, [" + sql.CYDB + "].[dbo].[vwICBill_11] b, [" + sql.CYDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and d.FitemID = a.FitemID and (d.Fnumber like '11%' or d.Fnumber like '12.C.02%') " +
                                "and b.fuse like '%" + textBox1.Text + "' " +
                                "group by b.fuse, b.FBaseUnitID) as a " +
                                "on b.fbillno = a.fuse " +
                                "where b.fbillno like '%" + textBox1.Text + "' " +
                                "union all " +
                                "select* from item as b inner join " +
                                "(select b.fuse, SUM(b.FBaseQty) as FBaseQty, b.FBaseUnitID " +
                                "from [" + sql.CKDB + "].[dbo].[ICStockbillentry] a, [" + sql.CKDB + "].[dbo].[vwICBill_11] b, [" + sql.CKDB + "].[dbo].[T_ICItem] d " +
                                "where a.FinterID= b.FinterID and a.FentryID = b.FentryID and d.FitemID = a.FitemID and (d.Fnumber like '11%' or d.Fnumber like '12.C.02%') " +
                                "and b.fuse like '%" + textBox1.Text + "' " +
                                "group by b.fuse, b.FBaseUnitID) as a " +
                                "on b.fbillno = a.fuse " +
                                "where b.fbillno like '%" + textBox1.Text + "'), " +
                                "main2 as ( " +
                                "select b.*,case when mcode = '小计' then a.minus else '0' end as minus from calculatePw as a " +
                                "full outer join " +
                                "cleanNull as b " +
                                "on a.owflow = b.owflow), " +
                                "prepare as ( " +
                                "select owflow, mcode, mname, popcs, pokg, pppcs, ppkg, pwweight, pwpersent, " +
                                "case when minus is null   then(case when b.fbaseqty is null then '0' else b.fbaseqty-pokg end) " +
                                "else minus end as minus " +
                                "from main2 as a ,stockOutTable as b) " +
                                "select* " +
                                "from( " +
                                "select case when grouping(owflow) = 1 then '合计'  " +
                                "else mcode end as mcode, " +
                                "mname, " +
                                "case when grouping(owflow) = 1 then '0' else sum(popcs) end as popcs, " +
                                "case when grouping(owflow) = 1 then '0' else sum(pokg) end as pokg, " +
                                "case when grouping(owflow) = 1 then '0' else sum(pppcs) end as pppcs , " +
                                "case when grouping(owflow) = 1 then '0' else sum(ppkg) end as ppkg, " +
                                "case when grouping(owflow) = 1 then sum(pwweight)/2 else sum(pwweight) end as pw, " +
                                "case when grouping(owflow) = 1 then '0' else sum(pwpersent) end as pwp, " +
                                "sum(minus) as minus " +
                                "from prepare group by owflow, mcode, mname with rollup )a " +
                                "where mname is not null  or mcode = '合计'";

                 Load_Data(query);
            }
            if (checkbox3.Checked == true)
            {
                dgvHeaderText();
                dgv1.Columns[2].Visible = true;

                string query = @"with
                                item as 
                                (
                                select F_102,F_108,F_110,f_122,F_123,a.FitemID,Fnumber,b.fbillno from ["+ sql.CYDB +"].[dbo].[T_ICItem]a, ["+ sql.CYDB +"].[dbo].[ICMO] b " +
                                "where b.FITEMID =a.FitemID " +
                                "union all " +
                                "select F_102,F_108,F_110,f_122,F_123,a.FitemID,Fnumber,b.fbillno from ["+ sql.CKDB +"].[dbo].[T_ICItem]a, ["+ sql.CKDB +"].[dbo].[ICMO] b " +
                                "where b.FITEMID =a.FitemID " +
                                "), " +
                                "ph as  " +
                                "( " +
                                "select OWflow,mcode,Mname,munit,convert(varchar(10),Odate,120) as pdate,ohour,phour,pppcs,a.oid,Mspeed,e.* " +
                                "from [ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[Machine] c,item e " +
                                "where a.id = b.poid and c.mcode = a.omachinecode and e.fbillno = a.oid and phour<>0.01" +
                                "), " +
                                "main as  " +
                                "( " +
                                "select *,case when F_108 = '0' or F_110 = '0' or F_108 = '0' or F_102 = '0' then '0'  " +
                                "else (case when MUnit = 'KG' then  " +
                                "(case when Fnumber like '12.C%' then (Phour*MSpeed*60*1000)/(F_122) " +
                                "else (Phour*MSpeed*60*1000)/(F_122+F_123) end) " +
                                "when MUnit = '张'  then Phour*MSpeed*60*F_110 " +
                                "when MUnit = '箱'  then Phour*MSpeed*60*F_102 " +
                                "when MUnit = '米'  then Phour*(Mspeed*60*1000/F_108)*F_110 " +
                                "else Phour*MSpeed*60 end)  " +
                                "end as Pforecastpcs, " +
                                "case when F_108 = '0' or F_110 = '0' or F_108 = '0' or F_102 = '0' then '0'  " +
                                "else  " +
                                "(case when MUnit = 'KG' then  " +
                                "(case when Fnumber like '12.C%' then (Ohour*MSpeed*60*1000)/(F_122) " +
                                "else (Ohour*MSpeed*60*1000)/(F_122+F_123) end) " +
                                "when MUnit = '张'  then Ohour*MSpeed*60*F_110 " +
                                "when MUnit = '箱'  then Ohour*MSpeed*60*F_102 " +
                                "when MUnit = '米'  then Ohour*(Mspeed*60*1000/F_108)*F_110 " +
                                "else Ohour*MSpeed*60 end)  " +
                                "end as Oforecastpcs from ph " +
                                "where oid like '%" + textBox1.Text + "' " +
                                "), " +
                                "sumQuery as " +
                                "( " +
                                "select OWflow, case when grouping(mcode)= 1 and grouping(OWflow) = 0 then '小计' else mcode end as mcode " +
                                ",mname,pdate,sum(ohour) as ohour,sum(phour) as phour,sum(Oforecastpcs) as fopcs,sum(Pforecastpcs) as fppcs, " +
                                "sum(pppcs) as pppcs from main " +
                                "group by OWflow,mcode,mname,pdate with rollup " +
                                ") " +
                                "select " +
                                "mcode,isnull(mname, '') as mname,isnull(pdate, '') as pdate, " +
                                "isnull(ohour, '') as ohour,isnull(phour, '') as phour,convert(float, ceiling(fopcs * 100) / 100) as fopcs,convert(float, ceiling(fppcs * 100) / 100) as fppcs, " +
                                "convert(float, ceiling(pppcs * 100) / 100) as ppcs,convert(decimal(18, 2), (pppcs / fopcs) * 100) as jiadon from sumQuery " +
                                "where (mname <> '' and pdate<>'') or mcode = '小计'";

                Load_Data(query);
            }
            if (checkbox4.Checked == true)
            {
                dgvHeaderText();
                dgv1.Columns[2].Visible = false;

                string query = @"with
                                item as 
                                (
                                select F_102,F_108,F_110,f_122,F_123,a.FitemID,Fnumber,b.fbillno from ["+ sql.CYDB +"].[dbo].[T_ICItem]a , ["+ sql.CYDB +"].[dbo].[ICMO]b " +
                                "where b.FITEMID =a.FitemID " +
                                "union all " +
                                "select F_102,F_108,F_110,f_122,F_123,a.FitemID,Fnumber,b.fbillno from ["+ sql.CKDB +"].[dbo].[T_ICItem]a, ["+ sql.CKDB +"].[dbo].[ICMO] b " +
                                "where b.FITEMID =a.FitemID " +
                                "), " +
                                "ph as  " +
                                "( " +
                                "select OWflow,mcode,Mname,munit,convert(varchar(10),odate,120) as pdate,ohour,phour,pppcs,a.oid,Mspeed,e.* " +
                                "from [ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[Machine] c,item e " +
                                "where a.id = b.poid and c.mcode = a.omachinecode and e.fbillno = a.oid and phour<>0.01" +
                                "), " +
                                "main as ( " +
                                "select *,case when F_108 = '0' or F_110 = '0' or F_108 = '0' or F_102 = '0' then '0'  " +
                                "else (case when MUnit = 'KG' then  " +
                                "(case when Fnumber like '12.C%' then (Phour*MSpeed*60*1000)/(F_122) " +
                                "else (Phour*MSpeed*60*1000)/(F_122+F_123) end) " +
                                "when MUnit = '张'  then Phour*MSpeed*60*F_110 " +
                                "when MUnit = '箱'  then Phour*MSpeed*60*F_102 " +
                                "when MUnit = '米'  then Phour*(Mspeed*60*1000/F_108)*F_110 " +
                                "else Phour*MSpeed*60 end)  " +
                                "end as Pforecastpcs, " +
                                "case when F_108 = '0' or F_110 = '0' or F_108 = '0' or F_102 = '0' then '0'  " +
                                "else  " +
                                "(case when MUnit = 'KG' then  " +
                                "(case when Fnumber like '12.C%' then (Ohour*MSpeed*60*1000)/(F_122) " +
                                "else (Ohour*MSpeed*60*1000)/(F_122+F_123) end) " +
                                "when MUnit = '张'  then Ohour*MSpeed*60*F_110 " +
                                "when MUnit = '箱'  then Ohour*MSpeed*60*F_102 " +
                                "when MUnit = '米'  then Ohour*(Mspeed*60*1000/F_108)*F_110 " +
                                "else Ohour*MSpeed*60 end)  " +
                                "end as Oforecastpcs from ph " +
                                "where oid like '%" + textBox1.Text + "' " +
                                "), " +
                                "sumQuery as " +
                                "( " +
                                "select OWflow, case when grouping(mcode)= 1 and grouping(OWflow) = 0 then '小计' else mcode end as mcode " +
                                ",mname,sum(ohour) as ohour,sum(phour) as phour,sum(Oforecastpcs) as fopcs,sum(Pforecastpcs) as fppcs, " +
                                "sum(pppcs) as pppcs from main " +
                                "group by OWflow,mcode,mname with rollup " +
                                ") " +
                                "select mcode, isnull(mname, '') as mname, " +
                                "isnull(ohour, '') as ohour,isnull(phour, '') as phour,convert(float, ceiling(fopcs * 100) / 100) as fopcs,convert(float, ceiling(fppcs * 100) / 100) as fppcs, " +
                                "convert(float, ceiling(pppcs * 100) / 100) as ppcs, " +
                                "convert(decimal(18, 2), (pppcs / fopcs) * 100) as jiadon " +
                                "from sumQuery " +
                                "where mname <> ''  or mcode = '小计'";

                Load_Data(query);
                Cursor = Cursors.WaitCursor;
            }
            else if (checkbox1.Checked == false && checkbox2.Checked == false && checkbox3.Checked == false && checkbox4.Checked == false || textBox1.Text == "")
            {
                MessageBox.Show("请勾选选项或填入空值");
            }
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
            if (dgv1.Rows.Count != 0 && textBox1.Text != "")
            {
                if (checkbox1.Checked == true)
                {
                    exportPath = path + @"\生产任务单数量明细表导出";
                    filePath = inputPath + @"\生产任务单数量明细表";
                }
                else if (checkbox2.Checked == true)
                {
                    exportPath = path + @"\生产任务单数量汇总表导出";
                    filePath = inputPath + @"\生产任务单数量汇总表";
                }
                else if (checkbox3.Checked == true)
                {
                    exportPath = path + @"\生产任务单绩效明细表导出";
                    filePath = inputPath + @"\生产任务单绩效明细表";
                }
                else if (checkbox4.Checked == true)
                {
                    exportPath = path + @"\生产任务单绩效汇总表导出";
                    filePath = inputPath + @"\生产任务单绩效汇总表";
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

                if (checkbox1.Checked == true)
                {
                    wSheet.Name = "生产任务单数量明细表";

                    wSheet.Cells[2, 1] = "生产任务单数量明细表  " + "（" + textBox1.Text + "）";
                }
                else if (checkbox2.Checked == true)
                {
                    wSheet.Name = "生产任务单数量汇总表";

                    wSheet.Cells[2, 1] = "生产任务单数量汇总表  " + "（" + textBox1.Text + "）";
                }
                else if (checkbox3.Checked == true)
                {
                    wSheet.Name = "生产任务单绩效明细表";

                    wSheet.Cells[2, 1] = "生产任务单绩效明细表  " + "（" + textBox1.Text + "）";
                }
                else if (checkbox4.Checked == true)
                {
                    wSheet.Name = "生产任务单绩效汇总表";

                    wSheet.Cells[2, 1] = "生产任务单绩效汇总表  " + "（" + textBox1.Text + "）";
                }

                wRange = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // storing Each row and column value to excel sheet

                if (checkbox1.Checked == true || checkbox3.Checked == true)
                {
                    wSheet.Cells[3, 2] = textBox2.Text;
                    wSheet.Cells[3, 4] = "订单量（" + Unit1.Text + "）";
                    wSheet.Cells[3, 5] = txtOrder.Text;
                    wSheet.Cells[3, 6] = "入库量（" + Unit3.Text + "）";
                    wSheet.Cells[3, 7] = txtStorage.Text;
                    wSheet.Cells[3, 8] = "生产领料（" + Unit5.Text + "）";
                    wSheet.Cells[3, 9] = txtPicking.Text;
                    wSheet.Cells[4, 2] = txtAnnounce.Text;
                    wSheet.Cells[4, 5] = txtOrderPcs.Text;
                    wSheet.Cells[4, 7] = txtStoragePcs.Text;

                    for (int i = 0; i < dgv1.Rows.Count; i++)
                    {
                        wSheet.Cells[i + 6, 1] = Convert.ToString(dgv1.Rows[i].Cells[0].Value);
                        wSheet.Cells[i + 6, 2] = Convert.ToString(dgv1.Rows[i].Cells[1].Value);
                        if (Convert.ToString(dgv1.Rows[i].Cells[2].Value) == "小计")
                        {
                            wRange = wSheet.Range[wSheet.Cells[i + 6, 1], wSheet.Cells[i + 6, 9]];
                            wRange.Select();
                            wRange.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                        }
                        wSheet.Cells[i + 6, 3] = Convert.ToString(dgv1.Rows[i].Cells[2].Value);
                        wSheet.Cells[i + 6, 4] = Convert.ToString(dgv1.Rows[i].Cells[3].Value);
                        wSheet.Cells[i + 6, 5] = Convert.ToString(dgv1.Rows[i].Cells[4].Value);
                        wSheet.Cells[i + 6, 6] = Convert.ToString(dgv1.Rows[i].Cells[5].Value);
                        wSheet.Cells[i + 6, 7] = Convert.ToString(dgv1.Rows[i].Cells[6].Value);
                        wSheet.Cells[i + 6, 8] = Convert.ToString(dgv1.Rows[i].Cells[7].Value);
                        wSheet.Cells[i + 6, 9] = Convert.ToString(dgv1.Rows[i].Cells[8].Value);
                    }
                }
                else if (checkbox2.Checked == true)
                {
                    wSheet.Cells[3, 2] = textBox2.Text;
                    wSheet.Cells[3, 3] = "订单量（" + Unit1.Text + "）";
                    wSheet.Cells[3, 4] = txtOrder.Text;
                    wSheet.Cells[3, 5] = "入库量（" + Unit3.Text + "）";
                    wSheet.Cells[3, 6] = txtStorage.Text;
                    wSheet.Cells[3, 7] = "生产领料（" + Unit5.Text + "）";
                    wSheet.Cells[3, 8] = txtPicking.Text;
                    wSheet.Cells[4, 2] = txtAnnounce.Text;
                    wSheet.Cells[4, 4] = txtOrderPcs.Text;
                    wSheet.Cells[4, 6] = txtStoragePcs.Text;

                    for (int i = 0; i < dgv1.Rows.Count; i++)
                    {
                        wSheet.Cells[i + 6, 1] = Convert.ToString(dgv1.Rows[i].Cells[0].Value);
                        wSheet.Cells[i + 6, 2] = Convert.ToString(dgv1.Rows[i].Cells[1].Value);
                        if (Convert.ToString(dgv1.Rows[i].Cells[1].Value) == "小计")
                        {
                            wRange = wSheet.Range[wSheet.Cells[i + 6, 1], wSheet.Cells[i + 6, 8]];
                            wRange.Select();
                            wRange.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                        }
                        wSheet.Cells[i + 6, 3] = Convert.ToString(dgv1.Rows[i].Cells[3].Value);
                        wSheet.Cells[i + 6, 4] = Convert.ToString(dgv1.Rows[i].Cells[4].Value);
                        wSheet.Cells[i + 6, 5] = Convert.ToString(dgv1.Rows[i].Cells[5].Value);
                        wSheet.Cells[i + 6, 6] = Convert.ToString(dgv1.Rows[i].Cells[6].Value);
                        wSheet.Cells[i + 6, 7] = Convert.ToString(dgv1.Rows[i].Cells[7].Value);
                        wSheet.Cells[i + 6, 8] = Convert.ToString(dgv1.Rows[i].Cells[8].Value);
                        wSheet.Cells[i + 6, 9] = Convert.ToString(dgv1.Rows[i].Cells[9].Value);
                    }
                }
                else if (checkbox4.Checked == true)
                {
                    wSheet.Cells[3, 2] = textBox2.Text;
                    wSheet.Cells[3, 3] = "订单量（" + Unit1.Text + "）";
                    wSheet.Cells[3, 4] = txtOrder.Text;
                    wSheet.Cells[3, 5] = "入库量（" + Unit3.Text + "）";
                    wSheet.Cells[3, 6] = txtStorage.Text;
                    wSheet.Cells[3, 7] = "生产领料（" + Unit5.Text + "）";
                    wSheet.Cells[3, 8] = txtPicking.Text;
                    wSheet.Cells[4, 2] = txtAnnounce.Text;
                    wSheet.Cells[4, 4] = txtOrderPcs.Text;
                    wSheet.Cells[4, 6] = txtStoragePcs.Text;

                    for (int i = 0; i < dgv1.Rows.Count; i++)
                    {
                        wSheet.Cells[i + 6, 1] = Convert.ToString(dgv1.Rows[i].Cells[0].Value);
                        wSheet.Cells[i + 6, 2] = Convert.ToString(dgv1.Rows[i].Cells[1].Value);
                        if (Convert.ToString(dgv1.Rows[i].Cells[1].Value) == "小计")
                        {
                            wRange = wSheet.Range[wSheet.Cells[i + 6, 1], wSheet.Cells[i + 6, 8]];
                            wRange.Select();
                            wRange.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                        }
                        wSheet.Cells[i + 6, 3] = Convert.ToString(dgv1.Rows[i].Cells[3].Value);
                        wSheet.Cells[i + 6, 4] = Convert.ToString(dgv1.Rows[i].Cells[4].Value);
                        wSheet.Cells[i + 6, 5] = Convert.ToString(dgv1.Rows[i].Cells[5].Value);
                        wSheet.Cells[i + 6, 6] = Convert.ToString(dgv1.Rows[i].Cells[6].Value);
                        wSheet.Cells[i + 6, 7] = Convert.ToString(dgv1.Rows[i].Cells[7].Value);
                        wSheet.Cells[i + 6, 8] = Convert.ToString(dgv1.Rows[i].Cells[8].Value);
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
            else
            {
                MessageBox.Show("请确认是否有资料");
            }
        }

        private void Load_TitleData()
        {
            if (textBox1.Text != "")
            {
                String Fbillno = textBox1.Text;
                string query = @"select a.Fbillno,a.Fmodel,a.F_149,a.FAuxQty,a.需求pcs,a.Qty as 產出,a.PCS as 產出PCS,a.FCUUnitName,isnull(c.Qty,0) as 金蝶領料,isnull(c.FbaseUnitID,'kg') as 領料單位 from 
                                ((select v1.Fbillno,v1.Fmodel,v1.F_149,v1.FAuxQty,v1.需求pcs,isnull(v2.Qty,0) as Qty,V1.FName as FCUUnitName,isnull(v2.PCS,0) as PCS from (select a.Fbillno,b.Fmodel,a.FAuxQty,a.FAuxQty*b.F_102 as 需求pcs,a.FStatus,c.FName,b.F_149 from [" + sql.CYDB + "].[dbo].[ICMO] a,[" + sql.CYDB + "].[dbo].[t_ICItem] b,[" + sql.CYDB + "].[dbo].[t_measureUnit] c where a. FitemId = b.FitemID and a.FUnitID = c.FMeasureUnitID) v1 left join " +
                                "(select a.FbatchNo, d.Fmodel, Sum(a.FAuxQty) as Qty, b.FCUUnitName, Sum(a.FAuxQty) * b.FEntrySelfA0245 as PCS from["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_2] b,["+ sql.CYDB +"].[dbo].[t_ICItem] d where a.FinterID = b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID group by a.FbatchNo, d.Fmodel, b.FEntrySelfA0245, b.FCUUnitName) v2 on v1.Fbillno = v2.FbatchNo)  union " +
                                "(select v1.Fbillno, v1.Fmodel, v1.F_149, v1.FAuxQty, v1.需求pcs, isnull(v2.Qty, 0) as Qty, V1.FName as FCUUnitName, isnull(v2.PCS, 0) as PCS from(select a.Fbillno, b.Fmodel, a.FAuxQty, a.FAuxQty * b.F_102 as 需求pcs, a.FStatus, c.FName, b.F_149 from["+ sql.CKDB +"].[dbo].[ICMO] a,["+ sql.CKDB +"].[dbo].[t_ICItem] b,["+ sql.CKDB +"].[dbo].[t_measureUnit] c where a.FitemId = b.FitemID and a.FUnitID = c.FMeasureUnitID) v1 left join " +
                                "(select a.FbatchNo, d.Fmodel, Sum(isnull(a.FAuxQty, 0)) as Qty, b.FCUUnitName, Sum(isnull(a.FAuxQty, 0)) * b.FEntrySelfA0245 as PCS from["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_2] b,["+ sql.CKDB +"].[dbo].[t_ICItem] d where a.FinterID = b.FinterID and a.FentryID = b.FentryID and a.FitemID = d.FitemID group by a.FbatchNo, d.Fmodel, b.FEntrySelfA0245, b.FCUUnitName) v2 on v1.Fbillno = v2.Fbatchno)) a left join " +
                                "(select b.Fuse, SUM(b.FBaseQty) as Qty, b.FbaseUnitID  from( " +
                                "(select b.Fuse, d.Fmodel, SUM(b.FBaseQty) as FBaseQty, b.FBaseUnitID from["+ sql.CYDB +"].[dbo].[ICStockbillentry] a,["+ sql.CYDB +"].[dbo].[vwICBill_11] b,["+ sql.CYDB +"].[dbo].[ICStockbill] c,["+ sql.CYDB +"].[dbo].[T_ICItem] d where a.FinterID = b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and(d.Fnumber like '11%' or d.Fnumber like '12.C.02%') group by b.Fuse, d.Fmodel, b.FBaseUnitID) union " +
                                "(select b.Fuse, d.Fmodel, SUM(b.FcuQty) as FBaseQty, b.FUnitID as BaseUnitID from["+ sql.CYDB +"].[dbo].[ICSTJGBillEntry] a,["+ sql.CYDB +"].[dbo].[vwICBill_137] b,["+ sql.CYDB +"].[dbo].[ICSTJGBill] c,["+ sql.CYDB +"].[dbo].[T_ICItem] d where a.FinterID = b.FinterID and a.FEntryID = b.FEntryID and a.FinterID = c.FinterID and a.FItemID = d.FItemID and(d.Fnumber like '11%' or d.Fnumber like '12.C.02%')group by b.Fuse, d.Fmodel, b.FUnitID) union " +
                                "(select b.Fuse, d.Fmodel, SUM(b.FBaseQty) as FBaseQty, b.FBaseUnitID from["+ sql.CKDB +"].[dbo].[ICStockbillentry] a,["+ sql.CKDB +"].[dbo].[vwICBill_11] b,["+ sql.CKDB +"].[dbo].[ICStockbill] c,["+ sql.CKDB +"].[dbo].[T_ICItem] d where a.FinterID = b.FinterID and a.FentryID = b.FentryID and a.FinterID = c.FinterID and d.FitemID = a.FitemID and(d.Fnumber like '11%' or d.Fnumber like '12.C.02%') group by b.Fuse, d.Fmodel, b.FBaseUnitID) union " +
                                "(select b.Fuse, d.Fmodel, SUM(b.FcuQty) as FBaseQty, b.FUnitID as BaseUnitID from["+ sql.CKDB +"].[dbo].[ICSTJGBillEntry] a,["+ sql.CKDB +"].[dbo].[vwICBill_137] b,["+ sql.CKDB +"].[dbo].[ICSTJGBill] c,["+ sql.CKDB +"].[dbo].[T_ICItem] d where a.FinterID = b.FinterID and a.FEntryID = b.FEntryID and a.FinterID = c.FinterID and a.FItemID = d.FItemID and(d.Fnumber like '11%' or d.Fnumber like '12.C.02%')group by b.Fuse, d.Fmodel, b.FUnitID)) b group by b.Fuse,b.FbaseUnitID) c on a.Fbillno = c.Fuse " +
                                "where a.Fbillno like '%" + Fbillno + "'";

                DataTable dt = new DataTable();
                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {
                    textBox2.Text = item["Fmodel"].ToString();
                    txtAnnounce.Text = item["F_149"].ToString();
                    txtOrder.Text = Convert.ToDecimal(item["FAuxQty"]).ToString("N0");
                    txtOrderPcs.Text = Convert.ToDecimal(item["需求pcs"]).ToString("N0");
                    txtStorage.Text = Convert.ToDecimal(item["產出"]).ToString("N0");
                    txtStoragePcs.Text = Convert.ToDecimal(item["產出PCS"]).ToString("N0");
                    txtPicking.Text = Convert.ToDecimal(item["金蝶領料"]).ToString("N0");
                    Unit1.Text = item["FCUUnitName"].ToString();
                    Unit3.Text = item["FCUUnitName"].ToString();
                    Unit5.Text = item["領料單位"].ToString();
                }
            }
        }

        private void Load_Data(string query)
        {
            if (checkbox1.Checked == true)
            {
                if (textBox1.Text != "")
                {
                    dgv1.Rows.Clear();
                    DataTable dt = new DataTable();
                    dt = sql.getQuery(query);

                    foreach (DataRow item in dt.Rows)
                    {
                        int n = dgv1.Rows.Add();
                        dgv1.Rows[n].Cells[0].Value = item["mcode"].ToString();
                        dgv1.Rows[n].Cells[1].Value = item["mname"].ToString();
                        dgv1.Rows[n].Cells[2].Value = item["pdate"].ToString();
                        dgv1.Rows[n].Cells[3].Value = Convert.ToDecimal(item["popcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["pokg"]).ToString("N2");
                        dgv1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["pppcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["ppkg"]).ToString("N2");
                        dgv1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["pwweight"]).ToString("N2");
                        dgv1.Rows[n].Cells[8].Value = Convert.ToDecimal(item["pwpersent"]).ToString("");
                    }
                    if (dgv1.Rows.Count == 0)
                    {
                        MessageBox.Show("查无资料");
                    }
                }
                else
                {
                    MessageBox.Show("请输入任务单号");
                    textBox1.Focus();
                }
            }

            if (checkbox2.Checked == true)
            {
                if (textBox1.Text != "")
                {
                    dgv1.Rows.Clear();
                    DataTable dt = new DataTable();
                    dt = sql.getQuery(query);

                    foreach (DataRow item in dt.Rows)
                    {
                        int n = dgv1.Rows.Add();
                        dgv1.Rows[n].Cells[0].Value = item["mcode"].ToString();
                        dgv1.Rows[n].Cells[1].Value = Convert.ToString(item["mname"]);
                        dgv1.Rows[n].Cells[3].Value = Convert.ToDecimal(item["popcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["pokg"]).ToString("N2");
                        dgv1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["pppcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["ppkg"]).ToString("N2");
                        dgv1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["pw"]).ToString("N2");
                        dgv1.Rows[n].Cells[8].Value = Convert.ToString(item["pwp"]);
                        dgv1.Rows[n].Cells[9].Value = Convert.ToDecimal(item["minus"]).ToString("N2");
                    }

                    for (int i = 0; i < dgv1.Rows.Count; i++)
                    {
                        if(Convert.ToDecimal(dgv1.Rows[i].Cells[9].Value) == 0)
                        {
                            dgv1.Rows[i].Cells[9].Value = "";
                        }
                    }

                    try
                    {
                        int index = dgv1.Rows.Count;
                        dgv1.Rows[index - 1].Cells[3].Value = "";
                        dgv1.Rows[index - 1].Cells[4].Value = "";
                        dgv1.Rows[index - 1].Cells[5].Value = "";
                        dgv1.Rows[index - 1].Cells[6].Value = "";
                        dgv1.Rows[index - 1].Cells[8].Value = "";
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("查无信息");
                    }


                    if (dgv1.Rows.Count == 0)
                    {
                        MessageBox.Show("查无资料");
                    }
                }
                else
                {
                    MessageBox.Show("请输入任务单号");
                    textBox1.Focus();
                }
            }

            if (checkbox3.Checked == true)
            {
                if (textBox1.Text != "")
                {
                    dgv1.Rows.Clear();
                    DataTable dt = new DataTable();
                    dt = sql.getQuery(query);

                    foreach (DataRow item in dt.Rows)
                    {
                        int n = dgv1.Rows.Add();
                        dgv1.Rows[n].Cells[0].Value = item["mcode"].ToString();
                        dgv1.Rows[n].Cells[1].Value = item["mname"].ToString();
                        dgv1.Rows[n].Cells[2].Value = item["pdate"].ToString();
                        dgv1.Rows[n].Cells[3].Value = Convert.ToDecimal(item["ohour"]).ToString();
                        dgv1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["phour"]).ToString();
                        dgv1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["fopcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["fppcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["ppcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[8].Value = Convert.ToDecimal(item["jiadon"]).ToString("");
                    }
                    if (dgv1.Rows.Count == 0)
                    {
                        MessageBox.Show("查无资料");
                    }
                }
                else
                {
                    MessageBox.Show("请输入任务单号");
                    textBox1.Focus();
                }
            }

            if (checkbox4.Checked == true)
            {
                if (textBox1.Text != "")
                {
                    dgv1.Rows.Clear();
                    DataTable dt = new DataTable();
                    dt = sql.getQuery(query);

                    foreach (DataRow item in dt.Rows)
                    {
                        int n = dgv1.Rows.Add();
                        dgv1.Rows[n].Cells[0].Value = item["mcode"].ToString();
                        dgv1.Rows[n].Cells[1].Value = item["mname"].ToString();
                        dgv1.Rows[n].Cells[3].Value = Convert.ToDecimal(item["ohour"]).ToString();
                        dgv1.Rows[n].Cells[4].Value = Convert.ToDecimal(item["phour"]).ToString();
                        dgv1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["fopcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["fppcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["ppcs"]).ToString("N0");
                        dgv1.Rows[n].Cells[8].Value = Convert.ToDecimal(item["jiadon"]).ToString("");
                    }
                    if (dgv1.Rows.Count == 0)
                    {
                        MessageBox.Show("查无资料");
                    }
                }
                else
                {
                    MessageBox.Show("请输入任务单号");
                    textBox1.Focus();
                }
            }

        }

        private void checkbox1_OnChange(object sender, EventArgs e)
        {
            if (checkbox1.Checked == true)
            {
                checkbox2.Checked = false;
                checkbox3.Checked = false;
                checkbox4.Checked = false;
            }
        }

        private void checkbox2_OnChange(object sender, EventArgs e)
        {
            if (checkbox2.Checked == true)
            {
                checkbox1.Checked = false;
                checkbox3.Checked = false;
                checkbox4.Checked = false;
            }
        }

        private void checkbox3_OnChange(object sender, EventArgs e)
        {
            if (checkbox3.Checked == true)
            {
                checkbox1.Checked = false;
                checkbox2.Checked = false;
                checkbox4.Checked = false;
            }
        }

        private void checkbox4_OnChange(object sender, EventArgs e)
        {
            if (checkbox4.Checked == true)
            {
                checkbox1.Checked = false;
                checkbox2.Checked = false;
                checkbox3.Checked = false;
            }
        }

        private void dgvHeaderText()
        {
            if (checkbox1.Checked == true)
            {
                dgv1.Columns[0].HeaderText = "机台编号";
                dgv1.Columns[1].HeaderText = "机台名称";
                dgv1.Columns[3].HeaderText = "领用数量(pcs)";
                dgv1.Columns[4].HeaderText = "领用数量(kg)";
                dgv1.Columns[5].HeaderText = "完工数量(pcs)";
                dgv1.Columns[6].HeaderText = "完工数量(kg)";
                dgv1.Columns[7].HeaderText = "报废(kg)";
                dgv1.Columns[8].HeaderText = "损耗率(%)";
                dgv1.Columns[9].Visible = false;
            }
            else if (checkbox2.Checked == true)
            {
                dgv1.Columns[0].HeaderText = "机台编号";
                dgv1.Columns[1].HeaderText = "机台名称";
                dgv1.Columns[3].HeaderText = "领用数量(pcs)";
                dgv1.Columns[4].HeaderText = "领用数量(kg)";
                dgv1.Columns[5].HeaderText = "完工数量(pcs)";
                dgv1.Columns[6].HeaderText = "完工数量(kg)";
                dgv1.Columns[7].HeaderText = "生产报废（kg）";
                dgv1.Columns[8].HeaderText = "损耗率(%)";
                dgv1.Columns[9].HeaderText = "调机报废（kg）";
                dgv1.Columns[9].Visible = true;
            }
            else if (checkbox3.Checked == true || checkbox4.Checked == true)
            {
                dgv1.Columns[3].HeaderText = "排程工时";
                dgv1.Columns[4].HeaderText = "工作时数";
                dgv1.Columns[5].HeaderText = "预计产能(pcs)";
                dgv1.Columns[6].HeaderText = "应有产能(pcs)";
                dgv1.Columns[7].HeaderText = "实际产能(pcs)";
                dgv1.Columns[8].HeaderText = "稼动率(%)";
                dgv1.Columns[9].Visible = false;
            }

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = btnSeek;
            }
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            Load_TitleData();
        }
    }
}
