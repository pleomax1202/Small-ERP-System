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
    public partial class 机台绩效报表 : Form
    {
        public 机台绩效报表()
        {
            InitializeComponent();
        }

        Sql sql = new Sql();

        private void btnSeek_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            if (textBox1.Text != "")
            {
                dataGridView1.Columns[1].Visible = true;
                dataGridView1.Columns[2].Visible = true;

                string query = @"select * from
                                ((select a.Mname,convert(char,a.Odate,23) as Odate,a.OOrder,SUM(a.OHour) as OHour,SUM(a.Phour) as PHour,SUM(a.POPcs) as POPcs,SUM(a.PPPcs) as PPPcs,SUM(PPPcs)/SUM(a.速度達成率) as 速度達成率,Sum(a.PPPcs)/SUM(a.績效) as 績效,(Sum(a.POPcs)-Sum(a.PPPcs))/Sum(a.POPcs) as 損耗率,SUM(a.PWWeight) as PWWeight from 
                                ((select c.Mname,convert(char,b.Odate,23) as Odate,b.OOrder,b.OHour,a.Phour,a.POPcs,a.PPPcs,
                                case when e.F_108 = '0' or e.F_110 = '0' or e.F_108 = '0' or e.F_102 = '0' then '0' else (case when c.MUnit = 'KG' then (case when e.Fnumber like '12.C%' then ((a.Phour*c.MSpeed*60*1000)/(e.F_122)) else ((a.Phour*c.MSpeed*60*1000)/(e.F_122+e.F_123)) end)
                                when c.MUnit = '张'  then (a.Phour*c.MSpeed*60*e.F_110) when c.MUnit = '箱'  then (a.Phour*c.MSpeed*60*e.F_102) when c.MUnit = '米'  then (a.Phour*(c.Mspeed*60*1000/e.F_108)*e.F_110) else (a.Phour*c.MSpeed*60) end) end as 速度達成率,
                                case when e.F_108 = '0' or e.F_110 = '0' or e.F_108 = '0' or e.F_102 = '0' then '0' else (case when c.MUnit = 'KG' then  (case when e.Fnumber like '12.C%' then ((a.Phour*c.MSpeed*60*1000)/(e.F_122)) else ((a.Phour*c.MSpeed*60*1000)/(e.F_122+e.F_123)) end)
                                when c.MUnit = '张'  then (b.Ohour*c.MSpeed*60*e.F_110) when c.MUnit = '箱'  then (b.Ohour*c.MSpeed*60*e.F_102) when c.MUnit = '米'  then (b.Ohour*(c.Mspeed*60*1000/e.F_108)*e.F_110) else (b.Ohour*c.MSpeed*60) end) end as 績效,case when a.POPcs = a.PPPcs then '0' else (a.POPcs-a.PPPcs)/a.POPcs end as 損耗率 ,a.PWWeight from 
                                [ChengyiYuntech].[dbo].[ScanRecord] a,[ChengyiYuntech].[dbo].[ProduceOrder] b,[ChengyiYuntech].[dbo].[machine] c,[" + sql.CYDB + "].[dbo].[ICMO] d,[" + sql.CYDB + "].[dbo].[T_ICItem] e " +
                                "where a.POID = b.ID and b.OMachineCode = c.Mcode and d.Fbillno = b.OID and d.FitemID = e.FitemID and b.OSample = '0' and a.PHour <> '0.01' and b.Odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and c.Mcode = '" + textBox1.Text + "') union " +
                                "(select c.Mname, convert(char, b.Odate, 23) as Odate, b.OOrder, b.OHour, a.Phour, a.POPcs, a.PPPcs, (case when c.MUnit = 'KG' then((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122 + e.F_123)) " +
                                "when c.MUnit = '张'  then(a.Phour * c.MSpeed * 60 * e.F_110) when c.MUnit = '箱'  then(a.Phour * c.MSpeed * 60 * e.F_102) when c.MUnit = '米'  then(a.Phour * (c.Mspeed * 60 * 1000 / e.F_108) * e.F_110) else (a.Phour * c.MSpeed * 60) end) as 速度達成率, " +
                                "(case when c.MUnit = 'KG' then((b.Ohour * c.MSpeed * 60 * 1000) / (e.F_122 + e.F_123)) when c.MUnit = '张'  then(b.Ohour * c.MSpeed * 60 * e.F_110) when c.MUnit = '箱'  then(b.Ohour * c.MSpeed * 60 * e.F_102) when c.MUnit = '米'  then(b.Ohour * (c.Mspeed * 60 * 1000 / e.F_108) * e.F_110) else (b.Ohour * c.MSpeed * 60) end) as 績效, " +
                                "case when a.POPcs = a.PPPcs then '0' else (a.POPcs - a.PPPcs) / a.POPcs end as 損耗率,a.PWWeight from[ChengyiYuntech].[dbo].[ScanRecord] a,[ChengyiYuntech].[dbo].[ProduceOrder] b,[ChengyiYuntech].[dbo].[machine] c,["+ sql.CKDB +"].[dbo].[ICMO] d,["+ sql.CKDB +"].[dbo].[T_ICItem] " +
                                "e " +
                                "where a.POID = b.ID and b.OMachineCode = c.Mcode and d.Fbillno = b.OID and d.FitemID = e.FitemID and b.OSample = '0' and b.Odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "'  and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and c.Mcode = '" + textBox1.Text + "')) a group by a.Mname, convert(char, a.Odate,23),a.OOrder) union " +
                                "(select '机台小计' as Mname,'3000-12-17' as Odate,'' as OOrder, Sum(v1.Ohour) as Ohour, Sum(v1.Phour) as Phour, Sum(v1.POPcs) as POPcs, Sum(v1.PPPcs) as PPPcs, Sum(v1.PPPcs)/SUM(v1.速度達成率) as 速度達成率, Sum(v1.PPPcs)/SUM(v1.績效) as 績效, (Sum(v1.POPcs)-Sum(v1.PPPcs))/Sum(v1.POPcs) as 損耗率,SUM(v1.PWWeight) as PWWeight from( " +
                                "(select c.Mname, b.Odate, b.OHour, a.Phour, a.POPcs, a.PPPcs,case when e.F_108 = '0' or e.F_110 = '0' or e.F_108 = '0' or e.F_102 = '0' then '0' else  " +
                                "(case when c.MUnit = 'KG' then (case when e.Fnumber like '12.C%' then (a.Phour* c.MSpeed*60*1000)/(e.F_122) else (a.Phour* c.MSpeed*60*1000)/(e.F_122+e.F_123) end) when c.MUnit = '张'  then (a.Phour* c.MSpeed*60*e.F_110) when c.MUnit = '箱'  then (a.Phour* c.MSpeed*60*e.F_102) when c.MUnit = '米'  then (a.Phour*(c.Mspeed*60*1000/e.F_108)*e.F_110) else (a.Phour* c.MSpeed*60) end) end as 速度達成率,case when e.F_108 = '0' or e.F_110 = '0' or e.F_108 = '0' or e.F_102 = '0' then '0' else  " +
                                "(case when c.MUnit = 'KG' then  (case when e.Fnumber like '12.C%' then (a.Phour* c.MSpeed*60*1000)/(e.F_122) else (a.Phour* c.MSpeed*60*1000)/(e.F_122+e.F_123) end) " +
                                "when c.MUnit = '张'  then (b.Ohour* c.MSpeed*60*e.F_110) when c.MUnit = '箱'  then (b.Ohour* c.MSpeed*60*e.F_102) when c.MUnit = '米'  then (b.Ohour*(c.Mspeed*60*1000/e.F_108)*e.F_110) else (b.Ohour* c.MSpeed*60) end) end as 績效,case when a.POPcs = a.PPPcs then '0' else (a.POPcs-a.PPPcs)/a.POPcs end as 損耗率 ,a.PWWeight from " +
                                "[ChengyiYuntech].[dbo].[ScanRecord] a,[ChengyiYuntech].[dbo].[ProduceOrder] b,[ChengyiYuntech].[dbo].[machine] c,["+ sql.CYDB +"].[dbo].[ICMO] d,["+ sql.CYDB + "].[dbo].[T_ICItem] e where a.POID = b.ID and b.OMachineCode = c.Mcode and d.Fbillno = b.OID and d.FitemID = e.FitemID and b.OSample = '0' and b.Odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "'  and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and c.Mcode = '" + textBox1.Text + "') union " +
                                "(select c.Mname, b.Odate, b.OHour, a.Phour, a.POPcs, a.PPPcs, (case when c.MUnit = 'KG' then (a.Phour* c.MSpeed*60*1000)/(e.F_122+e.F_123) when c.MUnit = '张'  then (a.Phour* c.MSpeed*60*e.F_110) when c.MUnit = '箱'  then (a.Phour* c.MSpeed*60*e.F_102) when c.MUnit = '米'  then (a.Phour*(c.Mspeed*60*1000/e.F_108)*e.F_110) else (a.Phour* c.MSpeed*60) end) as 速度達成率, " +
                                "(case when c.MUnit = 'KG' then (b.Ohour* c.MSpeed*60*1000)/(e.F_122+e.F_123) when c.MUnit = '张'  then (b.Ohour* c.MSpeed*60*e.F_110) when c.MUnit = '箱'  then (b.Ohour* c.MSpeed*60*e.F_102) when c.MUnit = '米'  then (b.Ohour*(c.Mspeed*60*1000/e.F_108)*e.F_110) else (b.Ohour* c.MSpeed*60) end) as 績效, " +
                                "case when a.POPcs = a.PPPcs then '0' else (a.POPcs-a.PPPcs)/a.POPcs end as 損耗率,a.PWWeight from[ChengyiYuntech].[dbo].[ScanRecord] a,[ChengyiYuntech].[dbo].[ProduceOrder] b,[ChengyiYuntech].[dbo].[machine] c,["+ sql.CKDB +"].[dbo].[ICMO] d,["+ sql.CKDB +"].[dbo].[T_ICItem] " +
                                "e " +
                                "where a.POID = b.ID and b.OMachineCode = c.Mcode and d.Fbillno = b.OID and d.FitemID = e.FitemID and b.OSample = '0' and a.PHour<> '0.01' and b.Odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "'  and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' and c.Mcode = '" + textBox1.Text + "')) v1 group by v1.Mname)) a order by a.Odate, a.OOrder";


                Load_Data(query);
            }
            else if (textBox1.Text == "" && checkbox1.Checked)
            {
                dataGridView1.Columns[1].Visible = false;
                dataGridView1.Columns[2].Visible = false;

                string query = @"select b.Mname,b.Mcode,isnull(a.Ohour,0) as Ohour,isnull(a.Phour,0) as Phour,isnull(a.POPcs,0) as POPcs,isnull(a.PPPcs,0) as PPPcs,isnull(a.速度達成率,0) as 速度達成率,isnull(a.績效,0) as 績效,isnull(a.損耗率,0) as 損耗率,isnull(a.PWWeight,0) as PWWeight from [ChengyiYuntech].[dbo].[machine] b left join
                                ((select v1.Mname,v1.Mcode,SUM(v1.OHour) as OHour,SUM(v1.PHour) as PHour,SUM(v1.POPcs) as POPcs,SUM(v1.PPPcs) as PPPcs,
                                SUM(v1.PPPcs)/SUM(v1.應有產量) as 速度達成率,SUM(v1.PPPcs)/SUM(v1.預計產量) as 績效,
                                case when SUM(v1.POPcs) = SUM(v1.PPPcs) then '0' else (SUM(v1.POPcs)-SUM(v1.PPPcs))/SUM(v1.POPcs) end as 損耗率,SUM(v1.PWWeight) as PWWeight from
                                ((select c.Mname,b.Odate,c.Mcode,b.OHour,a.Phour,a.POPcs,a.PPPcs,
                                case when e.F_102 = '0' or e.F_108 = '0' or e.F_110 = '0' then '0' else 
                                (case when c.MUnit = 'KG' then (case when e.Fnumber  like '12.C%' then ((a.Phour*c.MSpeed*60*1000)/(e.F_122)) else 
                                ((a.Phour*c.MSpeed*60*1000)/(e.F_122+e.F_123)) end)
                                when c.MUnit = '张'  then (a.Phour*c.MSpeed*60*e.F_110) 
                                when c.MUnit = '箱'  then (a.Phour*c.MSpeed*60*e.F_102) 
                                when c.MUnit = '米'  then (a.Phour*(c.Mspeed*60*1000/e.F_108)*e.F_110)
                                 else (a.Phour*c.MSpeed*60) end) end as 應有產量,
                                case when e.F_102 = '0' or e.F_108 = '0' or e.F_110 = '0' then '0' 
                                else (case when c.MUnit = 'KG' then  (case when e.Fnumber  like '12.C%' then ((a.Phour*c.MSpeed*60*1000)/(e.F_122)) 
                                else ((a.Phour*c.MSpeed*60*1000)/(e.F_122+e.F_123)) end)
                                when c.MUnit = '张'  then (b.Ohour*c.MSpeed*60*e.F_110) 
                                when c.MUnit = '箱'  then (b.Ohour*c.MSpeed*60*e.F_102) 
                                when c.MUnit = '米'  then (b.Ohour*(c.Mspeed*60*1000/e.F_108)*e.F_110) 
                                else (b.Ohour*c.MSpeed*60) end) end as 預計產量,a.PWWeight from 
                                [ChengyiYuntech].[dbo].[ScanRecord] a,
                                [ChengyiYuntech].[dbo].[ProduceOrder] b,
                                [ChengyiYuntech].[dbo].[machine] c,
                                ["+ sql.CYDB +"].[dbo].[ICMO] d," +
                                "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                "where a.POID = b.ID and b.OMachineCode = c.Mcode and d.Fbillno = b.OID and d.FitemID = e.FitemID and b.OSample = '0')union " +
                                "(select c.Mname,b.Odate,c.Mcode,b.OHour,a.Phour,a.POPcs,a.PPPcs, " +
                                "case when e.F_108 = '0' or e.F_110 = '0' or e.F_102 = '0' then '0' else  " +
                                "(case when c.MUnit = 'KG' then (case when e.Fnumber like '12.C%' then ((a.Phour*c.MSpeed*60*1000)/(e.F_122)) else  " +
                                "((a.Phour*c.MSpeed*60*1000)/(e.F_122+e.F_123)) end) " +
                                "when c.MUnit = '张'  then (a.Phour*c.MSpeed*60*e.F_110)  " +
                                "when c.MUnit = '箱'  then (a.Phour*c.MSpeed*60*e.F_102)  " +
                                "when c.MUnit = '米'  then (a.Phour*(c.Mspeed*60*1000/e.F_108)*e.F_110) " +
                                "else (a.Phour*c.MSpeed*60) end) end as 應有產量, " +
                                "case when e.F_108 = '0' or e.F_110 = '0' or e.F_102 = '0' then '0'  " +
                                "else (case when c.MUnit = 'KG' then  (case when e.Fnumber  like '12.C%' then ((a.Phour*c.MSpeed*60*1000)/(e.F_122))  " +
                                "else ((a.Phour*c.MSpeed*60*1000)/(e.F_122+e.F_123)) end) " +
                                "when c.MUnit = '张'  then (b.Ohour*c.MSpeed*60*e.F_110)  " +
                                "when c.MUnit = '箱'  then (b.Ohour*c.MSpeed*60*e.F_102)  " +
                                "when c.MUnit = '米'  then (b.Ohour*(c.Mspeed*60*1000/e.F_108)*e.F_110)  " +
                                "else (b.Ohour*c.MSpeed*60) end) end as 預計產量,a.PWWeight from  " +
                                "[ChengyiYuntech].[dbo].[ScanRecord] a, " +
                                "[ChengyiYuntech].[dbo].[ProduceOrder] b, " +
                                "[ChengyiYuntech].[dbo].[machine] c,  " +
                                "["+ sql.CKDB +"].[dbo].[ICMO] d, " +
                                "["+ sql.CKDB +"].[dbo].[T_ICItem] e " +
                                "where a.POID = b.ID and b.OMachineCode = c.Mcode and d.Fbillno = b.OID and d.FitemID = e.FitemID and b.OSample = '0')) v1 " +
                                "where v1.Odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "'  " +
                                "group by v1.Mname,v1.Mcode)) a on b.Mcode = a.Mcode union " +
                                "(select '机台小计' as MName, 'zz' as Mcode, SUM(a.Ohour) as OHour, SUM(a.Phour) as PHour, SUM(a.POPcs) as POPcs, SUM(a.PPPcs) as PPPcs, '' as 速度達成率, '' as 績效, '' as 損耗率, SUM(a.PWWeight) as PWWeight from " +
                                "((select v1.Mname, v1.Mcode, SUM(v1.OHour) as OHour, SUM(v1.PHour) as PHour, SUM(v1.POPcs) as POPcs, SUM(v1.PPPcs) as PPPcs, " +
                                "SUM(v1.PPPcs) / SUM(v1.應有產量) as 速度達成率, SUM(v1.PPPcs) / SUM(v1.預計產量) as 績效, " +
                                "case when SUM(v1.POPcs) = SUM(v1.PPPcs) then '0' else (SUM(v1.POPcs) - SUM(v1.PPPcs)) / SUM(v1.POPcs) end as 損耗率, SUM(v1.PWWeight) as PWWeight from " +
                                "(select c.Mname, b.Odate, c.Mcode, b.OHour, a.Phour, a.POPcs, a.PPPcs, " +
                                "case when e.F_102 = '0' or e.F_108 = '0' or e.F_110 = '0' then '0' else  " +
                                "(case when c.MUnit = 'KG' then(case when e.Fnumber  like '12.C%' then((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122)) else  " +
                                "((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122 + e.F_123)) end) " +
                                "when c.MUnit = '张'  then(a.Phour * c.MSpeed * 60 * e.F_110) " +
                                "when c.MUnit = '箱'  then(a.Phour * c.MSpeed * 60 * e.F_102) " +
                                "when c.MUnit = '米'  then(a.Phour * (c.Mspeed * 60 * 1000 / e.F_108) * e.F_110) " +
                                "else (a.Phour * c.MSpeed * 60) end) end as 應有產量, " +
                                "case when e.F_102 = '0' or e.F_108 = '0' or e.F_110 = '0' then '0'  " +
                                "else (case when c.MUnit = 'KG' then(case when e.Fnumber  like '12.C%' then((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122))  " +
                                "else ((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122 + e.F_123)) end) " +
                                "when c.MUnit = '张'  then(b.Ohour * c.MSpeed * 60 * e.F_110) " +
                                "when c.MUnit = '箱'  then(b.Ohour * c.MSpeed * 60 * e.F_102) " +
                                "when c.MUnit = '米'  then(b.Ohour * (c.Mspeed * 60 * 1000 / e.F_108) * e.F_110)  " +
                                "else (b.Ohour * c.MSpeed * 60) end) end as 預計產量, a.PWWeight from " +
                                "[ChengyiYuntech].[dbo].[ScanRecord] a, " +
                                "[ChengyiYuntech].[dbo].[ProduceOrder] b, " +
                                "[ChengyiYuntech].[dbo].[machine] c, " +
                                "["+ sql.CYDB +"].[dbo].[ICMO] d, " +
                                "["+ sql.CYDB +"].[dbo].[T_ICItem] e " +
                                "where a.POID = b.ID and b.OMachineCode = c.Mcode and d.Fbillno = b.OID and d.FitemID = e.FitemID and b.OSample = '0') v1 " +
                                "where v1.Odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "' " +
                                "group by v1.Mname,v1.Mcode)union " +
                                "(select v1.Mname, v1.Mcode, SUM(v1.OHour) as OHour, SUM(v1.PHour) as PHour, SUM(v1.POPcs) as POPcs, SUM(v1.PPPcs) as PPPcs, " +
                                "SUM(v1.PPPcs) / SUM(v1.應有產量) as 速度達成率, SUM(v1.PPPcs) / SUM(v1.預計產量) as 績效, " +
                                "case when SUM(v1.POPcs) = SUM(v1.PPPcs) then '0' else (SUM(v1.POPcs) - SUM(v1.PPPcs)) / SUM(v1.POPcs) end as 損耗率, SUM(v1.PWWeight) as PWWeight from " +
                                "(select c.Mname, b.Odate, c.Mcode, b.OHour, a.Phour, a.POPcs, a.PPPcs, " +
                                "case when e.F_108 = '0' or e.F_110 = '0' or e.F_102 = '0' then '0' else  " +
                                "(case when c.MUnit = 'KG' then(case when e.Fnumber like '12.C%' then((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122)) else  " +
                                "((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122 + e.F_123)) end) " +
                                "when c.MUnit = '张'  then(a.Phour * c.MSpeed * 60 * e.F_110) " +
                                "when c.MUnit = '箱'  then(a.Phour * c.MSpeed * 60 * e.F_102) " +
                                "when c.MUnit = '米'  then(a.Phour * (c.Mspeed * 60 * 1000 / e.F_108) * e.F_110) " +
                                "else (a.Phour * c.MSpeed * 60) end) end as 應有產量, " +
                                "case when e.F_108 = '0' or e.F_110 = '0' or e.F_102 = '0' then '0'  " +
                                "else (case when c.MUnit = 'KG' then(case when e.Fnumber  like '12.C%' then((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122))  " +
                                "else ((a.Phour * c.MSpeed * 60 * 1000) / (e.F_122 + e.F_123)) end) " +
                                "when c.MUnit = '张'  then(b.Ohour * c.MSpeed * 60 * e.F_110) " +
                                "when c.MUnit = '箱'  then(b.Ohour * c.MSpeed * 60 * e.F_102) " +
                                "when c.MUnit = '米'  then(b.Ohour * (c.Mspeed * 60 * 1000 / e.F_108) * e.F_110)  " +
                                "else (b.Ohour * c.MSpeed * 60) end) end as 預計產量,a.PWWeight from " +
                                "[ChengyiYuntech].[dbo].[ScanRecord] a, " +
                                "[ChengyiYuntech].[dbo].[ProduceOrder] b, " +
                                "[ChengyiYuntech].[dbo].[machine] c, " +
                                "["+ sql.CKDB +"].[dbo].[ICMO] d, " +
                                "["+ sql.CKDB +"].[dbo].[T_ICItem] " +
                                "e " +
                                "where a.POID = b.ID and b.OMachineCode = c.Mcode and d.Fbillno = b.OID and d.FitemID = e.FitemID and b.OSample = '0') v1 " +
                                "where v1.Odate between '" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and '" + dateTimePicker2.Value.ToString("yyyyMMdd") + "'  " +
                                "group by v1.Mname,v1.Mcode))a) order by b.Mcode";

                Load_Data(query);

                Cursor = Cursors.Default;
            }
            else if (textBox1.Text == "" && checkbox1.Checked == false)
            {
                MessageBox.Show("请填写机台编号或勾选");
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            checkbox1.Checked = false;
        }

        private void checkbox1_OnChange(object sender, EventArgs e)
        {
            if (checkbox1.Checked == true)
            {
                textBox1.Text = "";
            }
        }

        private void Load_Data(string query)
        {
            dataGridView1.Rows.Clear();

            if (textBox1.Text != "")
            {
                DataTable dt = new DataTable();
                dt = sql.getQuery(query);

                foreach (DataRow item in dt.Rows)
                {

                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = item["Mname"].ToString();
                    dataGridView1.Rows[n].Cells[1].Value = Convert.ToDateTime(item["Odate"]).ToString("yyyy/MM/dd");
                    dataGridView1.Rows[n].Cells[2].Value = item["OOrder"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["Ohour"].ToString();
                    dataGridView1.Rows[n].Cells[4].Value = item["Phour"].ToString();
                    dataGridView1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["POPcs"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["PPPcs"]).ToString("N0");
                    dataGridView1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["速度達成率"]).ToString("p");
                    dataGridView1.Rows[n].Cells[8].Value = Convert.ToDecimal(item["績效"]).ToString("p");
                    dataGridView1.Rows[n].Cells[9].Value = Convert.ToDecimal(item["損耗率"]).ToString("p");
                    dataGridView1.Rows[n].Cells[10].Value = item["PWWeight"].ToString();
                }

                dt.Clear();

                int idx = dataGridView1.Rows.Count;
                dataGridView1.Rows[idx - 1].Cells[1].Value = "";

                if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("查无信息");
                }
            }

            if (textBox1.Text == "" && checkbox1.Checked)
            {
                DataTable dt2 = new DataTable();
                dt2 = sql.getQuery(query);

                try
                {
                    foreach (DataRow item in dt2.Rows)
                    {
                        int n = dataGridView1.Rows.Add();
                        dataGridView1.Rows[n].Cells[0].Value = item["Mname"].ToString();
                        dataGridView1.Rows[n].Cells[3].Value = item["Ohour"].ToString();
                        dataGridView1.Rows[n].Cells[4].Value = item["Phour"].ToString();
                        dataGridView1.Rows[n].Cells[5].Value = Convert.ToDecimal(item["POPcs"]).ToString("N0");
                        dataGridView1.Rows[n].Cells[6].Value = Convert.ToDecimal(item["PPPcs"]).ToString("N0");
                        dataGridView1.Rows[n].Cells[7].Value = Convert.ToDecimal(item["速度達成率"]).ToString("p");
                        dataGridView1.Rows[n].Cells[8].Value = Convert.ToDecimal(item["績效"]).ToString("p");
                        dataGridView1.Rows[n].Cells[9].Value = Convert.ToDecimal(item["損耗率"]).ToString("p");
                        dataGridView1.Rows[n].Cells[10].Value = item["PWWeight"].ToString();
                    }
                    dt2.Clear();

                    if (dataGridView1.Rows.Count == 0)
                    {
                        MessageBox.Show("查无信息");
                    }
                }
                catch (Exception)
                {
                    if (dataGridView1.Rows.Count == 1)
                    {
                        MessageBox.Show("无信息汇总");
                    }
                }
            }
        }

        private void Export_Data()
        {
            if (textBox1.Text != "")
            {
                if (dataGridView1.Rows.Count != 0)
                {
                    Excel.Application excelApp;
                    Excel._Workbook wBook;
                    Excel._Worksheet wSheet;
                    Excel.Range wRange;

                    string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    string inputPath = System.Environment.CurrentDirectory;
                    string exportPath = path + @"\机台绩效明细表导出";
                    string filePath = inputPath + @"\机台绩效明细表";

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

                    wSheet.Name = "机台绩效明细表";

                    wSheet.Cells[2, 1] = "机台绩效明细表    " + dateTimePicker1.Value.ToString("yyyy/MM/dd") + " - " + dateTimePicker2.Value.ToString("yyyy/MM/dd");
                    wRange = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                    wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


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
            else if (textBox1.Text == "" && checkbox1.Checked == true)
            {
                if (dataGridView1.Rows.Count != 0)
                {
                    Excel.Application excelApp;
                    Excel._Workbook wBook;
                    Excel._Worksheet wSheet;
                    Excel.Range wRange;

                    string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    string inputPath = System.Environment.CurrentDirectory;
                    string exportPath = path + @"\机台绩效汇总表导出";
                    string filePath = inputPath + @"\机台绩效汇总表";

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

                    wSheet.Name = "机台绩效汇总表";

                    wSheet.Cells[2, 1] = "机台绩效汇总表    " + dateTimePicker1.Value.ToString("yyyy/MM/dd") + " - " + dateTimePicker2.Value.ToString("yyyy/MM/dd");
                    wRange = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[2, 1]];
                    wRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                    // storing Each row and column value to excel sheet
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        wSheet.Cells[i + 4, 1] = Convert.ToString(dataGridView1.Rows[i].Cells[0].Value);
                        wSheet.Cells[i + 4, 2] = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
                        wSheet.Cells[i + 4, 3] = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value);
                        wSheet.Cells[i + 4, 4] = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                        wSheet.Cells[i + 4, 5] = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);
                        wSheet.Cells[i + 4, 6] = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value);
                        wSheet.Cells[i + 4, 7] = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value);
                        wSheet.Cells[i + 4, 8] = Convert.ToString(dataGridView1.Rows[i].Cells[9].Value);
                        wSheet.Cells[i + 4, 9] = Convert.ToString(dataGridView1.Rows[i].Cells[10].Value);
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
        }
    }
}
