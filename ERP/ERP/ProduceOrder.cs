using Combination.Detail;
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
using System.Runtime.InteropServices;

namespace Combination
{
    public partial class ProduceOrder : Form
    {
        string User, hour;
        int textLength;
        int count = 0;
        string value1;
        string convertIndex, convertIndex2;
        string model1, model2, model3, model4, model5, model6, model7;
        string getdata1_0, getdata1_1, getdata1_2, getdata1_3, getdata1_4, getdata1_5, getdata1_6, getdata1_7, getdata1_8, getdata1_9, getdata1_10, getdata1_11, getdata1_12, getdata1_13, getdata1_14, getdata1_15, getdata1_16;
        string MOrder, WID, ID;
        string dm1, dm2, dm3;
        string checkMachineName, checkCodeName1, checkCodeName2;

        Sql sql = new Sql();

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetWindow(IntPtr hWnd, int uCmd);
        int GW_CHILD = 5;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);
        public const int EM_SETREADONLY = 0xcf;

        private void ProduceOrder2_Load(object sender, EventArgs e)
        {
            textBox1.Focus();
            this.ActiveControl = textBox1;
            comboBox2.SelectedIndex = 0;
            DataTable dt = new DataTable();
            dt = sql.getQuery(@"SELECT[WID],[WName],[WNote] FROM[ChengyiYuntech].[dbo].[WorkFlow]");

            foreach (DataRow item in dt.Rows)
            {
                comboBox1.Items.Add(item["WName"].ToString());
            }
        }

        

        private int countI;

        public ProduceOrder(string auth, string SUser)
        {
            InitializeComponent();
            User = SUser;
            label22.Text = auth;
            label22.Visible = false;
            IntPtr editHandle = GetWindow(comboBox1.Handle, GW_CHILD);
            SendMessage(editHandle, EM_SETREADONLY, 1, 0);
        }

        private void Clear()
        {
            textBox2.Text = "";
            textBox6.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            blank1.Text = "0";
            blank2.Text = "0";
            blank3.Text = "0";
            blank4.Text = "0";
            blank5.Text = "0";
            blank6.Text = "0";
            blank7.Text = "0";
            checkBox1.Checked = false;
            comboBox2.SelectedIndex = 0;
            this.ActiveControl = comboBox2;
        }

        private void textBox6_Leave(object sender, EventArgs e)
        {
            Load_Data();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textLength > textBox1.Text.Length)
            {
                Clear();
            }
            else
            {
                DataTable dt = new DataTable();
                dt = sql.getQuery(@"select [Mcode], [Mname], [Mhour], [MOrder],right(left(Mcode,2),1) as 刀模ab判断 from [ChengyiYuntech].[dbo].[machine] where Mcode = '" + textBox1.Text + "'");

                foreach (DataRow item in dt.Rows)
                {
                    textBox4.Text = item["Mname"].ToString();
                    hour = item["Mhour"].ToString();
                    textBox2.Text = hour;
                    MOrder = item["MOrder"].ToString();
                    checkMachineName = item["刀模ab判断"].ToString();
                }

                DataTable dt2 = new DataTable();
                dt2 = sql.getQuery(@"SELECT [WName] FROM [ChengyiYuntech].[dbo].[WorkFlow] WHERE [WID] = '" + MOrder + "'");
                foreach (DataRow item in dt2.Rows)
                {
                    comboBox1.Text = item["WName"].ToString();
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = sql.getQuery("SELECT ID FROM [ChengyiYuntech].[dbo].[WorkFlow] where WName = '" + comboBox1.Text + "'");
            ID = dt.Rows[0][0].ToString();
            Load_Data();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Save_Data();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                DialogResult result = MessageBox.Show("确定删除?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (result == DialogResult.OK)
                {
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        dataGridView1.Rows.Remove(row);
                    }
                    count--;
                }
            }
            else
            {
                MessageBox.Show("请选择要删除的行");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                try
                {
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        getdata1_0 = Convert.ToDateTime(dataGridView1.Rows[i].Cells[0].Value).ToString("yyyyMMdd");
                        getdata1_1 = Convert.ToString(dataGridView1.Rows[i].Cells[1].Value);
                        getdata1_2 = Convert.ToString(dataGridView1.Rows[i].Cells[2].Value);
                        getdata1_3 = Convert.ToString(dataGridView1.Rows[i].Cells[3].Value);
                        getdata1_4 = Convert.ToString(dataGridView1.Rows[i].Cells[4].Value);
                        getdata1_5 = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);
                        getdata1_6 = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value);
                        getdata1_7 = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value);
                        getdata1_8 = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value);
                        getdata1_9 = Convert.ToString(dataGridView1.Rows[i].Cells[9].Value);
                        getdata1_10 = Convert.ToString(dataGridView1.Rows[i].Cells[10].Value);
                        getdata1_11 = Convert.ToString(dataGridView1.Rows[i].Cells[11].Value);
                        getdata1_12 = Convert.ToString(dataGridView1.Rows[i].Cells[12].Value);
                        getdata1_13 = Convert.ToString(dataGridView1.Rows[i].Cells[13].Value);
                        getdata1_14 = Convert.ToString(dataGridView1.Rows[i].Cells[14].Value);
                        getdata1_16 = Convert.ToString(dataGridView1.Rows[i].Cells[16].Value);

                        if (dataGridView1.Rows[i].Cells[15].Value.ToString() == "是")
                        {
                            getdata1_15 = "1";
                        }
                        else
                        {
                            getdata1_15 = "0";
                        }

                        DataTable dt = new DataTable();
                        dt = sql.getQuery(@"SELECT [SID],[SUser] FROM [ChengyiYuntech].[dbo].[Staff] WHERE [SID] = '" + User + "'");

                        DataTable dt2 = new DataTable();
                        dt2 = sql.getQuery(@"SELECT * FROM[ChengyiYuntech].[dbo].[WorkFlow] WHERE [WName] = '" + getdata1_16 + "'");

                        foreach (DataRow item in dt2.Rows)
                        {
                            WID = item["ID"].ToString();
                        }


                        var SqlQuery = @"INSERT INTO [chengyiYuntech].[dbo].[ProduceOrder] ([OID],[OMachineCode],[OCcode],[OPName],[OOrder],[OHour],[ODate],[OStaff],[OStatus],[O_001],[O_002],[O_003],[O_004],[O_005],[O_006],[O_007],[OSample],[OWFlow])
                                VALUES ('" + getdata1_3 + "','" + getdata1_2 + "','" + getdata1_6 + "'" +
                            ",'" + getdata1_4 + "','" + getdata1_1 + "','" + getdata1_5 + "','" + getdata1_0 + "'" +
                            ",'" + Convert.ToInt32(dt.Rows[0][0]).ToString() + "','0','" + getdata1_8 + "','" + getdata1_9 + "','" + getdata1_10 + "','" + getdata1_11 + "','" + getdata1_12 + "','" + getdata1_13 + "','" + getdata1_14 + "','" + getdata1_15 + "','" + WID + "')";

                        sql.sqlCmd(SqlQuery);
                    }

                    MessageBox.Show("已成功下达");
                    dataGridView1.Rows.Clear();
                    count = 0;
                }
                catch (Exception)
                {
                    MessageBox.Show("该制造单号已排程");
                }
                    
                
            }
            else
            {
                MessageBox.Show("请确认表格内是否有值");
            }
        }

        private void Save_Data()
        {
            value1 = textBox1.Text;
            if (textBox1.Text != "" && textBox6.Text != "" && comboBox2.Text != "")
            {
                if(checkBox1.Checked == true && value1.Substring(0, 1) == "C" || value1.Substring(0, 1) == "c")
                {
                    if (textBox6.Text != "" && textBox1.Text != "" && textBox5.Text != "" && textBox2.Text != "" && textBox4.Text != "")
                    {
                        this.dataGridView1.Rows.Add();
                        dataGridView1.Rows[count].Cells[0].Value = dateTimePicker1.Value.ToString("yyyy/MM/dd");
                        dataGridView1.Rows[count].Cells[1].Value = comboBox2.Text;
                        dataGridView1.Rows[count].Cells[2].Value = textBox1.Text;
                        dataGridView1.Rows[count].Cells[3].Value = textBox6.Text;
                        dataGridView1.Rows[count].Cells[4].Value = textBox5.Text;
                        dataGridView1.Rows[count].Cells[5].Value = textBox2.Text;
                        dataGridView1.Rows[count].Cells[6].Value = textBox3.Text;
                        dataGridView1.Rows[count].Cells[7].Value = label22.Text;
                        dataGridView1.Rows[count].Cells[8].Value = "";
                        dataGridView1.Rows[count].Cells[9].Value = "";
                        dataGridView1.Rows[count].Cells[10].Value = "";
                        dataGridView1.Rows[count].Cells[11].Value = "";
                        dataGridView1.Rows[count].Cells[12].Value = "";
                        dataGridView1.Rows[count].Cells[13].Value = "";
                        dataGridView1.Rows[count].Cells[14].Value = "";
                        dataGridView1.Rows[count].Cells[15].Value = "否";
                        dataGridView1.Rows[count].Cells[16].Value = comboBox1.Text;
                        if (checkBox1.Checked)
                        {
                            dataGridView1.Rows[count].Cells[15].Value = "是";
                        }
                        
                        count++;
                        textBox1.Text = "";
                        Clear();
                    }
                }
                else if (value1.Substring(0, 1) == "C" || value1.Substring(0, 1) == "c")
                {
                    if (textBox6.Text != "" && textBox1.Text != "" && textBox3.Text != "" && textBox5.Text != "" && textBox2.Text != "" && textBox4.Text != "")
                    {
                        this.dataGridView1.Rows.Add();
                        dataGridView1.Rows[count].Cells[0].Value = dateTimePicker1.Value.ToString("yyyy/MM/dd");
                        dataGridView1.Rows[count].Cells[1].Value = comboBox2.Text;
                        dataGridView1.Rows[count].Cells[2].Value = textBox1.Text;
                        dataGridView1.Rows[count].Cells[3].Value = textBox6.Text;
                        dataGridView1.Rows[count].Cells[4].Value = textBox5.Text;
                        dataGridView1.Rows[count].Cells[5].Value = textBox2.Text;
                        dataGridView1.Rows[count].Cells[6].Value = textBox3.Text;
                        dataGridView1.Rows[count].Cells[7].Value = label22.Text;
                        dataGridView1.Rows[count].Cells[8].Value = "";
                        dataGridView1.Rows[count].Cells[9].Value = "";
                        dataGridView1.Rows[count].Cells[10].Value = "";
                        dataGridView1.Rows[count].Cells[11].Value = "";
                        dataGridView1.Rows[count].Cells[12].Value = "";
                        dataGridView1.Rows[count].Cells[13].Value = "";
                        dataGridView1.Rows[count].Cells[14].Value = "";
                        dataGridView1.Rows[count].Cells[15].Value = "否";
                        dataGridView1.Rows[count].Cells[16].Value = comboBox1.Text;
                        if (checkBox1.Checked)
                        {
                            dataGridView1.Rows[count].Cells[15].Value = "是";
                        }
                        count++;
                        textBox1.Text = "";
                        Clear();
                    }
                    else
                    {
                        MessageBox.Show("此制造单号无刀模");
                    }
                }
                else if (value1.Substring(0, 1) == "B" || value1.Substring(0, 1) == "b")
                {
                    if (textBox6.Text != "" && textBox1.Text != "" && textBox5.Text != "" && textBox2.Text != "" && textBox4.Text != "")
                    {
                        this.dataGridView1.Rows.Add();
                        dataGridView1.Rows[count].Cells[0].Value = dateTimePicker1.Value.ToString("yyyy/MM/dd");
                        dataGridView1.Rows[count].Cells[1].Value = comboBox2.Text;
                        dataGridView1.Rows[count].Cells[2].Value = textBox1.Text;
                        dataGridView1.Rows[count].Cells[3].Value = textBox6.Text;
                        dataGridView1.Rows[count].Cells[4].Value = textBox5.Text;
                        dataGridView1.Rows[count].Cells[5].Value = textBox2.Text;
                        dataGridView1.Rows[count].Cells[6].Value = textBox3.Text;
                        dataGridView1.Rows[count].Cells[7].Value = label22.Text;
                        dataGridView1.Rows[count].Cells[8].Value = model1;
                        dataGridView1.Rows[count].Cells[9].Value = model2;
                        dataGridView1.Rows[count].Cells[10].Value = model3;
                        dataGridView1.Rows[count].Cells[11].Value = model4;
                        dataGridView1.Rows[count].Cells[12].Value = model5;
                        dataGridView1.Rows[count].Cells[13].Value = model6;
                        dataGridView1.Rows[count].Cells[14].Value = model7;
                        dataGridView1.Rows[count].Cells[15].Value = "否";
                        dataGridView1.Rows[count].Cells[16].Value = comboBox1.Text;
                        if (checkBox1.Checked)
                        {
                            dataGridView1.Rows[count].Cells[15].Value = "是";
                        }
                        count++;
                        textBox1.Text = "";
                        Clear();
                    }
                }
                else if (value1.Substring(0, 1) != "B" || value1.Substring(0, 1) != "b" || value1.Substring(0, 1) != "C" || value1.Substring(0, 1) != "c")
                {
                    if (textBox6.Text != "" && textBox1.Text != "" && textBox5.Text != "" && textBox2.Text != "" && textBox4.Text != "")
                    {
                        this.dataGridView1.Rows.Add();
                        dataGridView1.Rows[count].Cells[0].Value = dateTimePicker1.Value.ToString("yyyy/MM/dd");
                        dataGridView1.Rows[count].Cells[1].Value = comboBox2.Text;
                        dataGridView1.Rows[count].Cells[2].Value = textBox1.Text;
                        dataGridView1.Rows[count].Cells[3].Value = textBox6.Text;
                        dataGridView1.Rows[count].Cells[4].Value = textBox5.Text;
                        dataGridView1.Rows[count].Cells[5].Value = textBox2.Text;
                        dataGridView1.Rows[count].Cells[6].Value = textBox3.Text;
                        dataGridView1.Rows[count].Cells[7].Value = label22.Text;
                        dataGridView1.Rows[count].Cells[8].Value = "";
                        dataGridView1.Rows[count].Cells[9].Value = "";
                        dataGridView1.Rows[count].Cells[10].Value = "";
                        dataGridView1.Rows[count].Cells[11].Value = "";
                        dataGridView1.Rows[count].Cells[12].Value = "";
                        dataGridView1.Rows[count].Cells[13].Value = "";
                        dataGridView1.Rows[count].Cells[14].Value = "";
                        dataGridView1.Rows[count].Cells[15].Value = "否";
                        dataGridView1.Rows[count].Cells[16].Value = comboBox1.Text;
                        if (checkBox1.Checked)
                        {
                            dataGridView1.Rows[count].Cells[15].Value = "是";
                        }
                        count++;
                        textBox1.Text = "";
                        Clear();
                    }
                    else
                    {
                        MessageBox.Show("请确认是否有空值或无此笔制造单号");
                    }
                }
            }
            else
            {
                MessageBox.Show("请确认是否有空值或无此笔制造单号");
            }
        }

        private void Load_Data()
        {
            try
            {
                if (textBox6.Text != "")
                {
                    DataTable dt = new DataTable();
                    dt = sql.getQuery(@"select a.Fmodel,a.刀模1,left(right(a.刀模1,4),1) as 刀模1符合,a.刀模2,left(right(a.刀模2,4),1) as 刀模2符合,a.断张刀模,a.柔印版號1,a.柔印版號2,a.柔印版號3,a.柔印版號4,a.柔印版號5,a.柔印版號6,a.柔印版號7,a.FAuxQty from 
                                                            ((select a.FBillNo,b.Fmodel,b.F_111 as 刀模1,b.F_125 as 刀模2,b.F_181 as 断张刀模,b.F_182 as 柔印版號1,b.F_183 as 柔印版號2,b.F_184 as 柔印版號3,b.F_185 as 柔印版號4,b.F_186 as 柔印版號5,b.F_187 as 柔印版號6,b.F_188 as 柔印版號7,a.FAuxQty from ["+ sql.CYDB +"].[dbo].[ICMO] a,["+ sql.CYDB +"].[dbo].[t_ICItem] b where a.FitemID= b.FitemID) union " +
                                                            "(select a.FBillNo,b.Fmodel,b.F_111 as 刀模1,b.F_125 as 刀模2,null as 断张刀模,null as 柔印版號1,null as 柔印版號2,null as 柔印版號3,null as 柔印版號4,null as 柔印版號5,null as 柔印版號6,null as 柔印版號7,a.FAuxQty from ["+ sql.CKDB +"].[dbo].[ICMO] a,["+ sql.CKDB +"].[dbo].[t_ICItem] b where a.FitemID= b.FitemID)) a " +
                                                            "where a.FBillNo = '" + textBox6.Text + "'");

                    foreach (DataRow item in dt.Rows)
                    {
                        textBox5.Text = item["FModel"].ToString();
                        model1 = item["柔印版號1"].ToString();
                        model2 = item["柔印版號2"].ToString();
                        model3 = item["柔印版號3"].ToString();
                        model4 = item["柔印版號4"].ToString();
                        model5 = item["柔印版號5"].ToString();
                        model6 = item["柔印版號6"].ToString();
                        model7 = item["柔印版號7"].ToString();
                        dm3 = item["断张刀模"].ToString();
                        checkCodeName1 = item["刀模1符合"].ToString();
                        checkCodeName2 = item["刀模2符合"].ToString();
                    }
                }

                //
                if (textBox1.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "")
                {
                    DataTable dt2 = new DataTable();
                    dt2 = sql.getQuery(@"select * from ((select c.MUnit,c.MoutUnit,b.FModel,b.F_110,b.F_108,c.MSpeed,a.Fbillno,c.Mcode,(case when c.MUnit = 'KG' then (case when b.Fnumber like '12.C%' then (b.F_122)/(1000*c.MSpeed*60) else (b.F_123+b.F_122)/(1000*c.MSpeed*60) end) when c.MUnit = '张' then 1/(b.F_110*c.MSpeed*60) when c.MUnit = '箱' then 1/(b.F_102*c.MSpeed*60) when c.MUnit = '米' then b.F_108/(b.F_110*1000*c.MSpeed*60) else 1/(c.MSpeed*60) end) as 转换系数,case when c.MoutUnit = '米' then convert(float,b.F_108)/(convert(float,b.F_110)*1000) when c.MoutUnit = '张' then 1/convert(float,b.F_110) when c.MoutUnit = '箱' then 1/convert(float,b.F_102) else 1 end 转换系数2 
                                        from [" + sql.CYDB + "].[dbo].[ICMO] a,[" + sql.CYDB + "].[dbo].[t_ICItem] b ,[ChengyiYuntech].[dbo].[machine] c where a.FitemID= b.FitemID) union all " +
                                        "(select c.MUnit,c.MoutUnit, b.FModel, b.F_110, b.F_108, c.MSpeed, a.Fbillno, c.Mcode, (case when c.MUnit = 'KG' then(case when b.Fnumber like '12.C%' then(b.F_122) / (1000 * c.MSpeed * 60) else (b.F_123 + b.F_122) / (1000 * c.MSpeed * 60) end) when c.MUnit = '张' then 1 / (b.F_110 * c.MSpeed * 60) when c.MUnit = '箱' then 1 / (b.F_102 * c.MSpeed * 60) when c.MUnit = '米' then b.F_108 / (b.F_110 * 1000 * c.MSpeed * 60) else 1 / (c.MSpeed * 60) end) as 转换系数,case when c.MoutUnit = '米' then convert(float, b.F_108) / (convert(float, b.F_110) * 1000) when c.MoutUnit = '张' then 1 / convert(float, b.F_110) when c.MoutUnit = '箱' then 1 / convert(float, b.F_102) else 1 end 转换系数2 " + 
                                        "from["+ sql.CKDB +"].[dbo].[ICMO] a,["+ sql.CKDB +"].[dbo].[t_ICItem] b,[ChengyiYuntech].[dbo].[machine] c where a.FitemID = b.FitemID)) z  " + 
                                        "where z.FBillNo = '" + textBox6.Text + "' and z.MCode = '" + textBox1.Text + "'");

                    foreach (DataRow item in dt2.Rows)
                    {
                        convertIndex = item["转换系数"].ToString();
                        convertIndex2 = item["转换系数2"].ToString();
                        label25.Text = item["MoutUnit"].ToString();
                    }
                }

                //LabelChanged
                string value1 = textBox1.Text;
                string value2 = textBox6.Text;
                string dateTimeNow = DateTime.Now.ToString("yyyyMMdd");

                if (textBox1.Text != "" && textBox6.Text != "")
                {
                    string query = @"select * from ((select v1.*,isnull(v2.已完成,0) as 已完成,isnull(v3.已排程,0) as 已排程,case when (v1.需求数-isnull(v2.已完成,0)) < 0 then '0' else  (v1.需求数-isnull(v2.已完成,0)) end as 未完成,case when (v1.需求数-isnull(v2.已完成,0)-isnull(v3.已排程,0)) < 0 then '0' else (v1.需求数-isnull(v2.已完成,0)-isnull(v3.已排程,0)) end as 未排程 from
                                                            (select a.FbillNo,c.FName,b.Fmodel,a.FAuxQty,(case when c.FName ='KG' then (case when b.Fnumber like '12.C%' then a.FAuxQty*1000/(b.F_122) else a.FAuxQty*1000/(b.F_123+b.F_122) end) when c.FName ='张' then a.FAuxQty*b.F_110 when c.FName ='箱' then a.FAuxQty*b.F_102 when c.FName ='板' then a.FAuxQty*b.F_102 when c.FName ='米' then a.FAuxQty*b.F_110*1000/b.F_108 else a.FAuxQty end) as 需求数
                                                            from ["+ sql.CYDB +"].[dbo].[ICMO] a,["+ sql.CYDB +"].[dbo].[t_ICItem] b,["+ sql.CYDB +"].[dbo].[t_measureUnit] c where a.FitemID = b.FitemID and a.FUnitID = c.FmeasureUnitID and a.FbillNo = '" + textBox6.Text + "' ) v1 left join   " +
                                                            "(select a.OID, (case when e.MOutUnit = 'KG' then(case when c.Fnumber like '12.C%' then sum(b.PPQty) * 1000 / (c.F_122) else sum(b.PPQty) * 1000 / (c.F_123 + c.F_122) end) when e.MOutUnit = '张' then sum(b.PPQty) * c.F_110 when e.MOutUnit = '箱' then sum(b.PPQty) * c.F_102 when e.MOutUnit = '米' then sum(b.PPQty) * c.F_110 * 1000 / c.F_108 else sum(b.PPQty) end) as 已完成 " +
                                                            "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[WorkFlow] f,["+ sql.CYDB +"].[dbo].[T_Icitem] c,["+ sql.CYDB +"].[dbo].[ICMO] d,[ChengyiYuntech].[dbo].[Machine] e " +
                                                            "where a.ID = b.POID and a.OStatus = '1' and a.OID = d.Fbillno and d.FitemID = c.FitemID and a.Owflow = f.ID and a.OMachineCode = e.MCode and  f.WName = '" + comboBox1.Text + "' and a.OID = '" + textBox6.Text + "' and a.ODate <= '" + dateTimeNow + "' group by a.OID, c.F_102, c.F_110, c.F_122, c.F_123, c.F_108, e.MOutUnit, c.Fnumber) v2 on v1.FbillNo = v2.OID left join " +
                                                            "(select a.OID, (case when b.MUnit = 'KG' then (case when c.Fnumber like '12.C%' then(SUM(a.Ohour) * 1000 * b.MSpeed * 60) / (c.F_122) else (SUM(a.Ohour) * 1000 * b.MSpeed * 60) / (c.F_123 + c.F_122) end) when b.MUnit = '张' then SUM(a.Ohour)*c.F_110 * b.MSpeed * 60 when b.MUnit = '箱' then SUM(a.Ohour)*c.F_102 * b.MSpeed * 60 when b.MUnit = '米' then(SUM(a.Ohour) * b.MSpeed * 60 * c.F_110 * 1000) / c.F_108 else SUM(a.Ohour) * b.MSpeed * 60 end) as 已排程 " +
                                                            "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[Machine] b,[ChengyiYuntech].[dbo].[WorkFlow] e,["+ sql.CYDB +"].[dbo].[T_Icitem] c,["+ sql.CYDB +"].[dbo].[ICMO] " +
                                                            "d " +
                                                            "where a.OMachinecode = b.Mcode and a.OWflow = e.ID and a.OID = d.Fbillno and d.FitemID = c.FitemID and a.OStatus = '0' and e.Wname = '" + comboBox1.Text + "' and a.OID = '" + textBox6.Text + "' and a.ODate >= '" + dateTimeNow + "'  group by a.OID, c.F_102, c.F_110, c.F_122, c.F_123, c.F_108, b.MSpeed, b.MUnit, c.Fnumber) v3 on v1.FbillNo = v3.OID) union " +
                                                            "(select v1.*, isnull(v2.已完成,0) as 已完成,isnull(v3.已排程,0) as 已排程,(v1.需求数-isnull(v2.已完成,0)) as 未完成,(v1.需求数-isnull(v2.已完成,0)-isnull(v3.已排程,0)) as 未排程 from " +
                                                            "(select a.FbillNo, c.FName, b.Fmodel, a.FAuxQty, (case when c.FName = 'KG' then (case when b.Fnumber like '12.C%' then a.FAuxQty*1000/(b.F_122) else a.FAuxQty*1000/(b.F_123+b.F_122) end) when c.FName ='张' then a.FAuxQty* b.F_110 when c.FName = '箱' then a.FAuxQty* b.F_102 when c.FName = '板' then a.FAuxQty* b.F_102 when c.FName = '米' then a.FAuxQty* b.F_110*1000/b.F_108 else a.FAuxQty end) as 需求数 " +
                                                            "from["+ sql.CKDB +"].[dbo].[ICMO] a,["+ sql.CKDB +"].[dbo].[t_ICItem] b,["+ sql.CKDB +"].[dbo].[t_measureUnit] c where a.FitemID = b.FitemID and a.FUnitID = c.FmeasureUnitID and a.FbillNo = '" + textBox6.Text + "' ) v1 left join " +
                                                            "(select a.OID, (case when e.MOutUnit = 'KG' then (case when c.Fnumber like '12.C%' then sum(b.PPQty) *1000/(c.F_122) else sum(b.PPQty) *1000/(c.F_123+c.F_122) end) when e.MOutUnit ='张' then sum(b.PPQty) * c.F_110 when e.MOutUnit = '箱' then sum(b.PPQty) * c.F_102 when e.MOutUnit = '米' then sum(b.PPQty) * c.F_110*1000/c.F_108 else sum(b.PPQty) end) as 已完成 " +
                                                            "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b,[ChengyiYuntech].[dbo].[WorkFlow] f,["+ sql.CKDB +"].[dbo].[T_Icitem] c,["+ sql.CKDB +"].[dbo].[ICMO] d,[ChengyiYuntech].[dbo].[Machine] " +
                                                            "e " +
                                                            "where a.ID = b.POID and a.OStatus = '1' and a.OID = d.Fbillno and d.FitemID = c.FitemID and a.Owflow = f.ID and a.OMachineCode = e.MCode  and f.Wname = '" + comboBox1.Text + "' and a.OID = '" + textBox6.Text + "' and a.ODate <= '" + dateTimeNow + "' group by a.OID, c.F_102, c.F_110, c.F_122, c.F_123, c.F_108, e.MOutUnit, c.Fnumber) v2 on v1.FbillNo = v2.OID left join " +
                                                            "(select a.OID, (case when b.MUnit = 'KG' then (case when c.Fnumber like '12.C%' then (SUM(a.Ohour) *1000*b.MSpeed*60)/(c.F_122) else (SUM(a.Ohour) *1000*b.MSpeed*60)/(c.F_123+c.F_122) end)when b.MUnit = '张' then SUM(a.Ohour) * c.F_110* b.MSpeed*60 when b.MUnit = '箱' then SUM(a.Ohour) * c.F_102* b.MSpeed*60 when b.MUnit = '米' then (SUM(a.Ohour) * b.MSpeed*60*c.F_110*1000)/c.F_108 else SUM(a.Ohour)*b.MSpeed*60 end) as 已排程 " +
                                                            "from[ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[Machine] b,[ChengyiYuntech].[dbo].[WorkFlow] e,["+ sql.CKDB +"].[dbo].[T_Icitem] c,["+ sql.CKDB +"].[dbo].[ICMO] " +
                                                            "d " +
                                                            "where a.OMachinecode = b.Mcode and a.OWflow = e.ID and a.OID = d.Fbillno and d.FitemID = c.FitemID and a.OStatus = '0' and e.Wname = '" + comboBox1.Text + "' and a.OID = '" + textBox6.Text + "' and a.ODate >= '" + dateTimeNow + "'  group by a.OID, c.F_102, c.F_110, c.F_122, c.F_123, c.F_108, b.MSpeed, b.MUnit, c.Fnumber) v3 on v1.FbillNo = v3.OID)) a";

                    DataTable dt3 = new DataTable();
                    dt3 = sql.getQuery(query);

                    foreach (DataRow item in dt3.Rows)
                    {
                        blank1.Text = Convert.ToDecimal(item["需求数"]).ToString("0");
                        blank2.Text = Convert.ToDecimal(item["已完成"]).ToString("0");
                        blank3.Text = Convert.ToDecimal(item["已排程"]).ToString("0");
                        blank6.Text = Convert.ToDecimal(item["未完成"]).ToString("0");
                    }
                }

                if (textBox1.Text != "" && textBox6.Text != "")
                {
                    blank7.Text = blank1.Text;
                    DataTable dt = new DataTable();
                    dt = sql.getQuery(@"select a.OID,OWFlow as OMCode,SUM(isnull(b.PPPcs,0)) as PPPcs 
                                                            from [ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[ScanRecord] b
                                                            where a.ID = b.POID and a.OID = '" + textBox6.Text + "'" +
                                                            "group by a.OID,OWFlow");

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (ID == dt.Rows[i][1].ToString())
                        {
                            countI = i;
                            if (countI != 0)
                            {
                                countI--;
                                blank7.Text = dt.Rows[countI][2].ToString();
                            }
                            else
                            {
                                blank7.Text = blank1.Text;
                            }
                        }
                    }

                }

                if (textBox6.Text != "" && textBox1.Text != "" && comboBox2.Text != "")
                {
                    if (value1.Substring(0, 1) == "C" || value1.Substring(0, 1) == "c")
                    {
                        DateTime datetimeNow = DateTime.Now;
                        DataTable dt4 = new DataTable();
                        dt4 = sql.getQuery(@"select b.Fmodel,b.F_111 as 刀模1,F_125 as 刀模2,a.FAuxQty from ["+ sql.CYDB +"].[dbo].[ICMO] a " +
                                            ",["+ sql.CYDB +"].[dbo].[t_ICItem] b  where a.FitemID= b.FitemID and FBillNo = '" + textBox6.Text + "'");

                        foreach (DataRow item in dt4.Rows)
                        {
                            dm1 = item["刀模1"].ToString();
                            dm2 = item["刀模2"].ToString();

                            if (checkCodeName1 == checkMachineName || (checkMachineName == "C" && checkCodeName1 == "B" ) || (checkMachineName == "D" && checkCodeName1 == "B"))
                            {
                                textBox3.Text = dm1;
                            }
                            else
                            {
                                if (checkCodeName2 == checkMachineName || (checkMachineName == "C" && checkCodeName2 == "B") || (checkMachineName == "D" && checkCodeName1 == "B"))
                                {
                                    textBox3.Text = dm2;
                                }
                            }
                            if (comboBox1.Text == "断张")
                            {
                                textBox3.Text = dm3;
                            }
                        }
                        DataTable dt4Check = new DataTable();
                        dt4Check = sql.getQuery(@"select OMachineCode,OPName, OCcode as 刀模1, OID,* from [ChengyiYuntech].[dbo].[ProduceOrder] where ODate ='" + dateTimePicker1.Value.ToString("yyyyMMdd") + "' and OOrder = '" + comboBox2.Text + "' and OCcode = '" + dm1 + "'");

                        foreach (DataRow item2 in dt4Check.Rows)
                        {
                            string dmc1 = item2["刀模1"].ToString();
                            if (dmc1 != "")
                            {
                                if (textBox1.Text.ToUpper() == item2["OMachineCode"].ToString() && textBox6.Text != item2["OID"].ToString() || textBox1.Text.ToLower() == item2["OMachineCode"].ToString() && textBox6.Text != item2["OID"].ToString())
                                {
                                    if (checkCodeName1 == checkMachineName || (checkMachineName == "C" && checkCodeName1 == "B") || (checkMachineName == "D" && checkCodeName1 == "B"))
                                    {
                                        textBox3.Text = dm1;
                                    }
                                }
                                else
                                {
                                    if (checkCodeName2 == checkMachineName || (checkMachineName == "C" && checkCodeName2 == "B") || (checkMachineName == "D" && checkCodeName1 == "B"))
                                    {
                                        textBox3.Text = dm2;
                                    }
                                }
                                
                            }
                            if (comboBox1.Text == "断张")
                            {
                                textBox3.Text = dm3;
                            }
                        }
                        if (textBox3.Text == "")
                        {
                            MessageBox.Show("刀模已被使用");
                        }
                    }
                }

                try
                {
                    if (value1.Substring(0, 1) != "C" && value1.Substring(0, 1) != "c")
                    {
                        textBox3.Text = "";
                    }
                }
                catch (Exception)
                {

                }



                //Sub
                if (blank7.Text != "__________")
                {
                    float Sub = float.Parse(blank7.Text) - float.Parse(blank2.Text);
                    blank6.Text = Sub.ToString("0");
                    float Sub2 = float.Parse(blank7.Text) - float.Parse(blank2.Text) - float.Parse(blank3.Text);
                    blank4.Text = Sub2.ToString("0");
                    //mutiple
                    float multi = float.Parse(blank4.Text) * float.Parse(convertIndex);
                    blank5.Text = multi.ToString("0.0");
                    float value;
                    value = float.Parse(textBox2.Text) / float.Parse(convertIndex) * float.Parse(convertIndex2);
                    lblOrderQty.Text = value.ToString("0");
                }
            }
            catch (System.Data.SqlClient.SqlException)
            {

            }
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text != "")
            {
                this.ActiveControl = textBox1;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox6.Text != "")
            {
                Load_Data();
            }
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && textBox6.Text != "")
            {
                Load_Data();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == 0)
            {
                textBox2.Text = hour;
            }
            else
            {
                textBox2.Text = Convert.ToString(float.Parse(hour) + 0.5);
            }

        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            try
            {
                if (float.Parse(textBox2.Text) > float.Parse(blank5.Text))
                {
                    DialogResult result = MessageBox.Show("输入工时已大于剩余工时", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (result == DialogResult.OK)
                    {

                    }
                }
            }
            catch (Exception)
            {

            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = textBox6;
            }
        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.Focus();
            }
        }


        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = textBox2;
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar >= 0 && e.KeyChar <= 126)
            {
                textLength = textBox1.Text.Length;
            }
        }

        private void comboBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ActiveControl = textBox1;
            }
        }
    }
}
