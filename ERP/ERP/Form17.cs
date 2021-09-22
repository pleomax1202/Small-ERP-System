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

namespace Combination
{
    public partial class Form17 : Form
    {
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        Sql sql = new Sql();

        string[] WorkDT = new string[2];
        string conStr = "Data Source = 192.168.1.252; Initial Catalog = chengyiYuntech; User ID = SA; pwd = chengyi";
        int userID;
        public Form17(int auth)
        {
            userID = auth;
            InitializeComponent();
        }

        private void Form17_Load(object sender, EventArgs e)
        {
            textBox3.ReadOnly = true;
            textBox4.ReadOnly = true;
            BtnClear_Click(this, null);
            using (SqlConnection connection = new SqlConnection(conStr))
            {
                connection.Open();
                SqlCommand command = new SqlCommand("SELECT distinct SName FROM [ChengyiYuntech].[dbo].[Staff] where SID='" + userID + "'", connection);
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    layeredTextBox1.Text = reader[0].ToString();
                }
            }

        }

        private void Form17_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void TextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ds.Reset();
                ds1.Clear();

                string temp = textBox1.Text;
                BtnClear_Click(this, null);
                textBox1.Text = temp;
                using (SqlConnection connection = new SqlConnection(conStr))
                {
                    string[] WorkDT = GetTimeSpan();
                    SqlDataAdapter adapterSample = new SqlDataAdapter("SELECT Mcode,Mname FROM [ChengyiYuntech].[dbo].[Machine] WHERE Mcode='" + textBox1.Text + "';" + 
                        "select b.MName,a.OID,a.OPName,a.OCCode,b.MInUnit,b.MoutUnit,b.MWUnit,"+
                        "'1' as 入料參數, '1' as 出料參數,'1' as 報廢參數,'1' as 檢驗參數, a.ID,a.OOrder,a.Ohour,a.OSample,'1' as 檢驗參數2 from [ChengyiYuntech].[dbo].[ProduceOrder] a,[ChengyiYuntech].[dbo].[Machine] b " +
                        "where a.OMachineCode = b.Mcode and a.OStatus = '0' and b.Mcode = '" + textBox1.Text + "' and a.OOrder = '" + WorkDT[1] + "'  and a.ODate = '" + WorkDT[0] + "' order by a.ID asc", connection);
                    adapterSample.Fill(ds);
                    
                    if (ds.Tables[0].Rows.Count != 0 )
                    {
                        textBox1.Text = ds.Tables[0].Rows[0][0].ToString();
                        textBox2.Text = ds.Tables[0].Rows[0][1].ToString();
                        if (ds.Tables[1].Rows.Count != 0)
                        {
                            if(ds.Tables[1].Rows[0][14].ToString() != "1")
                            {
                                try
                                {
                                    SqlDataAdapter adapter = new SqlDataAdapter("SELECT  Mcode,Mname FROM [ChengyiYuntech].[dbo].[Machine] WHERE Mcode='" + textBox1.Text + "';" + "select a.MName,a.Fbillno,a.FModel,a.F_111,a.MInUnit,a.MOutUnit,a.MWUnit,a.入料參數,a.出料參數,a.报废參數,a.检验参数,a.ID,a.OOrder,a.OHour,a.OSample,a.检验参数2 from " +
                                "((select d.MCode, a.ODate, a.ID, a.OOrder, a.OSample, a.OHour, d.MName, b.Fbillno, c.FModel, c.F_111, d.MInUnit, d.MOutUnit, d.MWUnit," +
                                "(case when d.MInUnit = 'KG' then  (case when c.Fnumber  like '12.C%' then 1000 / (c.F_123) else 1000 / (c.F_123 + c.F_122) end) when d.MInUnit = '张' then c.F_110 when d.MInUnit = '箱' then c.F_102 when d.MInUnit = '米' then c.F_110 * 1000 / (c.F_108) else 1 end) as 入料參數," +
                                "(case when d.MOutUnit = 'KG' then(case when c.Fnumber  like '12.C%' then 1000 / (c.F_122) else 1000 / (c.F_123 + c.F_122) end) when d.MOutUnit = '张' then c.F_110 when d.MOutUnit = '箱' then c.F_102 when d.MOutUnit = '米' then c.F_110 * 1000 / (c.F_108) else 1 end) as 出料參數," +
                                "case when c.Fnumber  like '12.C%' then(case when d.MWUnit = 'PCS' then(c.F_122) / 1000 when d.MWUnit = '张' then c.F_110 * (c.F_122) / 1000 when d.MWUnit = '箱' then c.F_102 * (c.F_122) / 1000 when d.MWUnit = '米' then c.F_110 * 1000 * (c.F_122) / (c.F_108 * 1000) else 1 end) else " +
                                "(case when d.MWUnit = 'PCS' then(c.F_123 + c.F_122) / 1000 when d.MWUnit = '张' then c.F_110 * (c.F_123 + c.F_122) / 1000 when d.MWUnit = '箱' then c.F_102 * (c.F_123 + c.F_122) / 1000 when d.MWUnit = '米' then c.F_110 * 1000 * (c.F_123 + c.F_122) / (c.F_108 * 1000) else 1 end) end as 报废參數,case when c.Fnumber  like '12.C%' then (c.F_122)/1000 else (c.F_123+c.F_122)/1000 end as 检验参数,case when c.Fnumber  like '12.C%' then (c.F_123)/1000 else (c.F_123+c.F_122)/1000 end as 检验参数2 " +
                                "from [ChengyiYuntech].[dbo].[ProduceOrder] a,(select Fbillno, FitemID from["+ sql.CYDB +"].[dbo].[ICMO]) b ,(select FItemID, F_102, F_108, F_110, F_122, F_123, FModel, F_111, Fnumber from["+ sql.CYDB +"].[dbo].[T_ICITem]) c,[ChengyiYuntech].[dbo].[Machine] " +
                                "d where a.OID = b.FbillNo and a.OStatus = '0' and a.OMachinecode = d.Mcode and b.FitemID = c.FitemID) union(select d.MCode, a.ODate, a.ID, a.OOrder, a.OSample, a.OHour, d.MName, b.Fbillno, c.FModel, c.F_111, d.MInUnit, d.MOutUnit, d.MWUnit," +
                                "(case when d.MInUnit = 'KG' then 1000/(c.F_123+c.F_122) when d.MInUnit ='张' then c.F_110 when d.MInUnit = '箱' then c.F_102 when d.MInUnit = '米' then c.F_110*1000/(c.F_108) else 1 end) as 入料參數," +
                                 "(case when d.MOutUnit ='KG' then 1000/(c.F_123+c.F_122) when d.MOutUnit ='张' then c.F_110 when d.MOutUnit = '箱' then c.F_102 when d.MOutUnit = '米' then c.F_110*1000/(c.F_108) else 1 end) as 出料參數," +
                                "(case when d.MWUnit ='PCS' then (c.F_123+c.F_122)/1000 when d.MWUnit ='张' then c.F_110*(c.F_123+c.F_122)/1000 when d.MWUnit ='箱' then c.F_102*(c.F_123+c.F_122)/1000 when d.MWUnit ='米' then c.F_110*1000*(c.F_123+c.F_122)/(c.F_108*1000) else 1 end) as 报废參數,case when c.Fnumber  like '12.C%' then (c.F_122)/1000 else (c.F_123+c.F_122)/1000 end as 检验参数,case when c.Fnumber  like '12.C%' then (c.F_123)/1000 else (c.F_123+c.F_122)/1000 end as 检验参数2 " +
                                 "from[ChengyiYuntech].[dbo].[ProduceOrder] a,(select Fbillno, FitemID from ["+ sql.CKDB +"].[dbo].[ICMO]) b ,(select FItemID, F_102, F_108, F_110, F_122, F_123, FModel, F_111, Fnumber from ["+ sql.CKDB +"].[dbo].[T_ICITem]) c,[ChengyiYuntech].[dbo].[Machine] " +
                                "d where a.OID = b.FbillNo and a.OStatus = '0' and a.OMachinecode = d.Mcode and b.FitemID = c.FitemID)) a " +
                                 "where a.Mcode = '" + textBox1.Text + "' and a.OOrder = '" + WorkDT[1] + "' and a.ODate = '" + WorkDT[0] + "' and a.OSample = '0'  order by a.ID asc", connection);
                                    ds.Reset();
                                    adapter.Fill(ds);
                                    layeredTextBox4.Visible = false;
                                    layeredTextBox5.Visible = false;
                                    layeredTextBox6.Visible = true;
                                    layeredTextBox7.Visible = false;
                                    layeredTextBox8.Visible = false;
                                    layeredTextBox10.Visible = false;
                                    label6.Visible = true;
                                    label3.Visible = false;
                                    label8.Visible = false;
                                    label12.Visible = false;
                                    label13.Visible = false;
                                    label14.Visible = false;
                                    label19.Visible = false;
                                    textBox5.Visible = true;
                                    textBox12.Visible = false;
                                }
                                catch (SqlException ex)
                                {
                                    if (ex.Number == 8134)
                                        MessageBox.Show("金蝶自定义资料未补齐!", "查询失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    else
                                        MessageBox.Show(ex.Message, "查询失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    BtnClear_Click(this, null);
                                }
                            }
                            else
                            {
                                layeredTextBox4.Visible = false;
                                layeredTextBox5.Visible = false;
                                layeredTextBox6.Visible = false;
                                layeredTextBox7.Visible = false;
                                layeredTextBox8.Visible = false;
                                layeredTextBox10.Visible = false;
                                label6.Visible = false;
                                label3.Visible = false;
                                label8.Visible = false;
                                label12.Visible = false;
                                label13.Visible = false;
                                label14.Visible = false;
                                label19.Visible = false;
                                textBox5.Visible = false;
                                textBox12.Visible = false;

                            }
                            textBox9.Text = ds.Tables[1].Rows[0][2].ToString();
                            layeredTextBox9.Text = ds.Tables[1].Rows[0][1].ToString();
                            textBox14.Text = ds.Tables[1].Rows[0][12].ToString();
                            layeredTextBox2.Text = ds.Tables[1].Rows[0][4].ToString();
                            layeredTextBox3.Text = ds.Tables[1].Rows[0][5].ToString();
                            layeredTextBox6.Text = ds.Tables[1].Rows[0][6].ToString();

                            if (textBox1.Text.StartsWith("BB"))
                            {
                                SqlDataAdapter adapter1 = new SqlDataAdapter("select TOP 1 OID, [O_001],[O_002],[O_003],[O_004],[O_005],[O_006],[O_007] FROM[ChengyiYuntech].[dbo].[ProduceOrder] where OID='" + layeredTextBox9.Text + "'", connection);
                                adapter1.Fill(ds1);
                                if (ds1.Tables[0].Rows.Count == 0) MessageBox.Show("没有找到版号！", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                else
                                {
                                    skinButtonBain.Visible = true;
                                    skinButtonBain.Enabled = true;
                                    skinButtonBain.Text = "版号 " + ds1.Tables[0].Rows[0]["O_001"].ToString() + "...";
                                    skinListBox1.Items[1].Text = "刀模编号 ";
                                    skinListBox1.Items[0].Text = "版号 " + ds1.Tables[0].Rows[0]["O_001"].ToString();
                                    skinListBox1.SelectedIndex = 0;
                                   /* MessageBox.Show(ds1.Tables[0].Rows[0]["O_001"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_002"].ToString() + "\n" +
                                        ds1.Tables[0].Rows[0]["O_003"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_004"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_005"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_006"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_007"].ToString(), "版号");
                                    */
                                }

                            }
                            else if (textBox1.Text.StartsWith("C"))
                            {
                                skinButtondao.Visible = true;
                                skinButtondao.Enabled = true;
                                skinButtondao.Text = "刀模编号 " + ds.Tables[1].Rows[0][3].ToString();
                                skinListBox1.Items[0].Text = "版号 ";
                                skinListBox1.Items[1].Text = "刀模编号 " + ds.Tables[1].Rows[0][3].ToString();
                                skinListBox1.SelectedIndex = 1;
                               // MessageBox.Show(ds.Tables[1].Rows[0][3].ToString(), "刀模编号");
                                SqlDataAdapter adapter1 = new SqlDataAdapter("select * from ((select Fmodel,Fdeleted,F_111,F_125 from ["+ sql.CYDB +"].[dbo].[T_ICITem] where FDeleted = '0' " +
                                    "and Fmodel like '%刀模%' ) union (select Fmodel, Fdeleted, F_111,F_125 from["+ sql.CKDB +"].[dbo].[T_ICITem] where FDeleted = '0'and Fmodel like '%刀模%')) a" +
                                    " where a.F_111 = '" + ds.Tables[1].Rows[0][3].ToString() + "' or a.F_125 ='" + ds.Tables[1].Rows[0][3].ToString() +"'", connection);
                                adapter1.Fill(ds1);
                                if (ds1.Tables[0].Rows.Count == 0) return;// MessageBox.Show("没有找到刀模名称！", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                else textBox8.Text = ds1.Tables[0].Rows[0][0].ToString();

                            }

                        }
                        else
                        {
                            MessageBox.Show("没有找到计划排程，请重试！", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            BtnClear_Click(this, null);
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("找不到该机台！", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        BtnClear_Click(this, null);
                    }
                }

            }

        }
        /// <summary>
        /// 控制项事件区
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)13 && e.KeyChar != (char)8 && e.KeyChar != (char)46)
            {
                MessageBox.Show("格式错误！输入字串格式不正确", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Handled = true;
            }
        }

        private void TextBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)13 && e.KeyChar != (char)8 && e.KeyChar != (char)46)
            {
                MessageBox.Show("格式错误！输入字串格式不正确", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Handled = true;
            }
        }

        private void TextBox6_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)13 && e.KeyChar != (char)8 && e.KeyChar != (char)46)
            {
                MessageBox.Show("格式错误！输入字串格式不正确", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Handled = true;
            }
        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox3.Text) && (ds.Tables.Count > 1 && ds.Tables[1].Rows.Count != 0))
            {
                layeredTextBox4.Text = (float.Parse(textBox3.Text) * float.Parse(ds.Tables[1].Rows[0][7].ToString())).ToString();
                layeredTextBox7.Text = (float.Parse(layeredTextBox4.Text) * float.Parse(ds.Tables[1].Rows[0][15].ToString())).ToString();
            }
        }

        private void TextBox4_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox4.Text) && (ds.Tables.Count > 1 && ds.Tables[1].Rows.Count != 0))
            {
                if (ds.Tables[1].Rows[0][14].ToString() != "1" && !string.IsNullOrEmpty(layeredTextBox10.Text))
                { 
                    if (float.Parse(layeredTextBox10.Text) < 0)
                    {
                        MessageBox.Show("［产出数量与报废数量总和］大于［领料数量］！", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        textBox3.Text = "";
                        layeredTextBox4.Text = "";
                        layeredTextBox7.Text = "";
                        layeredTextBox10.Text = "";
                        textBox5.Text = "";
                        textBox12.Text = "";
                    }
                }
                layeredTextBox5.Text = (float.Parse(textBox4.Text) * float.Parse(ds.Tables[1].Rows[0][8].ToString())).ToString();
                layeredTextBox8.Text = (float.Parse(layeredTextBox5.Text) * float.Parse(ds.Tables[1].Rows[0][10].ToString())).ToString();
            }
        }

        private void TextBox6_TextChanged(object sender, EventArgs e)
        {

            if (!string.IsNullOrEmpty(textBox6.Text) && ds.Tables.Count != 0 && ds.Tables[1].Rows.Count != 0)
            {
                if (float.Parse(textBox6.Text) < 0 || float.Parse(textBox6.Text) > 12)
                {
                    MessageBox.Show("时数输入范围错误！", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox6.Text = "";
                }
                else if ((float.Parse(textBox6.Text) < float.Parse(ds.Tables[1].Rows[0][13].ToString())) && ds.Tables[1].Rows[0][14].ToString() != "1")
                {
                    label18.Visible = true;
                    checkedListBox2.Visible = true;

                }
                else
                {
                    label18.Visible = false;
                    checkedListBox2.Visible = false;
                    foreach (int index in checkedListBox2.CheckedIndices)
                    {
                        checkedListBox2.SetItemChecked(index, false);
                    }
                    checkedListBox2.SelectedIndex = -1;
                }
            }
        }

        private void LayeredTextBox4_TextChanged(object sender, EventArgs e)
        {
            float inPcs, outPcs;
          
            if (!string.IsNullOrEmpty(layeredTextBox4.Text))
                inPcs = float.Parse(layeredTextBox4.Text);
            else
                inPcs = 0;
            if (!string.IsNullOrEmpty(layeredTextBox5.Text))
                outPcs = float.Parse(layeredTextBox5.Text);
            else
                outPcs = 0;
            layeredTextBox10.Text = (inPcs - outPcs).ToString();
        }

        private void LayeredTextBox5_TextChanged(object sender, EventArgs e)
        {
            float inPcs, outPcs;

            if (!string.IsNullOrEmpty(layeredTextBox4.Text))
                inPcs = float.Parse(layeredTextBox4.Text);
            else
                inPcs = 0;
            if (!string.IsNullOrEmpty(layeredTextBox5.Text))
                outPcs = float.Parse(layeredTextBox5.Text);
            else
                outPcs = 0;
            layeredTextBox10.Text = (inPcs - outPcs).ToString();
        }

        private void LayeredTextBox10_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(layeredTextBox10.Text) && ds.Tables.Count > 1 && ds.Tables[1].Rows.Count != 0)
            {
                if (float.Parse(layeredTextBox10.Text) > 0 && ds.Tables[1].Rows[0][14].ToString() != "1")
                {
                    label17.Visible = true;
                    checkedListBox1.Visible = true;
                }
                else
                {
                    label17.Visible = false;
                    checkedListBox1.Visible = false;
                    foreach (int index in checkedListBox1.CheckedIndices)
                    {
                        checkedListBox1.SetItemChecked(index, false);
                    }
                    checkedListBox1.SelectedIndex = -1;
                }
                textBox12.Text = (float.Parse(layeredTextBox10.Text) * float.Parse(ds.Tables[1].Rows[0][10].ToString())).ToString();
                textBox5.Text = (float.Parse(textBox12.Text) / float.Parse(ds.Tables[1].Rows[0][9].ToString())).ToString();

            }
        }

        private void TextBox3_MouseDown(object sender, MouseEventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox6.Text))
            {
                textBox3.ReadOnly = false;
                textBox4.ReadOnly = false;
            }
            else
            {
                MessageBox.Show("请先输入工时！", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox3.ReadOnly = true;
                textBox4.ReadOnly = true;
            }
        }

        private void TextBox4_MouseDown(object sender, MouseEventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox6.Text))
            {
                textBox3.ReadOnly = false;
                textBox4.ReadOnly = false;
            }
            else
            {
                MessageBox.Show("请先输入工时！", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox3.ReadOnly = true;
                textBox4.ReadOnly = true;
            }
        }

        private void TextBox3_Leave(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox3.Text) && (ds.Tables.Count > 1 && ds.Tables[1].Rows.Count != 0) &&
                 ds.Tables[1].Rows[0][14].ToString() != "1" && !string.IsNullOrEmpty(layeredTextBox10.Text))
            {
                if (float.Parse(layeredTextBox10.Text) < 0)
                {
                    MessageBox.Show("［产出数量与报废数量总和］大于［领料数量］！", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox3.Text = "";
                    layeredTextBox4.Text = "";
                    layeredTextBox7.Text = "";
                    layeredTextBox10.Text = "";
                    textBox5.Text = "";
                    textBox12.Text = "";
                }
            }
        }

        /// <summary>
        /// 按钮与呼叫功能区
        /// </summary>
        /// <returns></returns>
        public string[] GetTimeSpan()//判断当前时间是否在工作时间段内  
        {          
            string _strWorkingDayAM = "09:00";//09:00-23:30 限定白班輸入
            string _strWorkingDayPM = "23:30";
            string _strWorkingDayMID = "23:59";
            TimeSpan dspWorkingDayAM = DateTime.Parse(_strWorkingDayAM).TimeOfDay;
            TimeSpan dspWorkingDayPM = DateTime.Parse(_strWorkingDayPM).TimeOfDay;
            TimeSpan dspWorkingDayMID = DateTime.Parse(_strWorkingDayMID).TimeOfDay;
            DateTime t1 = Convert.ToDateTime(DateTime.Now.ToString("HH:mm"));
            TimeSpan dspNow = t1.TimeOfDay;
            if (dspNow >= dspWorkingDayAM && dspNow <= dspWorkingDayPM)
            {
                WorkDT[0] = DateTime.Now.ToString("yyyy-MM-dd");
                WorkDT[1] = "白班";
            }
            else if (dspNow <= dspWorkingDayMID && dspNow >= dspWorkingDayPM)
            {
                WorkDT[0] = DateTime.Now.ToString("yyyy-MM-dd");
                WorkDT[1] = "晚班";
            }
            else
            {
                WorkDT[0] = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                WorkDT[1] = "晚班";

            }
            return WorkDT;
        }

    
        public string GetItem(CheckedListBox Box) //取出勾选项目
        {
            string strCollected = string.Empty;
            for (int i = 0; i < Box.Items.Count; i++)
            {
                if (Box.GetItemChecked(i))
                {
                    if (strCollected == string.Empty)
                    {
                        strCollected = Box.GetItemText(Box.Items[i]);
                    }
                    else
                    {
                        strCollected = strCollected + "/" + Box.GetItemText(Box.Items[i]);
                    }
                }
            }
            return strCollected;
        }

        private void BtnSubmit_Click(object sender, EventArgs e)
        {
            TextBox3_TextChanged(this, null);
            TextBox4_TextChanged(this, null);
            TextBox6_TextChanged(this, null);
            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox3.Text) ||
                string.IsNullOrEmpty(textBox4.Text) || string.IsNullOrEmpty(textBox5.Text) || string.IsNullOrEmpty(textBox6.Text) ||
                (checkedListBox1.Visible == true && checkedListBox1.CheckedItems.Count == 0) || (checkedListBox2.Visible == true && checkedListBox2.CheckedItems.Count == 0))
            {
                MessageBox.Show("请将信息补充完整！", "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (ds.Tables[1].Rows[0][14].ToString() != "1" && !string.IsNullOrEmpty(layeredTextBox10.Text))
            {
                if(float.Parse(layeredTextBox10.Text) < 0)
                {
                    MessageBox.Show("［产出数量与报废数量总和］大于［领料数量］！", "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            using (SqlConnection connection = new SqlConnection(conStr))
            {
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                SqlTransaction transaction;
                // Start a local transaction.
                transaction = connection.BeginTransaction("Trans");
                // Must assign both transaction object and connection 
                // to Command object for a pending local transaction
                command.Connection = connection;
                command.Transaction = transaction;
                try
                {
                    command.CommandText = "INSERT INTO[chengyiYuntech].[dbo].[ScanRecord] (POID, PDate, POQty, PPQty, POPcs, PPPcs, PHour, PStaff, PNote1, PNote2, PWQty, PWWeight) " +
                    "VALUES (@POID, @PDate, @POQty, @PPQty, @POPcs, @PPPcs, @PHour, @PStaff, @PNote1, @PNote2, @PWQty, @PWWeight)";
                    command.Parameters.AddWithValue("@POID", ds.Tables[1].Rows[0][11].ToString());
                    command.Parameters.AddWithValue("@PDate", DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
                    command.Parameters.AddWithValue("@POQty", float.Parse(textBox3.Text).ToString("0.00"));
                    command.Parameters.AddWithValue("@PPQty", float.Parse(textBox4.Text).ToString("0.00"));
                    command.Parameters.AddWithValue("@POPcs", float.Parse(layeredTextBox4.Text).ToString("0.00"));
                    command.Parameters.AddWithValue("@PPPcs", float.Parse(layeredTextBox5.Text).ToString("0.00"));
                    command.Parameters.AddWithValue("@PHour", ((float.Parse(textBox6.Text)==0)? 0.01:float.Parse(textBox6.Text)).ToString("0.00"));
                    command.Parameters.AddWithValue("@PStaff", userID);
                    command.Parameters.AddWithValue("@PNote1", GetItem(checkedListBox1));
                    command.Parameters.AddWithValue("@PNote2", GetItem(checkedListBox2));
                    command.Parameters.AddWithValue("@PWQty", float.Parse(textBox5.Text).ToString("0.00"));
                    command.Parameters.AddWithValue("@PWWeight", float.Parse(textBox12.Text).ToString("0.00"));
                    command.ExecuteNonQuery();

                    command.CommandText = "UPDATE[ChengyiYuntech].[dbo].[ProduceOrder] SET OStatus = '1' " +
                            "WHERE ID = '" + ds.Tables[1].Rows[0][11].ToString() + "'";
                    command.ExecuteNonQuery();
                    // Attempt to commit the transaction.
                    transaction.Commit();
                    MessageBox.Show("提交成功！", "成功", MessageBoxButtons.OK);
                    BtnClear_Click(this, null);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("提交至数据库时发生错误!" + "Commit Exception Type:" + ex.GetType() + "Message:" + ex.Message, "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // Attempt to roll back the transaction. 
                    MessageBox.Show(ex.Message, "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    try
                    {
                        transaction.Rollback();
                    }
                    catch (Exception ex2)
                    {
                        // This catch block will handle any errors that may have occurred 
                        // on the server that would cause the rollback to fail, such as 
                        // a closed connection.
                        MessageBox.Show("复原数据库时发生错误！" + "  Rollback Exception Type:" + ex2.GetType() + "  Message:" + ex2.Message, "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        // Attempt to roll back the transaction. 
                    }
                }
            }

        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text)) System.Environment.Exit(0);
            else
            {
                DialogResult dialogResult = MessageBox.Show("确定要退出系统吗？", "提示信息", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes) System.Environment.Exit(0);
            }
        }

        public void BtnClear_Click(object sender, EventArgs e) //清空
        {
            skinButtonBain.Visible = false;
            skinButtondao.Visible = false;
            skinButtonBain.Enabled = false;
            skinButtondao.Enabled = false;
            ds.Reset();
            ds1.Clear();

            skinListBox1.SelectedIndex = -1;
            skinListBox1.Items[0].Text = "版号";
            skinListBox1.Items[1].Text = "刀模编号";
            label17.Visible = false;
            checkedListBox1.Visible = false;
            label18.Visible = false;
            checkedListBox2.Visible = false;
            foreach (Control con in Controls)
            {
                if (con is TextBox && con != textBox14 && con != layeredTextBox1)
                {
                    con.Text = string.Empty;
                }

            }
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox12.Text = "";
            layeredTextBox2.Text = "";
            layeredTextBox3.Text = "";
            layeredTextBox4.Text = "";
            layeredTextBox5.Text = "";
            layeredTextBox6.Text = "";
            layeredTextBox7.Text = "";
            layeredTextBox8.Text = "";
            layeredTextBox9.Text = "";
            layeredTextBox10.Text = "";
            foreach (int index in checkedListBox1.CheckedIndices)
            {
                checkedListBox1.SetItemChecked(index, false);
            }
            checkedListBox1.SelectedIndex = -1;
            foreach (int index in checkedListBox2.CheckedIndices)
            {
                checkedListBox2.SetItemChecked(index, false);
            }
            checkedListBox2.SelectedIndex = -1;

        }

        private void SkinButtondao_Click(object sender, EventArgs e)
        {
            MessageBox.Show(ds.Tables[1].Rows[0][3].ToString(), "刀模编号");
        }

        private void SkinButtonBain_Click(object sender, EventArgs e)
        {
            MessageBox.Show(ds1.Tables[0].Rows[0]["O_001"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_002"].ToString() + "\n" +
                                       ds1.Tables[0].Rows[0]["O_003"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_004"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_005"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_006"].ToString() + "\n" + ds1.Tables[0].Rows[0]["O_007"].ToString(), "版号");
        }

        private void CheckedListBox2_MouseEnter(object sender, EventArgs e)
        {
            checkedListBox2.Location = new Point(29, 217);
            checkedListBox2.Size = new Size(660, 400);
            label18.Location = new Point(29, 185);
            label18.Size = new Size(660, 32);

        }

        private void CheckedListBox2_MouseLeave(object sender, EventArgs e)
        {
            checkedListBox2.Location = new Point(29, 413);
            checkedListBox2.Size = new Size(482, 172);
            label18.Location = new Point(29, 383);
            label18.Size = new Size(482, 32);
        }

        
    }
}

