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
using Combination.Detail;

namespace Combination
{
    public partial class tempBasicSetting : Form
    {
        public tempBasicSetting()
        {
            InitializeComponent();
        }

        Sql sql = new Sql();

        private void btnAdd_Click(object sender, EventArgs e)
        {
            tempBasicSetting_Add tbsa = new tempBasicSetting_Add();
            tbsa.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                int index = dataGridView1.CurrentRow.Index;
                tempBasicSetting_Edit tbse = new tempBasicSetting_Edit(dataGridView1, index);
                tbse.Show();
            }
            else
            {
                MessageBox.Show("请选择要编辑的行");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                DialogResult result = MessageBox.Show("清确认是否删除该行", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    int index = dataGridView1.CurrentRow.Index;
                    var SqlQuery = @"DELETE FROM [chengyifbsscctest].[dbo].[temperatureinfo] WHERE [ID] = '" + dataGridView1.Rows[index].Cells[6].Value + "'";
                    sql.sqlCmd(SqlQuery);

                    MessageBox.Show("该行已成功移除");
                }
            }
            else
            {
                MessageBox.Show("请选择资料行");
            }
        }

        private void 温湿度基础设置_Load(object sender, EventArgs e)
        {
            Load_Data();
        }

        private void Load_Data()
        {
            this.dataGridView1.Rows.Clear();

            DataTable dt = new DataTable();
            dt = sql.getQuery(@"select * from [chengyifbsscctest].[dbo].[temperatureinfo] ");

            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = Convert.ToString(item["machineno"]);
                dataGridView1.Rows[n].Cells[1].Value = Convert.ToString(item["area"]);
                dataGridView1.Rows[n].Cells[2].Value = Convert.ToString(item["temperature"]);
                dataGridView1.Rows[n].Cells[3].Value = Convert.ToString(item["temperaturechange"]);
                dataGridView1.Rows[n].Cells[4].Value = Convert.ToString(item["wet"]);
                dataGridView1.Rows[n].Cells[5].Value = Convert.ToString(item["wetchange"]);
                dataGridView1.Rows[n].Cells[6].Value = Convert.ToString(item["ID"]);
            }
        }
    }
}
