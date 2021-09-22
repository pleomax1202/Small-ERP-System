using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Combination.Detail;
using System.Data.SqlClient;

namespace Combination
{
    public partial class 权限角色管理 : Form
    {
        public 权限角色管理()
        {
            InitializeComponent();
        }

        Sql sql = new Sql();

        private void Load_Data()
        {
            dataGridView1.Rows.Clear();

            DataTable dt = new DataTable();
            dt = sql.getQuery(@"SELECT * FROM [ChengyiYuntech].[dbo].[Staff]");

            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = item["SUser"].ToString();
                dataGridView1.Rows[n].Cells[1].Value = item["SPassword"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = item["SName"].ToString();
                dataGridView1.Rows[n].Cells[3].Value = item["SRole"].ToString();
                dataGridView1.Rows[n].Cells[4].Value = item["SID"].ToString();
            }
        }

        private void btnSeek_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" & textBox2.Text == "")
            {
                Load_Data();
            }
            else
            {
                dataGridView1.Rows.Clear();

                DataTable dt = new DataTable();
                dt = sql.getQuery(@"SELECT * FROM [ChengyiYuntech].[dbo].[Staff] WHERE [SUser] = '" + textBox1.Text + "' OR [SName] = '" + textBox2.Text + "'");

                foreach (DataRow item in dt.Rows)
                {
                    int n = dataGridView1.Rows.Add();
                    dataGridView1.Rows[n].Cells[0].Value = item["SUser"].ToString();
                    dataGridView1.Rows[n].Cells[1].Value = item["SPassword"].ToString();
                    dataGridView1.Rows[n].Cells[2].Value = item["SName"].ToString();
                    dataGridView1.Rows[n].Cells[3].Value = item["SRole"].ToString();
                    dataGridView1.Rows[n].Cells[4].Value = item["SID"].ToString();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            auNew auNew = new auNew(btnSeek);

            auNew.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                int index = dataGridView1.CurrentRow.Index;
                auEdit auEdit = new auEdit(dataGridView1, index, btnSeek);
                auEdit.Show();
            }
            else
            {
                MessageBox.Show("请选择要编辑的行");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                DialogResult result = MessageBox.Show("确定删除?", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (result == DialogResult.OK)
                {
                    int index = dataGridView1.CurrentRow.Index;

                    var SqlQuery = "";
                    SqlQuery = @"DELETE FROM [ChengyiYuntech].[dbo].[Staff] WHERE [SID] = '" + dataGridView1.Rows[index].Cells[4].Value + "'";
                    sql.sqlCmd(SqlQuery);

                    MessageBox.Show("该行已成功移除");
                    btnSeek_Click(sender, e);
                }
            }
            else
            {
                MessageBox.Show("请选择要删除的行");
            }
        }
    }
}
