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
    public partial class WorkFlowBaicSetting : Form
    {
        string ID;
        Sql sql = new Sql();

        public WorkFlowBaicSetting()
        {
            InitializeComponent();
        }

        private void Load_Data()
        {
            DataTable dt = new DataTable();
            dt = sql.getQuery("SELECT * FROM[dbo].[WorkFlow]");

            foreach (DataRow item in dt.Rows)
            {
                int n = dataGridView1.Rows.Add();
                dataGridView1.Rows[n].Cells[0].Value = item["WID"].ToString();
                dataGridView1.Rows[n].Cells[1].Value = item["WName"].ToString();
                dataGridView1.Rows[n].Cells[2].Value = item["WNote"].ToString();
                dataGridView1.Rows[n].Cells[3].Value = item["ID"].ToString();
            }
        }

        private void WorkFlowBaicSetting_Load(object sender, EventArgs e)
        {
            Load_Data();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            int index = dataGridView1.CurrentRow.Index;
            WorkFlow_Edit wfe = new WorkFlow_Edit(dataGridView1, index);
            wfe.Show();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            WorkFlow_Add wfa = new WorkFlow_Add();
            wfa.Show();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                DialogResult result = MessageBox.Show("请确认是否删除该行", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    int index = dataGridView1.CurrentRow.Index;
                    ID = dataGridView1.Rows[index].Cells[3].Value.ToString();
                    var SqlQuery = @"DELETE FROM [dbo].[WorkFlow] WHERE ID = '" + ID + "'";

                    sql.sqlCmd(SqlQuery);
                    MessageBox.Show("该行已成功移除");
                    this.dataGridView1.Rows.Clear();
                    Load_Data();
                }
            }
            else
            {
                MessageBox.Show("请选择要删除的行");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Rows.Clear();
            Load_Data();
        }
    }
}
