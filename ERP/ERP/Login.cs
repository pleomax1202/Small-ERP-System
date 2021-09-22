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
    public partial class Login : Form
    {
        string SID;

        public Login()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            try
            {
                Sql sql = new Sql();
                DataTable dt = new DataTable();
                dt = sql.getQuery(@"SELECT * FROM [dbo].[Staff] WHERE [SUser] = '" + bunifuMaterialTextbox1.Text + "' and [SPassword] = '" + textBox1.Text + "'");

                SID = Convert.ToString(dt.Rows[0][0]);

                if (dt.Rows.Count == 1)
                {
                    if (Convert.ToString(dt.Rows[0][4]) == "操作员")
                    {
                        bunifuTransition1.HideSync(pictureBox1);
                        this.Hide();
                        Form17 frm17 = new Form17(int.Parse(SID));
                        frm17.Show();
                    }
                    else
                    {
                        bunifuTransition1.HideSync(pictureBox1);
                        this.Hide();
                        Dashboard dashboard = new Dashboard(SID);
                        dashboard.Show();
                    }

                }
                else
                {
                    MessageBox.Show("请输入正确的登录名或密码", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bunifuMaterialTextbox1.Text = "";
                    textBox1.Text = "";
                    bunifuMaterialTextbox1.Focus();
                }

            }
            catch (System.Data.SqlClient.SqlException)
            {
                DialogResult result = MessageBox.Show("请确认是否有网络链接");
                if(result == DialogResult.OK)
                {
                    Application.Exit();
                }
            }
            catch (System.IndexOutOfRangeException)
            {
                MessageBox.Show("请输入正确的登录名或密码", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bunifuMaterialTextbox1.Text = "";
                textBox1.Text = "";
                bunifuMaterialTextbox1.Focus();
            }
        }

        private void Login_Load(object sender, EventArgs e)
        {
            this.ActiveControl = bunifuMaterialTextbox1;
        }

        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                bunifuFlatButton1.Focus();
                bunifuFlatButton1_Click(this, new EventArgs());
            }
        }

        private void bunifuMaterialTextbox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                textBox1.Focus();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
