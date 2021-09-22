using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data.SqlClient;

namespace Combination
{
    public partial class Dashboard : Form
    {
        string LoginSID, SID, sRole;
        Sql sql = new Sql();
        public Dashboard(string auth)
        {
            InitializeComponent();
            LoginSID = auth;
            this.IsMdiContainer = true;

            DataTable dt = new DataTable();
            dt = sql.getQuery(@"SELECT * FROM [dbo].[Staff] WHERE [SID] = '" + auth + "'");

            this.Text = "宁波诚毅纸业有限公司    登录人：" + Convert.ToString(dt.Rows[0][1]);
            SID = Convert.ToString(dt.Rows[0][0]);
            sRole = Convert.ToString(dt.Rows[0][4]);
            ShowBtn();
        }


        private void Dashboard_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void btnProduceOrder_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("生产排程", new ProduceOrder(this.Text.Substring(18), LoginSID));
        }

        private void btnOrder_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("排程报表", new 排程报表(sRole));
        }

        private void btnMachineDaily_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("机台日报表", new 机台日报表());
        }

        private void btnMissionTot_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("任务单汇总表", new 任务单流程表());
        }

        private void btnFactoryRecord_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("车间输入记录", new FactoryInput(SID, sRole));
        }

        private void btnWorkFlow_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("生产工序信息", new WorkFlowBaicSetting());
        }

        private void btnMachineBasic_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("机台基础信息", new 机台基础信习查询());
        }

        private void btnKnifeModel_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("刀模报表", new 刀模报表());
        }

        private void btnMEfficient_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("机台绩效报表", new 机台绩效报表());
        }

        private void btnStorageCheck_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("入库检核表", new 入庫檢核表());
        }

        private void btnBoxLabel_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("外箱及标签查询", new CheckExternalBoxAndLabel());
        }

        private void btnOrderCheck_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("发货通知检核表", new 发货通知检核表());
        }

        private void PdNameDefInfo_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("品名与自定义", new PNameDefInfo());
        }

        private void temperatureRP_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("温控报表", new 温控报表());
        }

        private void btnMErrorCheck_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("串料检核表", new MErrorCheck());
        }

        private void btnAuth_Click(object sender, EventArgs e)
        {
            this.Add_TabPage("权限角色管理", new 权限角色管理());
        }

        private void btnBoxDemand_Click_1(object sender, EventArgs e)
        {
            this.Add_TabPage("外箱需求料表", new BoxDemand());
        }

        public void Add_TabPage(string str, Form myForm) //将标题添加进tabpage中
        {

            if (!this.tabControlCheckHave(this.MainTabControl, str))
            {
                this.MainTabControl.TabPages.Add(str);
                this.MainTabControl.SelectTab((int)(this.MainTabControl.TabPages.Count - 1));
                myForm.FormBorderStyle = FormBorderStyle.None;
                myForm.TopLevel = false;
                myForm.Dock = DockStyle.Fill;
                myForm.Show();
                myForm.Parent = this.MainTabControl.SelectedTab;
            }
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {
            
        }

        private void ShowBtn()
        {
            if (sRole == "系统管理员")
            {
                ListBtn();
                //Remove
                flowLayoutPanel1.Controls.Remove(btnAuth);
            }
            else if (sRole == "权限管理员")
            {
                ListBtn();
                //Remove
                flowLayoutPanel1.Controls.Remove(btnMachineBasic);
                flowLayoutPanel1.Controls.Remove(btnWorkFlow);
            }
            else if (sRole == "品检" || sRole == "干部")
            {
                ListBtn();
                //Remove
                flowLayoutPanel1.Controls.Remove(btnMachineBasic);
                flowLayoutPanel1.Controls.Remove(btnWorkFlow);
                flowLayoutPanel1.Controls.Remove(btnAuth);
            }
            else if (sRole == "验厂")
            {
                ListBtn();
                flowLayoutPanel1.Controls.Remove(btnProduceOrder);
                flowLayoutPanel1.Controls.Remove(btnOrder);
                flowLayoutPanel1.Controls.Remove(btnMachineDaily);
                flowLayoutPanel1.Controls.Remove(btnMEfficient);
                flowLayoutPanel1.Controls.Remove(btnBoxDemand);
                flowLayoutPanel1.Controls.Remove(btnBoxLabel);
                flowLayoutPanel1.Controls.Remove(btnStorageCheck);
                flowLayoutPanel1.Controls.Remove(temperatureRP);
                flowLayoutPanel1.Controls.Remove(btnKnifeModel);
                flowLayoutPanel1.Controls.Remove(PdNameDefInfo);
                flowLayoutPanel1.Controls.Remove(btnOrderCheck);
                flowLayoutPanel1.Controls.Remove(btnMErrorCheck);
                flowLayoutPanel1.Controls.Remove(btnMachineBasic);
                flowLayoutPanel1.Controls.Remove(btnWorkFlow);
                flowLayoutPanel1.Controls.Remove(btnAuth);
            }
            else if (sRole == "仓库") //btnStorageCheck, btnBoxLabel, btnBoxDemand
            {
                ListBtn();
                //Remove
                flowLayoutPanel1.Controls.Remove(btnProduceOrder);
                flowLayoutPanel1.Controls.Remove(btnFactoryRecord);
                flowLayoutPanel1.Controls.Remove(btnMachineDaily);
                flowLayoutPanel1.Controls.Remove(btnMEfficient);
                flowLayoutPanel1.Controls.Remove(btnMissionTot);
                flowLayoutPanel1.Controls.Remove(temperatureRP);
                flowLayoutPanel1.Controls.Remove(btnKnifeModel);
                flowLayoutPanel1.Controls.Remove(PdNameDefInfo);
                flowLayoutPanel1.Controls.Remove(btnMErrorCheck);
                flowLayoutPanel1.Controls.Remove(btnMachineBasic);
                flowLayoutPanel1.Controls.Remove(btnWorkFlow);
                flowLayoutPanel1.Controls.Remove(btnAuth);
            }
            else if (sRole == "最高管理员")
            {
                ListBtn();
            }
        }

        private void ListBtn()
        {
            flowLayoutPanel1.Controls.SetChildIndex(btnProduceOrder, 1);
            flowLayoutPanel1.Controls.SetChildIndex(btnOrder, 2);
            flowLayoutPanel1.Controls.SetChildIndex(btnFactoryRecord, 3);
            flowLayoutPanel1.Controls.SetChildIndex(btnMachineDaily, 4);
            flowLayoutPanel1.Controls.SetChildIndex(btnMEfficient, 5);
            flowLayoutPanel1.Controls.SetChildIndex(btnMissionTot, 6);
            flowLayoutPanel1.Controls.SetChildIndex(temperatureRP, 7);
            flowLayoutPanel1.Controls.SetChildIndex(btnKnifeModel, 8);
            flowLayoutPanel1.Controls.SetChildIndex(PdNameDefInfo, 9);
            flowLayoutPanel1.Controls.SetChildIndex(btnOrderCheck, 10);
            flowLayoutPanel1.Controls.SetChildIndex(btnStorageCheck, 11);
            flowLayoutPanel1.Controls.SetChildIndex(btnBoxLabel, 12);
            flowLayoutPanel1.Controls.SetChildIndex(btnBoxDemand, 13);
            flowLayoutPanel1.Controls.SetChildIndex(btnMErrorCheck, 14);
            flowLayoutPanel1.Controls.SetChildIndex(btnAuth, 15);
            flowLayoutPanel1.Controls.SetChildIndex(btnMachineBasic, 16);
            flowLayoutPanel1.Controls.SetChildIndex(btnWorkFlow, 17);
        }

        public bool tabControlCheckHave(TabControl tab, string tabName) //看tabpage中是否已有窗体
        {
            for (int i = 0; i < tab.TabCount; i++)
            {
                if (tab.TabPages[i].Text == tabName)
                {
                    tab.SelectedIndex = i;
                    return true;
                }
            }
            return false;
        }
    }
}
