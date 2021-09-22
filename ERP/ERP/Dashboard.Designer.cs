namespace Combination
{
    partial class Dashboard
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Dashboard));
            BunifuAnimatorNS.Animation animation1 = new BunifuAnimatorNS.Animation();
            BunifuAnimatorNS.Animation animation2 = new BunifuAnimatorNS.Animation();
            this.panel1 = new System.Windows.Forms.Panel();
            this.bunifuGradientPanel1 = new Bunifu.Framework.UI.BunifuGradientPanel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnBoxLabel = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnKnifeModel = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnAuth = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnMachineBasic = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnFactoryRecord = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnMissionTot = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnMachineDaily = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnOrder = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnProduceOrder = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnMEfficient = new Bunifu.Framework.UI.BunifuFlatButton();
            this.temperatureRP = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnOrderCheck = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnStorageCheck = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnWorkFlow = new Bunifu.Framework.UI.BunifuFlatButton();
            this.PdNameDefInfo = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnMErrorCheck = new Bunifu.Framework.UI.BunifuFlatButton();
            this.btnBoxDemand = new Bunifu.Framework.UI.BunifuFlatButton();
            this.bunifuTransition1 = new BunifuAnimatorNS.BunifuTransition(this.components);
            this.MainTabControl = new System.Windows.Forms.TabControl();
            this.PanelTransition = new BunifuAnimatorNS.BunifuTransition(this.components);
            this.panel1.SuspendLayout();
            this.bunifuGradientPanel1.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.bunifuGradientPanel1);
            this.PanelTransition.SetDecoration(this.panel1, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.panel1, BunifuAnimatorNS.DecorationType.None);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(230, 1033);
            this.panel1.TabIndex = 0;
            // 
            // bunifuGradientPanel1
            // 
            this.bunifuGradientPanel1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("bunifuGradientPanel1.BackgroundImage")));
            this.bunifuGradientPanel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bunifuGradientPanel1.Controls.Add(this.flowLayoutPanel1);
            this.PanelTransition.SetDecoration(this.bunifuGradientPanel1, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.bunifuGradientPanel1, BunifuAnimatorNS.DecorationType.None);
            this.bunifuGradientPanel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.bunifuGradientPanel1.GradientBottomLeft = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.bunifuGradientPanel1.GradientBottomRight = System.Drawing.Color.White;
            this.bunifuGradientPanel1.GradientTopLeft = System.Drawing.Color.White;
            this.bunifuGradientPanel1.GradientTopRight = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.bunifuGradientPanel1.Location = new System.Drawing.Point(0, 0);
            this.bunifuGradientPanel1.Name = "bunifuGradientPanel1";
            this.bunifuGradientPanel1.Quality = 10;
            this.bunifuGradientPanel1.Size = new System.Drawing.Size(230, 1033);
            this.bunifuGradientPanel1.TabIndex = 0;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.BackColor = System.Drawing.Color.Transparent;
            this.flowLayoutPanel1.Controls.Add(this.pictureBox1);
            this.flowLayoutPanel1.Controls.Add(this.btnBoxLabel);
            this.flowLayoutPanel1.Controls.Add(this.btnKnifeModel);
            this.flowLayoutPanel1.Controls.Add(this.btnAuth);
            this.flowLayoutPanel1.Controls.Add(this.btnMachineBasic);
            this.flowLayoutPanel1.Controls.Add(this.btnFactoryRecord);
            this.flowLayoutPanel1.Controls.Add(this.btnMissionTot);
            this.flowLayoutPanel1.Controls.Add(this.btnMachineDaily);
            this.flowLayoutPanel1.Controls.Add(this.btnOrder);
            this.flowLayoutPanel1.Controls.Add(this.btnProduceOrder);
            this.flowLayoutPanel1.Controls.Add(this.btnMEfficient);
            this.flowLayoutPanel1.Controls.Add(this.temperatureRP);
            this.flowLayoutPanel1.Controls.Add(this.btnOrderCheck);
            this.flowLayoutPanel1.Controls.Add(this.btnStorageCheck);
            this.flowLayoutPanel1.Controls.Add(this.btnWorkFlow);
            this.flowLayoutPanel1.Controls.Add(this.PdNameDefInfo);
            this.flowLayoutPanel1.Controls.Add(this.btnMErrorCheck);
            this.flowLayoutPanel1.Controls.Add(this.btnBoxDemand);
            this.PanelTransition.SetDecoration(this.flowLayoutPanel1, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.flowLayoutPanel1, BunifuAnimatorNS.DecorationType.None);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(230, 1033);
            this.flowLayoutPanel1.TabIndex = 12;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.bunifuTransition1.SetDecoration(this.pictureBox1, BunifuAnimatorNS.DecorationType.None);
            this.PanelTransition.SetDecoration(this.pictureBox1, BunifuAnimatorNS.DecorationType.None);
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(3, 10);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(3, 10, 3, 20);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(224, 146);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // btnBoxLabel
            // 
            this.btnBoxLabel.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnBoxLabel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnBoxLabel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBoxLabel.BorderRadius = 0;
            this.btnBoxLabel.ButtonText = "外箱及标签";
            this.btnBoxLabel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnBoxLabel, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnBoxLabel, BunifuAnimatorNS.DecorationType.None);
            this.btnBoxLabel.DisabledColor = System.Drawing.Color.Gray;
            this.btnBoxLabel.Iconcolor = System.Drawing.Color.Transparent;
            this.btnBoxLabel.Iconimage = null;
            this.btnBoxLabel.Iconimage_right = null;
            this.btnBoxLabel.Iconimage_right_Selected = null;
            this.btnBoxLabel.Iconimage_Selected = null;
            this.btnBoxLabel.IconMarginLeft = 0;
            this.btnBoxLabel.IconMarginRight = 0;
            this.btnBoxLabel.IconRightVisible = true;
            this.btnBoxLabel.IconRightZoom = 0D;
            this.btnBoxLabel.IconVisible = true;
            this.btnBoxLabel.IconZoom = 90D;
            this.btnBoxLabel.IsTab = false;
            this.btnBoxLabel.Location = new System.Drawing.Point(0, 176);
            this.btnBoxLabel.Margin = new System.Windows.Forms.Padding(0);
            this.btnBoxLabel.Name = "btnBoxLabel";
            this.btnBoxLabel.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnBoxLabel.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnBoxLabel.OnHoverTextColor = System.Drawing.Color.White;
            this.btnBoxLabel.selected = false;
            this.btnBoxLabel.Size = new System.Drawing.Size(230, 44);
            this.btnBoxLabel.TabIndex = 45;
            this.btnBoxLabel.Text = "外箱及标签";
            this.btnBoxLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnBoxLabel.Textcolor = System.Drawing.Color.Black;
            this.btnBoxLabel.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBoxLabel.Click += new System.EventHandler(this.btnBoxLabel_Click);
            // 
            // btnKnifeModel
            // 
            this.btnKnifeModel.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnKnifeModel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnKnifeModel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnKnifeModel.BorderRadius = 0;
            this.btnKnifeModel.ButtonText = "刀模版号报表";
            this.btnKnifeModel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnKnifeModel, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnKnifeModel, BunifuAnimatorNS.DecorationType.None);
            this.btnKnifeModel.DisabledColor = System.Drawing.Color.Gray;
            this.btnKnifeModel.Iconcolor = System.Drawing.Color.Transparent;
            this.btnKnifeModel.Iconimage = null;
            this.btnKnifeModel.Iconimage_right = null;
            this.btnKnifeModel.Iconimage_right_Selected = null;
            this.btnKnifeModel.Iconimage_Selected = null;
            this.btnKnifeModel.IconMarginLeft = 0;
            this.btnKnifeModel.IconMarginRight = 0;
            this.btnKnifeModel.IconRightVisible = true;
            this.btnKnifeModel.IconRightZoom = 0D;
            this.btnKnifeModel.IconVisible = true;
            this.btnKnifeModel.IconZoom = 90D;
            this.btnKnifeModel.IsTab = false;
            this.btnKnifeModel.Location = new System.Drawing.Point(0, 220);
            this.btnKnifeModel.Margin = new System.Windows.Forms.Padding(0);
            this.btnKnifeModel.Name = "btnKnifeModel";
            this.btnKnifeModel.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnKnifeModel.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnKnifeModel.OnHoverTextColor = System.Drawing.Color.White;
            this.btnKnifeModel.selected = false;
            this.btnKnifeModel.Size = new System.Drawing.Size(230, 44);
            this.btnKnifeModel.TabIndex = 39;
            this.btnKnifeModel.Text = "刀模版号报表";
            this.btnKnifeModel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnKnifeModel.Textcolor = System.Drawing.Color.Black;
            this.btnKnifeModel.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnKnifeModel.Click += new System.EventHandler(this.btnKnifeModel_Click);
            // 
            // btnAuth
            // 
            this.btnAuth.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnAuth.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnAuth.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAuth.BorderRadius = 0;
            this.btnAuth.ButtonText = "权限角色管理";
            this.btnAuth.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnAuth, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnAuth, BunifuAnimatorNS.DecorationType.None);
            this.btnAuth.DisabledColor = System.Drawing.Color.Gray;
            this.btnAuth.Iconcolor = System.Drawing.Color.Transparent;
            this.btnAuth.Iconimage = null;
            this.btnAuth.Iconimage_right = null;
            this.btnAuth.Iconimage_right_Selected = null;
            this.btnAuth.Iconimage_Selected = null;
            this.btnAuth.IconMarginLeft = 0;
            this.btnAuth.IconMarginRight = 0;
            this.btnAuth.IconRightVisible = true;
            this.btnAuth.IconRightZoom = 0D;
            this.btnAuth.IconVisible = true;
            this.btnAuth.IconZoom = 90D;
            this.btnAuth.IsTab = false;
            this.btnAuth.Location = new System.Drawing.Point(0, 264);
            this.btnAuth.Margin = new System.Windows.Forms.Padding(0);
            this.btnAuth.Name = "btnAuth";
            this.btnAuth.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnAuth.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnAuth.OnHoverTextColor = System.Drawing.Color.White;
            this.btnAuth.selected = false;
            this.btnAuth.Size = new System.Drawing.Size(230, 44);
            this.btnAuth.TabIndex = 38;
            this.btnAuth.Text = "权限角色管理";
            this.btnAuth.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnAuth.Textcolor = System.Drawing.Color.Black;
            this.btnAuth.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAuth.Click += new System.EventHandler(this.btnAuth_Click);
            // 
            // btnMachineBasic
            // 
            this.btnMachineBasic.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnMachineBasic.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMachineBasic.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnMachineBasic.BorderRadius = 0;
            this.btnMachineBasic.ButtonText = "机台基础信息";
            this.btnMachineBasic.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnMachineBasic, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnMachineBasic, BunifuAnimatorNS.DecorationType.None);
            this.btnMachineBasic.DisabledColor = System.Drawing.Color.Gray;
            this.btnMachineBasic.Iconcolor = System.Drawing.Color.Transparent;
            this.btnMachineBasic.Iconimage = null;
            this.btnMachineBasic.Iconimage_right = null;
            this.btnMachineBasic.Iconimage_right_Selected = null;
            this.btnMachineBasic.Iconimage_Selected = null;
            this.btnMachineBasic.IconMarginLeft = 0;
            this.btnMachineBasic.IconMarginRight = 0;
            this.btnMachineBasic.IconRightVisible = true;
            this.btnMachineBasic.IconRightZoom = 0D;
            this.btnMachineBasic.IconVisible = true;
            this.btnMachineBasic.IconZoom = 90D;
            this.btnMachineBasic.IsTab = false;
            this.btnMachineBasic.Location = new System.Drawing.Point(0, 308);
            this.btnMachineBasic.Margin = new System.Windows.Forms.Padding(0);
            this.btnMachineBasic.Name = "btnMachineBasic";
            this.btnMachineBasic.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMachineBasic.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnMachineBasic.OnHoverTextColor = System.Drawing.Color.White;
            this.btnMachineBasic.selected = false;
            this.btnMachineBasic.Size = new System.Drawing.Size(230, 44);
            this.btnMachineBasic.TabIndex = 37;
            this.btnMachineBasic.Text = "机台基础信息";
            this.btnMachineBasic.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnMachineBasic.Textcolor = System.Drawing.Color.Black;
            this.btnMachineBasic.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMachineBasic.Click += new System.EventHandler(this.btnMachineBasic_Click);
            // 
            // btnFactoryRecord
            // 
            this.btnFactoryRecord.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnFactoryRecord.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnFactoryRecord.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnFactoryRecord.BorderRadius = 0;
            this.btnFactoryRecord.ButtonText = "车间输入记录";
            this.btnFactoryRecord.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnFactoryRecord, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnFactoryRecord, BunifuAnimatorNS.DecorationType.None);
            this.btnFactoryRecord.DisabledColor = System.Drawing.Color.Gray;
            this.btnFactoryRecord.Iconcolor = System.Drawing.Color.Transparent;
            this.btnFactoryRecord.Iconimage = null;
            this.btnFactoryRecord.Iconimage_right = null;
            this.btnFactoryRecord.Iconimage_right_Selected = null;
            this.btnFactoryRecord.Iconimage_Selected = null;
            this.btnFactoryRecord.IconMarginLeft = 0;
            this.btnFactoryRecord.IconMarginRight = 0;
            this.btnFactoryRecord.IconRightVisible = true;
            this.btnFactoryRecord.IconRightZoom = 0D;
            this.btnFactoryRecord.IconVisible = true;
            this.btnFactoryRecord.IconZoom = 90D;
            this.btnFactoryRecord.IsTab = false;
            this.btnFactoryRecord.Location = new System.Drawing.Point(0, 352);
            this.btnFactoryRecord.Margin = new System.Windows.Forms.Padding(0);
            this.btnFactoryRecord.Name = "btnFactoryRecord";
            this.btnFactoryRecord.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnFactoryRecord.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnFactoryRecord.OnHoverTextColor = System.Drawing.Color.White;
            this.btnFactoryRecord.selected = false;
            this.btnFactoryRecord.Size = new System.Drawing.Size(230, 44);
            this.btnFactoryRecord.TabIndex = 36;
            this.btnFactoryRecord.Text = "车间输入记录";
            this.btnFactoryRecord.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnFactoryRecord.Textcolor = System.Drawing.Color.Black;
            this.btnFactoryRecord.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFactoryRecord.Click += new System.EventHandler(this.btnFactoryRecord_Click);
            // 
            // btnMissionTot
            // 
            this.btnMissionTot.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnMissionTot.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMissionTot.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnMissionTot.BorderRadius = 0;
            this.btnMissionTot.ButtonText = "任务单汇总表";
            this.btnMissionTot.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnMissionTot, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnMissionTot, BunifuAnimatorNS.DecorationType.None);
            this.btnMissionTot.DisabledColor = System.Drawing.Color.Gray;
            this.btnMissionTot.Iconcolor = System.Drawing.Color.Transparent;
            this.btnMissionTot.Iconimage = null;
            this.btnMissionTot.Iconimage_right = null;
            this.btnMissionTot.Iconimage_right_Selected = null;
            this.btnMissionTot.Iconimage_Selected = null;
            this.btnMissionTot.IconMarginLeft = 0;
            this.btnMissionTot.IconMarginRight = 0;
            this.btnMissionTot.IconRightVisible = true;
            this.btnMissionTot.IconRightZoom = 0D;
            this.btnMissionTot.IconVisible = true;
            this.btnMissionTot.IconZoom = 90D;
            this.btnMissionTot.IsTab = false;
            this.btnMissionTot.Location = new System.Drawing.Point(0, 396);
            this.btnMissionTot.Margin = new System.Windows.Forms.Padding(0);
            this.btnMissionTot.Name = "btnMissionTot";
            this.btnMissionTot.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMissionTot.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnMissionTot.OnHoverTextColor = System.Drawing.Color.White;
            this.btnMissionTot.selected = false;
            this.btnMissionTot.Size = new System.Drawing.Size(230, 44);
            this.btnMissionTot.TabIndex = 35;
            this.btnMissionTot.Text = "任务单汇总表";
            this.btnMissionTot.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnMissionTot.Textcolor = System.Drawing.Color.Black;
            this.btnMissionTot.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMissionTot.Click += new System.EventHandler(this.btnMissionTot_Click);
            // 
            // btnMachineDaily
            // 
            this.btnMachineDaily.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnMachineDaily.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMachineDaily.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnMachineDaily.BorderRadius = 0;
            this.btnMachineDaily.ButtonText = "机台日报表";
            this.btnMachineDaily.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnMachineDaily, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnMachineDaily, BunifuAnimatorNS.DecorationType.None);
            this.btnMachineDaily.DisabledColor = System.Drawing.Color.Gray;
            this.btnMachineDaily.Iconcolor = System.Drawing.Color.Transparent;
            this.btnMachineDaily.Iconimage = null;
            this.btnMachineDaily.Iconimage_right = null;
            this.btnMachineDaily.Iconimage_right_Selected = null;
            this.btnMachineDaily.Iconimage_Selected = null;
            this.btnMachineDaily.IconMarginLeft = 0;
            this.btnMachineDaily.IconMarginRight = 0;
            this.btnMachineDaily.IconRightVisible = true;
            this.btnMachineDaily.IconRightZoom = 0D;
            this.btnMachineDaily.IconVisible = true;
            this.btnMachineDaily.IconZoom = 90D;
            this.btnMachineDaily.IsTab = false;
            this.btnMachineDaily.Location = new System.Drawing.Point(0, 440);
            this.btnMachineDaily.Margin = new System.Windows.Forms.Padding(0);
            this.btnMachineDaily.Name = "btnMachineDaily";
            this.btnMachineDaily.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMachineDaily.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnMachineDaily.OnHoverTextColor = System.Drawing.Color.White;
            this.btnMachineDaily.selected = false;
            this.btnMachineDaily.Size = new System.Drawing.Size(230, 44);
            this.btnMachineDaily.TabIndex = 34;
            this.btnMachineDaily.Text = "机台日报表";
            this.btnMachineDaily.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnMachineDaily.Textcolor = System.Drawing.Color.Black;
            this.btnMachineDaily.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMachineDaily.Click += new System.EventHandler(this.btnMachineDaily_Click);
            // 
            // btnOrder
            // 
            this.btnOrder.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnOrder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnOrder.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnOrder.BorderRadius = 0;
            this.btnOrder.ButtonText = "排程报表";
            this.btnOrder.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnOrder, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnOrder, BunifuAnimatorNS.DecorationType.None);
            this.btnOrder.DisabledColor = System.Drawing.Color.Gray;
            this.btnOrder.Iconcolor = System.Drawing.Color.Transparent;
            this.btnOrder.Iconimage = null;
            this.btnOrder.Iconimage_right = null;
            this.btnOrder.Iconimage_right_Selected = null;
            this.btnOrder.Iconimage_Selected = null;
            this.btnOrder.IconMarginLeft = 0;
            this.btnOrder.IconMarginRight = 0;
            this.btnOrder.IconRightVisible = true;
            this.btnOrder.IconRightZoom = 0D;
            this.btnOrder.IconVisible = true;
            this.btnOrder.IconZoom = 90D;
            this.btnOrder.IsTab = false;
            this.btnOrder.Location = new System.Drawing.Point(0, 484);
            this.btnOrder.Margin = new System.Windows.Forms.Padding(0);
            this.btnOrder.Name = "btnOrder";
            this.btnOrder.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnOrder.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnOrder.OnHoverTextColor = System.Drawing.Color.White;
            this.btnOrder.selected = false;
            this.btnOrder.Size = new System.Drawing.Size(230, 44);
            this.btnOrder.TabIndex = 33;
            this.btnOrder.Text = "排程报表";
            this.btnOrder.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnOrder.Textcolor = System.Drawing.Color.Black;
            this.btnOrder.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOrder.Click += new System.EventHandler(this.btnOrder_Click);
            // 
            // btnProduceOrder
            // 
            this.btnProduceOrder.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnProduceOrder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnProduceOrder.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnProduceOrder.BorderRadius = 0;
            this.btnProduceOrder.ButtonText = "生产排程";
            this.btnProduceOrder.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnProduceOrder, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnProduceOrder, BunifuAnimatorNS.DecorationType.None);
            this.btnProduceOrder.DisabledColor = System.Drawing.Color.Gray;
            this.btnProduceOrder.Iconcolor = System.Drawing.Color.Transparent;
            this.btnProduceOrder.Iconimage = null;
            this.btnProduceOrder.Iconimage_right = null;
            this.btnProduceOrder.Iconimage_right_Selected = null;
            this.btnProduceOrder.Iconimage_Selected = null;
            this.btnProduceOrder.IconMarginLeft = 0;
            this.btnProduceOrder.IconMarginRight = 0;
            this.btnProduceOrder.IconRightVisible = true;
            this.btnProduceOrder.IconRightZoom = 0D;
            this.btnProduceOrder.IconVisible = true;
            this.btnProduceOrder.IconZoom = 90D;
            this.btnProduceOrder.IsTab = false;
            this.btnProduceOrder.Location = new System.Drawing.Point(0, 528);
            this.btnProduceOrder.Margin = new System.Windows.Forms.Padding(0);
            this.btnProduceOrder.Name = "btnProduceOrder";
            this.btnProduceOrder.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnProduceOrder.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnProduceOrder.OnHoverTextColor = System.Drawing.Color.White;
            this.btnProduceOrder.selected = false;
            this.btnProduceOrder.Size = new System.Drawing.Size(230, 44);
            this.btnProduceOrder.TabIndex = 32;
            this.btnProduceOrder.Text = "生产排程";
            this.btnProduceOrder.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnProduceOrder.Textcolor = System.Drawing.Color.Black;
            this.btnProduceOrder.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProduceOrder.Click += new System.EventHandler(this.btnProduceOrder_Click);
            // 
            // btnMEfficient
            // 
            this.btnMEfficient.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnMEfficient.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMEfficient.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnMEfficient.BorderRadius = 0;
            this.btnMEfficient.ButtonText = "机台绩效报表";
            this.btnMEfficient.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnMEfficient, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnMEfficient, BunifuAnimatorNS.DecorationType.None);
            this.btnMEfficient.DisabledColor = System.Drawing.Color.Gray;
            this.btnMEfficient.Iconcolor = System.Drawing.Color.Transparent;
            this.btnMEfficient.Iconimage = null;
            this.btnMEfficient.Iconimage_right = null;
            this.btnMEfficient.Iconimage_right_Selected = null;
            this.btnMEfficient.Iconimage_Selected = null;
            this.btnMEfficient.IconMarginLeft = 0;
            this.btnMEfficient.IconMarginRight = 0;
            this.btnMEfficient.IconRightVisible = true;
            this.btnMEfficient.IconRightZoom = 0D;
            this.btnMEfficient.IconVisible = true;
            this.btnMEfficient.IconZoom = 90D;
            this.btnMEfficient.IsTab = false;
            this.btnMEfficient.Location = new System.Drawing.Point(0, 572);
            this.btnMEfficient.Margin = new System.Windows.Forms.Padding(0);
            this.btnMEfficient.Name = "btnMEfficient";
            this.btnMEfficient.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMEfficient.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnMEfficient.OnHoverTextColor = System.Drawing.Color.White;
            this.btnMEfficient.selected = false;
            this.btnMEfficient.Size = new System.Drawing.Size(230, 44);
            this.btnMEfficient.TabIndex = 40;
            this.btnMEfficient.Text = "机台绩效报表";
            this.btnMEfficient.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnMEfficient.Textcolor = System.Drawing.Color.Black;
            this.btnMEfficient.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMEfficient.Click += new System.EventHandler(this.btnMEfficient_Click);
            // 
            // temperatureRP
            // 
            this.temperatureRP.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.temperatureRP.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.temperatureRP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.temperatureRP.BorderRadius = 0;
            this.temperatureRP.ButtonText = "温控报表";
            this.temperatureRP.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.temperatureRP, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.temperatureRP, BunifuAnimatorNS.DecorationType.None);
            this.temperatureRP.DisabledColor = System.Drawing.Color.Gray;
            this.temperatureRP.Iconcolor = System.Drawing.Color.Transparent;
            this.temperatureRP.Iconimage = null;
            this.temperatureRP.Iconimage_right = null;
            this.temperatureRP.Iconimage_right_Selected = null;
            this.temperatureRP.Iconimage_Selected = null;
            this.temperatureRP.IconMarginLeft = 0;
            this.temperatureRP.IconMarginRight = 0;
            this.temperatureRP.IconRightVisible = true;
            this.temperatureRP.IconRightZoom = 0D;
            this.temperatureRP.IconVisible = true;
            this.temperatureRP.IconZoom = 90D;
            this.temperatureRP.IsTab = false;
            this.temperatureRP.Location = new System.Drawing.Point(0, 616);
            this.temperatureRP.Margin = new System.Windows.Forms.Padding(0);
            this.temperatureRP.Name = "temperatureRP";
            this.temperatureRP.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.temperatureRP.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.temperatureRP.OnHoverTextColor = System.Drawing.Color.White;
            this.temperatureRP.selected = false;
            this.temperatureRP.Size = new System.Drawing.Size(230, 44);
            this.temperatureRP.TabIndex = 43;
            this.temperatureRP.Text = "温控报表";
            this.temperatureRP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.temperatureRP.Textcolor = System.Drawing.Color.Black;
            this.temperatureRP.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.temperatureRP.Click += new System.EventHandler(this.temperatureRP_Click);
            // 
            // btnOrderCheck
            // 
            this.btnOrderCheck.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnOrderCheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnOrderCheck.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnOrderCheck.BorderRadius = 0;
            this.btnOrderCheck.ButtonText = "发货通知检核表";
            this.btnOrderCheck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnOrderCheck, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnOrderCheck, BunifuAnimatorNS.DecorationType.None);
            this.btnOrderCheck.DisabledColor = System.Drawing.Color.Gray;
            this.btnOrderCheck.Iconcolor = System.Drawing.Color.Transparent;
            this.btnOrderCheck.Iconimage = null;
            this.btnOrderCheck.Iconimage_right = null;
            this.btnOrderCheck.Iconimage_right_Selected = null;
            this.btnOrderCheck.Iconimage_Selected = null;
            this.btnOrderCheck.IconMarginLeft = 0;
            this.btnOrderCheck.IconMarginRight = 0;
            this.btnOrderCheck.IconRightVisible = true;
            this.btnOrderCheck.IconRightZoom = 0D;
            this.btnOrderCheck.IconVisible = true;
            this.btnOrderCheck.IconZoom = 90D;
            this.btnOrderCheck.IsTab = false;
            this.btnOrderCheck.Location = new System.Drawing.Point(0, 660);
            this.btnOrderCheck.Margin = new System.Windows.Forms.Padding(0);
            this.btnOrderCheck.Name = "btnOrderCheck";
            this.btnOrderCheck.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnOrderCheck.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnOrderCheck.OnHoverTextColor = System.Drawing.Color.White;
            this.btnOrderCheck.selected = false;
            this.btnOrderCheck.Size = new System.Drawing.Size(230, 44);
            this.btnOrderCheck.TabIndex = 41;
            this.btnOrderCheck.Text = "发货通知检核表";
            this.btnOrderCheck.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnOrderCheck.Textcolor = System.Drawing.Color.Black;
            this.btnOrderCheck.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOrderCheck.Click += new System.EventHandler(this.btnOrderCheck_Click);
            // 
            // btnStorageCheck
            // 
            this.btnStorageCheck.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnStorageCheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnStorageCheck.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnStorageCheck.BorderRadius = 0;
            this.btnStorageCheck.ButtonText = "入库检核表";
            this.btnStorageCheck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnStorageCheck, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnStorageCheck, BunifuAnimatorNS.DecorationType.None);
            this.btnStorageCheck.DisabledColor = System.Drawing.Color.Gray;
            this.btnStorageCheck.Iconcolor = System.Drawing.Color.Transparent;
            this.btnStorageCheck.Iconimage = null;
            this.btnStorageCheck.Iconimage_right = null;
            this.btnStorageCheck.Iconimage_right_Selected = null;
            this.btnStorageCheck.Iconimage_Selected = null;
            this.btnStorageCheck.IconMarginLeft = 0;
            this.btnStorageCheck.IconMarginRight = 0;
            this.btnStorageCheck.IconRightVisible = true;
            this.btnStorageCheck.IconRightZoom = 0D;
            this.btnStorageCheck.IconVisible = true;
            this.btnStorageCheck.IconZoom = 90D;
            this.btnStorageCheck.IsTab = false;
            this.btnStorageCheck.Location = new System.Drawing.Point(0, 704);
            this.btnStorageCheck.Margin = new System.Windows.Forms.Padding(0);
            this.btnStorageCheck.Name = "btnStorageCheck";
            this.btnStorageCheck.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnStorageCheck.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnStorageCheck.OnHoverTextColor = System.Drawing.Color.White;
            this.btnStorageCheck.selected = false;
            this.btnStorageCheck.Size = new System.Drawing.Size(230, 44);
            this.btnStorageCheck.TabIndex = 42;
            this.btnStorageCheck.Text = "入库检核表";
            this.btnStorageCheck.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnStorageCheck.Textcolor = System.Drawing.Color.Black;
            this.btnStorageCheck.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStorageCheck.Click += new System.EventHandler(this.btnStorageCheck_Click);
            // 
            // btnWorkFlow
            // 
            this.btnWorkFlow.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnWorkFlow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnWorkFlow.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnWorkFlow.BorderRadius = 0;
            this.btnWorkFlow.ButtonText = "生产工序信息";
            this.btnWorkFlow.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnWorkFlow, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnWorkFlow, BunifuAnimatorNS.DecorationType.None);
            this.btnWorkFlow.DisabledColor = System.Drawing.Color.Gray;
            this.btnWorkFlow.Iconcolor = System.Drawing.Color.Transparent;
            this.btnWorkFlow.Iconimage = null;
            this.btnWorkFlow.Iconimage_right = null;
            this.btnWorkFlow.Iconimage_right_Selected = null;
            this.btnWorkFlow.Iconimage_Selected = null;
            this.btnWorkFlow.IconMarginLeft = 0;
            this.btnWorkFlow.IconMarginRight = 0;
            this.btnWorkFlow.IconRightVisible = true;
            this.btnWorkFlow.IconRightZoom = 0D;
            this.btnWorkFlow.IconVisible = true;
            this.btnWorkFlow.IconZoom = 90D;
            this.btnWorkFlow.IsTab = false;
            this.btnWorkFlow.Location = new System.Drawing.Point(0, 748);
            this.btnWorkFlow.Margin = new System.Windows.Forms.Padding(0);
            this.btnWorkFlow.Name = "btnWorkFlow";
            this.btnWorkFlow.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnWorkFlow.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnWorkFlow.OnHoverTextColor = System.Drawing.Color.White;
            this.btnWorkFlow.selected = false;
            this.btnWorkFlow.Size = new System.Drawing.Size(230, 44);
            this.btnWorkFlow.TabIndex = 44;
            this.btnWorkFlow.Text = "生产工序信息";
            this.btnWorkFlow.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnWorkFlow.Textcolor = System.Drawing.Color.Black;
            this.btnWorkFlow.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWorkFlow.Click += new System.EventHandler(this.btnWorkFlow_Click);
            // 
            // PdNameDefInfo
            // 
            this.PdNameDefInfo.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.PdNameDefInfo.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.PdNameDefInfo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.PdNameDefInfo.BorderRadius = 0;
            this.PdNameDefInfo.ButtonText = "品名与自定义";
            this.PdNameDefInfo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.PdNameDefInfo, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.PdNameDefInfo, BunifuAnimatorNS.DecorationType.None);
            this.PdNameDefInfo.DisabledColor = System.Drawing.Color.Gray;
            this.PdNameDefInfo.Iconcolor = System.Drawing.Color.Transparent;
            this.PdNameDefInfo.Iconimage = null;
            this.PdNameDefInfo.Iconimage_right = null;
            this.PdNameDefInfo.Iconimage_right_Selected = null;
            this.PdNameDefInfo.Iconimage_Selected = null;
            this.PdNameDefInfo.IconMarginLeft = 0;
            this.PdNameDefInfo.IconMarginRight = 0;
            this.PdNameDefInfo.IconRightVisible = true;
            this.PdNameDefInfo.IconRightZoom = 0D;
            this.PdNameDefInfo.IconVisible = true;
            this.PdNameDefInfo.IconZoom = 90D;
            this.PdNameDefInfo.IsTab = false;
            this.PdNameDefInfo.Location = new System.Drawing.Point(0, 792);
            this.PdNameDefInfo.Margin = new System.Windows.Forms.Padding(0);
            this.PdNameDefInfo.Name = "PdNameDefInfo";
            this.PdNameDefInfo.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.PdNameDefInfo.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.PdNameDefInfo.OnHoverTextColor = System.Drawing.Color.White;
            this.PdNameDefInfo.selected = false;
            this.PdNameDefInfo.Size = new System.Drawing.Size(230, 44);
            this.PdNameDefInfo.TabIndex = 46;
            this.PdNameDefInfo.Text = "品名与自定义";
            this.PdNameDefInfo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.PdNameDefInfo.Textcolor = System.Drawing.Color.Black;
            this.PdNameDefInfo.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PdNameDefInfo.Click += new System.EventHandler(this.PdNameDefInfo_Click);
            // 
            // btnMErrorCheck
            // 
            this.btnMErrorCheck.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnMErrorCheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMErrorCheck.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnMErrorCheck.BorderRadius = 0;
            this.btnMErrorCheck.ButtonText = "串料检核表";
            this.btnMErrorCheck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnMErrorCheck, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnMErrorCheck, BunifuAnimatorNS.DecorationType.None);
            this.btnMErrorCheck.DisabledColor = System.Drawing.Color.Gray;
            this.btnMErrorCheck.Iconcolor = System.Drawing.Color.Transparent;
            this.btnMErrorCheck.Iconimage = null;
            this.btnMErrorCheck.Iconimage_right = null;
            this.btnMErrorCheck.Iconimage_right_Selected = null;
            this.btnMErrorCheck.Iconimage_Selected = null;
            this.btnMErrorCheck.IconMarginLeft = 0;
            this.btnMErrorCheck.IconMarginRight = 0;
            this.btnMErrorCheck.IconRightVisible = true;
            this.btnMErrorCheck.IconRightZoom = 0D;
            this.btnMErrorCheck.IconVisible = true;
            this.btnMErrorCheck.IconZoom = 90D;
            this.btnMErrorCheck.IsTab = false;
            this.btnMErrorCheck.Location = new System.Drawing.Point(0, 836);
            this.btnMErrorCheck.Margin = new System.Windows.Forms.Padding(0);
            this.btnMErrorCheck.Name = "btnMErrorCheck";
            this.btnMErrorCheck.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnMErrorCheck.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnMErrorCheck.OnHoverTextColor = System.Drawing.Color.White;
            this.btnMErrorCheck.selected = false;
            this.btnMErrorCheck.Size = new System.Drawing.Size(230, 44);
            this.btnMErrorCheck.TabIndex = 47;
            this.btnMErrorCheck.Text = "串料检核表";
            this.btnMErrorCheck.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnMErrorCheck.Textcolor = System.Drawing.Color.Black;
            this.btnMErrorCheck.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMErrorCheck.Click += new System.EventHandler(this.btnMErrorCheck_Click);
            // 
            // btnBoxDemand
            // 
            this.btnBoxDemand.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnBoxDemand.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnBoxDemand.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBoxDemand.BorderRadius = 0;
            this.btnBoxDemand.ButtonText = "外箱需求料表";
            this.btnBoxDemand.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PanelTransition.SetDecoration(this.btnBoxDemand, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.btnBoxDemand, BunifuAnimatorNS.DecorationType.None);
            this.btnBoxDemand.DisabledColor = System.Drawing.Color.Gray;
            this.btnBoxDemand.Iconcolor = System.Drawing.Color.Transparent;
            this.btnBoxDemand.Iconimage = null;
            this.btnBoxDemand.Iconimage_right = null;
            this.btnBoxDemand.Iconimage_right_Selected = null;
            this.btnBoxDemand.Iconimage_Selected = null;
            this.btnBoxDemand.IconMarginLeft = 0;
            this.btnBoxDemand.IconMarginRight = 0;
            this.btnBoxDemand.IconRightVisible = true;
            this.btnBoxDemand.IconRightZoom = 0D;
            this.btnBoxDemand.IconVisible = true;
            this.btnBoxDemand.IconZoom = 90D;
            this.btnBoxDemand.IsTab = false;
            this.btnBoxDemand.Location = new System.Drawing.Point(0, 880);
            this.btnBoxDemand.Margin = new System.Windows.Forms.Padding(0);
            this.btnBoxDemand.Name = "btnBoxDemand";
            this.btnBoxDemand.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(238)))), ((int)(((byte)(158)))), ((int)(((byte)(159)))));
            this.btnBoxDemand.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.btnBoxDemand.OnHoverTextColor = System.Drawing.Color.White;
            this.btnBoxDemand.selected = false;
            this.btnBoxDemand.Size = new System.Drawing.Size(230, 44);
            this.btnBoxDemand.TabIndex = 48;
            this.btnBoxDemand.Text = "外箱需求料表";
            this.btnBoxDemand.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.btnBoxDemand.Textcolor = System.Drawing.Color.Black;
            this.btnBoxDemand.TextFont = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBoxDemand.Click += new System.EventHandler(this.btnBoxDemand_Click_1);
            // 
            // bunifuTransition1
            // 
            this.bunifuTransition1.AnimationType = BunifuAnimatorNS.AnimationType.Rotate;
            this.bunifuTransition1.Cursor = null;
            animation1.AnimateOnlyDifferences = true;
            animation1.BlindCoeff = ((System.Drawing.PointF)(resources.GetObject("animation1.BlindCoeff")));
            animation1.LeafCoeff = 0F;
            animation1.MaxTime = 1F;
            animation1.MinTime = 0F;
            animation1.MosaicCoeff = ((System.Drawing.PointF)(resources.GetObject("animation1.MosaicCoeff")));
            animation1.MosaicShift = ((System.Drawing.PointF)(resources.GetObject("animation1.MosaicShift")));
            animation1.MosaicSize = 0;
            animation1.Padding = new System.Windows.Forms.Padding(50);
            animation1.RotateCoeff = 1F;
            animation1.RotateLimit = 0F;
            animation1.ScaleCoeff = ((System.Drawing.PointF)(resources.GetObject("animation1.ScaleCoeff")));
            animation1.SlideCoeff = ((System.Drawing.PointF)(resources.GetObject("animation1.SlideCoeff")));
            animation1.TimeCoeff = 0F;
            animation1.TransparencyCoeff = 1F;
            this.bunifuTransition1.DefaultAnimation = animation1;
            // 
            // MainTabControl
            // 
            this.PanelTransition.SetDecoration(this.MainTabControl, BunifuAnimatorNS.DecorationType.None);
            this.bunifuTransition1.SetDecoration(this.MainTabControl, BunifuAnimatorNS.DecorationType.None);
            this.MainTabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.MainTabControl.Location = new System.Drawing.Point(230, 0);
            this.MainTabControl.Margin = new System.Windows.Forms.Padding(4);
            this.MainTabControl.Name = "MainTabControl";
            this.MainTabControl.SelectedIndex = 0;
            this.MainTabControl.Size = new System.Drawing.Size(1672, 1033);
            this.MainTabControl.TabIndex = 4;
            // 
            // PanelTransition
            // 
            this.PanelTransition.AnimationType = BunifuAnimatorNS.AnimationType.Transparent;
            this.PanelTransition.Cursor = null;
            animation2.AnimateOnlyDifferences = true;
            animation2.BlindCoeff = ((System.Drawing.PointF)(resources.GetObject("animation2.BlindCoeff")));
            animation2.LeafCoeff = 0F;
            animation2.MaxTime = 1F;
            animation2.MinTime = 0F;
            animation2.MosaicCoeff = ((System.Drawing.PointF)(resources.GetObject("animation2.MosaicCoeff")));
            animation2.MosaicShift = ((System.Drawing.PointF)(resources.GetObject("animation2.MosaicShift")));
            animation2.MosaicSize = 0;
            animation2.Padding = new System.Windows.Forms.Padding(0);
            animation2.RotateCoeff = 0F;
            animation2.RotateLimit = 0F;
            animation2.ScaleCoeff = ((System.Drawing.PointF)(resources.GetObject("animation2.ScaleCoeff")));
            animation2.SlideCoeff = ((System.Drawing.PointF)(resources.GetObject("animation2.SlideCoeff")));
            animation2.TimeCoeff = 0F;
            animation2.TransparencyCoeff = 1F;
            this.PanelTransition.DefaultAnimation = animation2;
            // 
            // Dashboard
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(162)))), ((int)(((byte)(201)))), ((int)(((byte)(198)))));
            this.ClientSize = new System.Drawing.Size(1902, 1033);
            this.Controls.Add(this.MainTabControl);
            this.Controls.Add(this.panel1);
            this.bunifuTransition1.SetDecoration(this, BunifuAnimatorNS.DecorationType.None);
            this.PanelTransition.SetDecoration(this, BunifuAnimatorNS.DecorationType.None);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Dashboard";
            this.Text = "宁波诚毅纸业有限公司";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Dashboard_FormClosing);
            this.Load += new System.EventHandler(this.Dashboard_Load);
            this.panel1.ResumeLayout(false);
            this.bunifuGradientPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private Bunifu.Framework.UI.BunifuGradientPanel bunifuGradientPanel1;
        private BunifuAnimatorNS.BunifuTransition bunifuTransition1;
        private BunifuAnimatorNS.BunifuTransition PanelTransition;
        private System.Windows.Forms.TabControl MainTabControl;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private Bunifu.Framework.UI.BunifuFlatButton btnBoxLabel;
        private Bunifu.Framework.UI.BunifuFlatButton btnKnifeModel;
        private Bunifu.Framework.UI.BunifuFlatButton btnAuth;
        private Bunifu.Framework.UI.BunifuFlatButton btnMachineBasic;
        private Bunifu.Framework.UI.BunifuFlatButton btnFactoryRecord;
        private Bunifu.Framework.UI.BunifuFlatButton btnMissionTot;
        private Bunifu.Framework.UI.BunifuFlatButton btnMachineDaily;
        private Bunifu.Framework.UI.BunifuFlatButton btnOrder;
        private Bunifu.Framework.UI.BunifuFlatButton btnProduceOrder;
        private Bunifu.Framework.UI.BunifuFlatButton btnMEfficient;
        private Bunifu.Framework.UI.BunifuFlatButton temperatureRP;
        private Bunifu.Framework.UI.BunifuFlatButton btnOrderCheck;
        private Bunifu.Framework.UI.BunifuFlatButton btnStorageCheck;
        private Bunifu.Framework.UI.BunifuFlatButton btnWorkFlow;
        private Bunifu.Framework.UI.BunifuFlatButton PdNameDefInfo;
        private Bunifu.Framework.UI.BunifuFlatButton btnMErrorCheck;
        private Bunifu.Framework.UI.BunifuFlatButton btnBoxDemand;
    }
}