namespace Combination
{
    partial class Login
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
            BunifuAnimatorNS.Animation animation1 = new BunifuAnimatorNS.Animation();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Login));
            this.bunifuElipse1 = new Bunifu.Framework.UI.BunifuElipse(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.bunifuImageButton1 = new Bunifu.Framework.UI.BunifuImageButton();
            this.bunifuMaterialTextbox2 = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.bunifuMaterialTextbox1 = new Bunifu.Framework.UI.BunifuMaterialTextbox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.bunifuFlatButton1 = new Bunifu.Framework.UI.BunifuFlatButton();
            this.bunifuTransition1 = new BunifuAnimatorNS.BunifuTransition(this.components);
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bunifuImageButton1)).BeginInit();
            this.SuspendLayout();
            // 
            // bunifuElipse1
            // 
            this.bunifuElipse1.ElipseRadius = 30;
            this.bunifuElipse1.TargetControl = this;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(222)))), ((int)(((byte)(95)))), ((int)(((byte)(95)))));
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.pictureBox1);
            this.bunifuTransition1.SetDecoration(this.panel1, BunifuAnimatorNS.DecorationType.None);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(248, 419);
            this.panel1.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.bunifuTransition1.SetDecoration(this.label4, BunifuAnimatorNS.DecorationType.None);
            this.label4.ForeColor = System.Drawing.Color.White;
            this.label4.Location = new System.Drawing.Point(12, 398);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 12);
            this.label4.TabIndex = 9;
            this.label4.Text = "Version 1. 4. 3";
            // 
            // pictureBox1
            // 
            this.bunifuTransition1.SetDecoration(this.pictureBox1, BunifuAnimatorNS.DecorationType.None);
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(32, 106);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(185, 146);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.SystemColors.Window;
            this.panel2.Controls.Add(this.textBox1);
            this.panel2.Controls.Add(this.bunifuImageButton1);
            this.panel2.Controls.Add(this.bunifuMaterialTextbox2);
            this.panel2.Controls.Add(this.bunifuMaterialTextbox1);
            this.panel2.Controls.Add(this.label3);
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.bunifuFlatButton1);
            this.bunifuTransition1.SetDecoration(this.panel2, BunifuAnimatorNS.DecorationType.None);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(570, 419);
            this.panel2.TabIndex = 1;
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.bunifuTransition1.SetDecoration(this.textBox1, BunifuAnimatorNS.DecorationType.None);
            this.textBox1.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox1.Location = new System.Drawing.Point(285, 243);
            this.textBox1.Name = "textBox1";
            this.textBox1.PasswordChar = '*';
            this.textBox1.Size = new System.Drawing.Size(206, 24);
            this.textBox1.TabIndex = 2;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
            // 
            // bunifuImageButton1
            // 
            this.bunifuImageButton1.BackColor = System.Drawing.SystemColors.Window;
            this.bunifuTransition1.SetDecoration(this.bunifuImageButton1, BunifuAnimatorNS.DecorationType.None);
            this.bunifuImageButton1.Image = ((System.Drawing.Image)(resources.GetObject("bunifuImageButton1.Image")));
            this.bunifuImageButton1.ImageActive = null;
            this.bunifuImageButton1.Location = new System.Drawing.Point(519, 3);
            this.bunifuImageButton1.Name = "bunifuImageButton1";
            this.bunifuImageButton1.Size = new System.Drawing.Size(39, 41);
            this.bunifuImageButton1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.bunifuImageButton1.TabIndex = 8;
            this.bunifuImageButton1.TabStop = false;
            this.bunifuImageButton1.Zoom = 10;
            this.bunifuImageButton1.Click += new System.EventHandler(this.bunifuImageButton1_Click);
            // 
            // bunifuMaterialTextbox2
            // 
            this.bunifuMaterialTextbox2.BackColor = System.Drawing.SystemColors.Window;
            this.bunifuMaterialTextbox2.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.bunifuTransition1.SetDecoration(this.bunifuMaterialTextbox2, BunifuAnimatorNS.DecorationType.None);
            this.bunifuMaterialTextbox2.Enabled = false;
            this.bunifuMaterialTextbox2.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bunifuMaterialTextbox2.ForeColor = System.Drawing.Color.Black;
            this.bunifuMaterialTextbox2.HintForeColor = System.Drawing.Color.Empty;
            this.bunifuMaterialTextbox2.HintText = "";
            this.bunifuMaterialTextbox2.isPassword = false;
            this.bunifuMaterialTextbox2.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(212)))), ((int)(((byte)(89)))), ((int)(((byte)(89)))));
            this.bunifuMaterialTextbox2.LineIdleColor = System.Drawing.Color.FromArgb(((int)(((byte)(212)))), ((int)(((byte)(89)))), ((int)(((byte)(89)))));
            this.bunifuMaterialTextbox2.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.bunifuMaterialTextbox2.LineThickness = 3;
            this.bunifuMaterialTextbox2.Location = new System.Drawing.Point(285, 241);
            this.bunifuMaterialTextbox2.Margin = new System.Windows.Forms.Padding(4);
            this.bunifuMaterialTextbox2.Name = "bunifuMaterialTextbox2";
            this.bunifuMaterialTextbox2.Size = new System.Drawing.Size(206, 31);
            this.bunifuMaterialTextbox2.TabIndex = 7;
            this.bunifuMaterialTextbox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            // 
            // bunifuMaterialTextbox1
            // 
            this.bunifuMaterialTextbox1.BackColor = System.Drawing.SystemColors.Window;
            this.bunifuMaterialTextbox1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.bunifuTransition1.SetDecoration(this.bunifuMaterialTextbox1, BunifuAnimatorNS.DecorationType.None);
            this.bunifuMaterialTextbox1.Font = new System.Drawing.Font("Century Gothic", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bunifuMaterialTextbox1.ForeColor = System.Drawing.Color.Black;
            this.bunifuMaterialTextbox1.HintForeColor = System.Drawing.Color.Empty;
            this.bunifuMaterialTextbox1.HintText = "";
            this.bunifuMaterialTextbox1.isPassword = false;
            this.bunifuMaterialTextbox1.LineFocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(212)))), ((int)(((byte)(89)))), ((int)(((byte)(89)))));
            this.bunifuMaterialTextbox1.LineIdleColor = System.Drawing.Color.FromArgb(((int)(((byte)(212)))), ((int)(((byte)(89)))), ((int)(((byte)(89)))));
            this.bunifuMaterialTextbox1.LineMouseHoverColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.bunifuMaterialTextbox1.LineThickness = 3;
            this.bunifuMaterialTextbox1.Location = new System.Drawing.Point(285, 155);
            this.bunifuMaterialTextbox1.Margin = new System.Windows.Forms.Padding(4);
            this.bunifuMaterialTextbox1.Name = "bunifuMaterialTextbox1";
            this.bunifuMaterialTextbox1.Size = new System.Drawing.Size(206, 31);
            this.bunifuMaterialTextbox1.TabIndex = 1;
            this.bunifuMaterialTextbox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.bunifuMaterialTextbox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.bunifuMaterialTextbox1_KeyDown);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.bunifuTransition1.SetDecoration(this.label3, BunifuAnimatorNS.DecorationType.None);
            this.label3.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(212)))), ((int)(((byte)(89)))), ((int)(((byte)(89)))));
            this.label3.Location = new System.Drawing.Point(280, 211);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 27);
            this.label3.TabIndex = 3;
            this.label3.Text = "密码：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.bunifuTransition1.SetDecoration(this.label2, BunifuAnimatorNS.DecorationType.None);
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(212)))), ((int)(((byte)(89)))), ((int)(((byte)(89)))));
            this.label2.Location = new System.Drawing.Point(280, 125);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(96, 27);
            this.label2.TabIndex = 2;
            this.label2.Text = "用户名：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.bunifuTransition1.SetDecoration(this.label1, BunifuAnimatorNS.DecorationType.None);
            this.label1.Font = new System.Drawing.Font("微軟正黑體", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(212)))), ((int)(((byte)(89)))), ((int)(((byte)(89)))));
            this.label1.Location = new System.Drawing.Point(279, 59);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(123, 35);
            this.label1.TabIndex = 2;
            this.label1.Text = "用户登录";
            // 
            // bunifuFlatButton1
            // 
            this.bunifuFlatButton1.Activecolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.bunifuFlatButton1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(233)))), ((int)(((byte)(95)))), ((int)(((byte)(95)))));
            this.bunifuFlatButton1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.bunifuFlatButton1.BorderRadius = 7;
            this.bunifuFlatButton1.ButtonText = "用户登入";
            this.bunifuFlatButton1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bunifuTransition1.SetDecoration(this.bunifuFlatButton1, BunifuAnimatorNS.DecorationType.None);
            this.bunifuFlatButton1.DisabledColor = System.Drawing.Color.Gray;
            this.bunifuFlatButton1.Iconcolor = System.Drawing.Color.Transparent;
            this.bunifuFlatButton1.Iconimage = null;
            this.bunifuFlatButton1.Iconimage_right = null;
            this.bunifuFlatButton1.Iconimage_right_Selected = null;
            this.bunifuFlatButton1.Iconimage_Selected = null;
            this.bunifuFlatButton1.IconMarginLeft = 0;
            this.bunifuFlatButton1.IconMarginRight = 0;
            this.bunifuFlatButton1.IconRightVisible = false;
            this.bunifuFlatButton1.IconRightZoom = 0D;
            this.bunifuFlatButton1.IconVisible = true;
            this.bunifuFlatButton1.IconZoom = 90D;
            this.bunifuFlatButton1.IsTab = false;
            this.bunifuFlatButton1.Location = new System.Drawing.Point(339, 336);
            this.bunifuFlatButton1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.bunifuFlatButton1.Name = "bunifuFlatButton1";
            this.bunifuFlatButton1.Normalcolor = System.Drawing.Color.FromArgb(((int)(((byte)(233)))), ((int)(((byte)(95)))), ((int)(((byte)(95)))));
            this.bunifuFlatButton1.OnHovercolor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.bunifuFlatButton1.OnHoverTextColor = System.Drawing.Color.White;
            this.bunifuFlatButton1.selected = false;
            this.bunifuFlatButton1.Size = new System.Drawing.Size(152, 43);
            this.bunifuFlatButton1.TabIndex = 3;
            this.bunifuFlatButton1.Text = "用户登入";
            this.bunifuFlatButton1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.bunifuFlatButton1.Textcolor = System.Drawing.Color.White;
            this.bunifuFlatButton1.TextFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bunifuFlatButton1.Click += new System.EventHandler(this.bunifuFlatButton1_Click);
            // 
            // bunifuTransition1
            // 
            this.bunifuTransition1.AnimationType = BunifuAnimatorNS.AnimationType.Particles;
            this.bunifuTransition1.Cursor = null;
            animation1.AnimateOnlyDifferences = true;
            animation1.BlindCoeff = ((System.Drawing.PointF)(resources.GetObject("animation1.BlindCoeff")));
            animation1.LeafCoeff = 0F;
            animation1.MaxTime = 1F;
            animation1.MinTime = 0F;
            animation1.MosaicCoeff = ((System.Drawing.PointF)(resources.GetObject("animation1.MosaicCoeff")));
            animation1.MosaicShift = ((System.Drawing.PointF)(resources.GetObject("animation1.MosaicShift")));
            animation1.MosaicSize = 1;
            animation1.Padding = new System.Windows.Forms.Padding(100, 50, 100, 150);
            animation1.RotateCoeff = 0F;
            animation1.RotateLimit = 0F;
            animation1.ScaleCoeff = ((System.Drawing.PointF)(resources.GetObject("animation1.ScaleCoeff")));
            animation1.SlideCoeff = ((System.Drawing.PointF)(resources.GetObject("animation1.SlideCoeff")));
            animation1.TimeCoeff = 2F;
            animation1.TransparencyCoeff = 0F;
            this.bunifuTransition1.DefaultAnimation = animation1;
            this.bunifuTransition1.Interval = 20;
            // 
            // Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(570, 419);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel2);
            this.bunifuTransition1.SetDecoration(this, BunifuAnimatorNS.DecorationType.None);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Login";
            this.Text = "Login";
            this.Load += new System.EventHandler(this.Login_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bunifuImageButton1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private Bunifu.Framework.UI.BunifuElipse bunifuElipse1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private Bunifu.Framework.UI.BunifuFlatButton bunifuFlatButton1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private Bunifu.Framework.UI.BunifuMaterialTextbox bunifuMaterialTextbox1;
        private BunifuAnimatorNS.BunifuTransition bunifuTransition1;
        private Bunifu.Framework.UI.BunifuMaterialTextbox bunifuMaterialTextbox2;
        private Bunifu.Framework.UI.BunifuImageButton bunifuImageButton1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label4;
    }
}