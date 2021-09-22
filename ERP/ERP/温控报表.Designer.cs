namespace Combination
{
    partial class 温控报表
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(温控报表));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnSeek = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.btnSetting = new Bunifu.Framework.UI.BunifuImageButton();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnSetting)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 11;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 4.216152F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 7.185273F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.85511F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.090909F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.66746F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 3.503563F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 14.72684F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 3.44418F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.090909F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.090909F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 9.090909F));
            this.tableLayoutPanel1.Controls.Add(this.btnExport, 9, 1);
            this.tableLayoutPanel1.Controls.Add(this.btnSeek, 8, 1);
            this.tableLayoutPanel1.Controls.Add(this.label2, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.dateTimePicker1, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.btnSetting, 10, 0);
            this.tableLayoutPanel1.Controls.Add(this.dataGridView1, 0, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 44F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 76F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1684, 994);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // btnExport
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.btnExport, 2);
            this.btnExport.Font = new System.Drawing.Font("微軟正黑體", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnExport.Location = new System.Drawing.Point(1380, 53);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(120, 54);
            this.btnExport.TabIndex = 90;
            this.btnExport.Text = "导出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnSeek
            // 
            this.btnSeek.Font = new System.Drawing.Font("微軟正黑體", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnSeek.Location = new System.Drawing.Point(1227, 53);
            this.btnSeek.Name = "btnSeek";
            this.btnSeek.Size = new System.Drawing.Size(120, 54);
            this.btnSeek.TabIndex = 89;
            this.btnSeek.Text = "查询";
            this.btnSeek.UseVisualStyleBackColor = true;
            this.btnSeek.Click += new System.EventHandler(this.btnSeek_Click);
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(79, 64);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 31);
            this.label2.TabIndex = 2;
            this.label2.Text = "检测日期";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.dateTimePicker1.CustomFormat = "";
            this.dateTimePicker1.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.dateTimePicker1.Location = new System.Drawing.Point(195, 60);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(227, 39);
            this.dateTimePicker1.TabIndex = 3;
            // 
            // btnSetting
            // 
            this.btnSetting.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnSetting.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(162)))), ((int)(((byte)(201)))), ((int)(((byte)(198)))));
            this.btnSetting.Image = ((System.Drawing.Image)(resources.GetObject("btnSetting.Image")));
            this.btnSetting.ImageActive = null;
            this.btnSetting.Location = new System.Drawing.Point(1583, 3);
            this.btnSetting.Name = "btnSetting";
            this.btnSetting.Size = new System.Drawing.Size(48, 44);
            this.btnSetting.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.btnSetting.TabIndex = 6;
            this.btnSetting.TabStop = false;
            this.btnSetting.Zoom = 10;
            this.btnSetting.Click += new System.EventHandler(this.btnSetting_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.tableLayoutPanel1.SetColumnSpan(this.dataGridView1, 11);
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(3, 157);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(1678, 834);
            this.dataGridView1.TabIndex = 91;
            // 
            // 温控报表
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(162)))), ((int)(((byte)(201)))), ((int)(((byte)(198)))));
            this.ClientSize = new System.Drawing.Size(1684, 994);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "温控报表";
            this.Text = "温控报表";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.btnSetting)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Button btnSeek;
        private System.Windows.Forms.Button btnExport;
        private Bunifu.Framework.UI.BunifuImageButton btnSetting;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}