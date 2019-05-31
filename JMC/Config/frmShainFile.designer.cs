namespace JMC.Config
{
    partial class frmShainFile
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmShainFile));
            this.dg = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.linkLabel4 = new System.Windows.Forms.LinkLabel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.linkLabel3 = new System.Windows.Forms.LinkLabel();
            this.button5 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.txtGrpName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dg)).BeginInit();
            this.SuspendLayout();
            // 
            // dg
            // 
            this.dg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dg.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dg.Location = new System.Drawing.Point(12, 12);
            this.dg.Name = "dg";
            this.dg.RowTemplate.Height = 21;
            this.dg.Size = new System.Drawing.Size(831, 301);
            this.dg.TabIndex = 9;
            this.dg.TabStop = false;
            this.dg.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dg_CellDoubleClick);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.SteelBlue;
            this.label3.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.ForeColor = System.Drawing.Color.White;
            this.label3.Location = new System.Drawing.Point(12, 388);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(131, 29);
            this.label3.TabIndex = 46;
            this.label3.Text = "ＣＳＶファイル名";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.Color.SteelBlue;
            this.label6.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.ForeColor = System.Drawing.Color.White;
            this.label6.Location = new System.Drawing.Point(12, 354);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(131, 31);
            this.label6.TabIndex = 51;
            this.label6.Text = "雇用種別";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtFileName
            // 
            this.txtFileName.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFileName.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtFileName.Location = new System.Drawing.Point(142, 388);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(637, 29);
            this.txtFileName.TabIndex = 2;
            // 
            // linkLabel4
            // 
            this.linkLabel4.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkLabel4.Image = ((System.Drawing.Image)(resources.GetObject("linkLabel4.Image")));
            this.linkLabel4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel4.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel4.Location = new System.Drawing.Point(481, 439);
            this.linkLabel4.Name = "linkLabel4";
            this.linkLabel4.Size = new System.Drawing.Size(92, 24);
            this.linkLabel4.TabIndex = 5;
            this.linkLabel4.TabStop = true;
            this.linkLabel4.Text = "更新する";
            this.linkLabel4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkLabel4.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel4_LinkClicked);
            // 
            // linkLabel2
            // 
            this.linkLabel2.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkLabel2.Image = ((System.Drawing.Image)(resources.GetObject("linkLabel2.Image")));
            this.linkLabel2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel2.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel2.Location = new System.Drawing.Point(582, 439);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(92, 24);
            this.linkLabel2.TabIndex = 6;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "削除する";
            this.linkLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // linkLabel1
            // 
            this.linkLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabel1.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkLabel1.Image = ((System.Drawing.Image)(resources.GetObject("linkLabel1.Image")));
            this.linkLabel1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel1.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel1.Location = new System.Drawing.Point(781, 439);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(62, 24);
            this.linkLabel1.TabIndex = 8;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "終了";
            this.linkLabel1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // linkLabel3
            // 
            this.linkLabel3.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.linkLabel3.Image = ((System.Drawing.Image)(resources.GetObject("linkLabel3.Image")));
            this.linkLabel3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.linkLabel3.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.linkLabel3.Location = new System.Drawing.Point(683, 439);
            this.linkLabel3.Name = "linkLabel3";
            this.linkLabel3.Size = new System.Drawing.Size(92, 24);
            this.linkLabel3.TabIndex = 7;
            this.linkLabel3.TabStop = true;
            this.linkLabel3.Text = "取り消し";
            this.linkLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.linkLabel3.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel3_LinkClicked);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(779, 388);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(64, 29);
            this.button5.TabIndex = 3;
            this.button5.Text = "参照...";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "社員",
            "パートタイマー"});
            this.comboBox1.Location = new System.Drawing.Point(142, 354);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(182, 31);
            this.comboBox1.TabIndex = 1;
            // 
            // txtGrpName
            // 
            this.txtGrpName.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtGrpName.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.txtGrpName.Location = new System.Drawing.Point(142, 322);
            this.txtGrpName.MaxLength = 100;
            this.txtGrpName.Name = "txtGrpName";
            this.txtGrpName.Size = new System.Drawing.Size(637, 29);
            this.txtGrpName.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.SteelBlue;
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(13, 322);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(131, 29);
            this.label1.TabIndex = 54;
            this.label1.Text = "グループ名";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmShainFile
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(855, 474);
            this.Controls.Add(this.txtGrpName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.linkLabel3);
            this.Controls.Add(this.linkLabel4);
            this.Controls.Add(this.linkLabel2);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.dg);
            this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "frmShainFile";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "給与グループ登録";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frm_FormClosing);
            this.Load += new System.EventHandler(this.frm_Load);
            this.Shown += new System.EventHandler(this.frmKintaiKbn_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dg)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dg;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.LinkLabel linkLabel4;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel linkLabel3;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TextBox txtGrpName;
        private System.Windows.Forms.Label label1;
    }
}