namespace JMC
{
    partial class frmComSelect
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmComSelect));
            this.dg1 = new System.Windows.Forms.DataGridView();
            this.btnOK = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.btnSel = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).BeginInit();
            this.SuspendLayout();
            // 
            // dg1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Gainsboro;
            this.dg1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dg1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dg1.Location = new System.Drawing.Point(7, 6);
            this.dg1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dg1.MultiSelect = false;
            this.dg1.Name = "dg1";
            this.dg1.ReadOnly = true;
            this.dg1.RowTemplate.Height = 21;
            this.dg1.Size = new System.Drawing.Size(759, 215);
            this.dg1.TabIndex = 1;
            this.dg1.TabStop = false;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(673, 298);
            this.btnOK.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(93, 30);
            this.btnOK.TabIndex = 0;
            this.btnOK.Text = "OK(&N)";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(3, 238);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(370, 23);
            this.label1.TabIndex = 2;
            this.label1.Text = "【暫定機能】※社員名簿ファイルを選択してください";
            // 
            // txtFileName
            // 
            this.txtFileName.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtFileName.Location = new System.Drawing.Point(7, 259);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(694, 29);
            this.txtFileName.TabIndex = 3;
            // 
            // btnSel
            // 
            this.btnSel.Font = new System.Drawing.Font("メイリオ", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnSel.Location = new System.Drawing.Point(701, 259);
            this.btnSel.Name = "btnSel";
            this.btnSel.Size = new System.Drawing.Size(65, 28);
            this.btnSel.TabIndex = 4;
            this.btnSel.Text = "参照...";
            this.btnSel.UseVisualStyleBackColor = true;
            this.btnSel.Click += new System.EventHandler(this.btnSel_Click);
            // 
            // frmComSelect
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(779, 337);
            this.Controls.Add(this.btnSel);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.dg1);
            this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "frmComSelect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "社員ファイル選択";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmComSelect_FormClosing);
            this.Load += new System.EventHandler(this.frmComSelect_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dg1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button btnSel;
    }
}

