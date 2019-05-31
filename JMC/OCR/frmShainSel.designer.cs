namespace JMC.OCR
{
    partial class frmShainSel
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmShainSel));
            this.dg1 = new System.Windows.Forms.DataGridView();
            this.btnSel = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).BeginInit();
            this.SuspendLayout();
            // 
            // dg1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Gainsboro;
            this.dg1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dg1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dg1.Location = new System.Drawing.Point(13, 13);
            this.dg1.Margin = new System.Windows.Forms.Padding(4);
            this.dg1.Name = "dg1";
            this.dg1.ReadOnly = true;
            this.dg1.RowTemplate.Height = 21;
            this.dg1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dg1.Size = new System.Drawing.Size(770, 583);
            this.dg1.TabIndex = 13;
            this.dg1.TabStop = false;
            this.dg1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dg1_CellDoubleClick);
            // 
            // btnSel
            // 
            this.btnSel.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnSel.Image = global::JMC.Properties.Resources.lens;
            this.btnSel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSel.Location = new System.Drawing.Point(611, 613);
            this.btnSel.Name = "btnSel";
            this.btnSel.Size = new System.Drawing.Size(83, 32);
            this.btnSel.TabIndex = 14;
            this.btnSel.Text = "選択(&S)";
            this.btnSel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.btnSel, "編集する勤務記録表を選択します");
            this.btnSel.UseVisualStyleBackColor = true;
            this.btnSel.Click += new System.EventHandler(this.btnSel_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnRtn.Image = global::JMC.Properties.Resources.undo__1_;
            this.btnRtn.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnRtn.Location = new System.Drawing.Point(700, 613);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(83, 32);
            this.btnRtn.TabIndex = 15;
            this.btnRtn.Text = "戻る(&E)";
            this.btnRtn.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.toolTip1.SetToolTip(this.btnRtn, "データ登録画面に戻ります");
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // frmShainSel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(796, 657);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnSel);
            this.Controls.Add(this.dg1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmShainSel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "出勤簿データ選択";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmShainSel_FormClosing);
            this.Load += new System.EventHandler(this.frmShainSel_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dg1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dg1;
        private System.Windows.Forms.Button btnSel;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}